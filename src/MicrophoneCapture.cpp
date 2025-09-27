#include "MicrophoneCapture.hpp"
#include "SerialStreamer.hpp"

#include <algorithm>
#include <cmath>
#include <cstring>
#include <fstream>
#include <limits>
#include <string>
#include <mmdeviceapi.h>
#include <functiondiscoverykeys_devpkey.h>
#include <ksmedia.h>
#include <vector>

using Microsoft::WRL::ComPtr;

namespace
{
    constexpr std::uint32_t kTargetSampleRate = 48000;

    void logMic(const std::string& message)
    {
        std::ofstream("pckvm.log", std::ios::app) << message << '\n';
    }

    std::wstring widen(const std::string& text)
    {
        if (text.empty())
        {
            return {};
        }
        const int required = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), nullptr, 0);
        if (required <= 0)
        {
            return {};
        }
        std::wstring result(static_cast<std::size_t>(required), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), result.data(), required);
        return result;
    }

    bool isFloatFormat(const WAVEFORMATEX* format)
    {
        if (!format)
        {
            return false;
        }
        if (format->wFormatTag == WAVE_FORMAT_IEEE_FLOAT)
        {
            return true;
        }
        if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE)
        {
            const auto* ext = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(format);
            return IsEqualGUID(ext->SubFormat, KSDATAFORMAT_SUBTYPE_IEEE_FLOAT);
        }
        return false;
    }

    bool isPcm16Format(const WAVEFORMATEX* format)
    {
        if (!format)
        {
            return false;
        }
        if (format->wFormatTag == WAVE_FORMAT_PCM && format->wBitsPerSample == 16)
        {
            return true;
        }
        if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE)
        {
            const auto* ext = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(format);
            return IsEqualGUID(ext->SubFormat, KSDATAFORMAT_SUBTYPE_PCM) && format->wBitsPerSample == 16;
        }
        return false;
    }

    std::vector<std::int16_t> mixDownToMono(const std::int16_t* samples,
                                            std::size_t frameCount,
                                            std::uint32_t channels)
    {
        if (!samples || frameCount == 0)
        {
            return {};
        }
        if (channels <= 1)
        {
            return std::vector<std::int16_t>(samples, samples + frameCount);
        }

        channels = std::max<std::uint32_t>(1, channels);

        std::vector<double> channelEnergy(channels, 0.0);
        for (std::size_t frame = 0; frame < frameCount; ++frame)
        {
            for (std::uint32_t ch = 0; ch < channels; ++ch)
            {
                const std::int16_t sample = samples[frame * channels + ch];
                channelEnergy[ch] += std::abs(static_cast<double>(sample));
            }
        }

        const auto maxIt = std::max_element(channelEnergy.begin(), channelEnergy.end());
        const std::uint32_t dominantChannel = static_cast<std::uint32_t>(std::distance(channelEnergy.begin(), maxIt));

        std::vector<std::int16_t> mono(frameCount);
        for (std::size_t frame = 0; frame < frameCount; ++frame)
        {
            mono[frame] = samples[frame * channels + dominantChannel];
        }
        return mono;
    }

    std::vector<std::int16_t> resampleLinear(const std::vector<std::int16_t>& input,
                                             std::uint32_t srcRate,
                                             std::uint32_t dstRate)
    {
        if (input.empty())
        {
            return {};
        }
        if (srcRate == dstRate || dstRate == 0)
        {
            return input;
        }

        if (srcRate == 0)
        {
            return input;
        }

        const double step = static_cast<double>(srcRate) / static_cast<double>(dstRate);
        if (step <= 0.0)
        {
            return input;
        }

        const std::size_t outSamples = static_cast<std::size_t>(std::ceil(static_cast<double>(input.size()) / step));
        std::vector<std::int16_t> output(outSamples);

        const std::size_t lastIndex = input.size() > 0 ? input.size() - 1 : 0;
        double srcPos = 0.0;
        for (std::size_t i = 0; i < outSamples; ++i)
        {
            const std::size_t idx = static_cast<std::size_t>(srcPos);
            if (idx >= lastIndex)
            {
                output[i] = input[lastIndex];
                output.resize(i + 1);
                break;
            }
            const double frac = srcPos - static_cast<double>(idx);
            const std::int16_t s0 = input[idx];
            const std::int16_t s1 = input[std::min(idx + 1, lastIndex)];
            const double sample = static_cast<double>(s0) + (static_cast<double>(s1) - static_cast<double>(s0)) * frac;
            output[i] = static_cast<std::int16_t>(std::clamp(sample, -32768.0, 32767.0));
            srcPos += step;
        }

        return output;
    }
}

MicrophoneCapture::MicrophoneCapture() = default;
MicrophoneCapture::~MicrophoneCapture()
{
    stop();
}

void MicrophoneCapture::start(const std::string& endpointId, SerialStreamer& streamer, bool enableAutoGain)
{
    stop();
    streamer_ = &streamer;
    autoGainEnabled_ = enableAutoGain;
    stopRequested_.store(false, std::memory_order_release);
    running_.store(true, std::memory_order_release);
    worker_ = std::thread(&MicrophoneCapture::captureThread, this, widen(endpointId));
}

void MicrophoneCapture::stop()
{
    stopRequested_.store(true, std::memory_order_release);
    if (captureEvent_)
    {
        SetEvent(captureEvent_);
    }

    if (worker_.joinable())
    {
        worker_.join();
    }

    running_.store(false, std::memory_order_release);
    stopRequested_.store(false, std::memory_order_release);
}

void MicrophoneCapture::captureThread(std::wstring endpointId)
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    const bool comInitialized = SUCCEEDED(hr) || hr == S_FALSE;

    if (!initializeClient(endpointId))
    {
        running_.store(false, std::memory_order_release);
        if (comInitialized)
        {
            CoUninitialize();
        }
        return;
    }

    if (FAILED(audioClient_->Start()))
    {
        logMic("[Mic] Failed to start audio client");
        releaseClient();
        running_.store(false, std::memory_order_release);
        if (comInitialized)
        {
            CoUninitialize();
        }
        return;
    }

    logMic("[Mic] Capture started");

    while (!stopRequested_.load(std::memory_order_acquire))
    {
        const DWORD waitResult = WaitForSingleObject(captureEvent_, 50);
        if (waitResult == WAIT_OBJECT_0 || waitResult == WAIT_TIMEOUT)
        {
            processAvailableAudio();
        }
        else
        {
            logMic("[Mic] WaitForSingleObject returned error");
            break;
        }
    }

    audioClient_->Stop();
    releaseClient();
    running_.store(false, std::memory_order_release);

    if (comInitialized)
    {
        CoUninitialize();
    }

    logMic("[Mic] Capture stopped");
}

bool MicrophoneCapture::initializeClient(const std::wstring& endpointId)
{
    std::lock_guard<std::mutex> lock(clientMutex_);

    ComPtr<IMMDeviceEnumerator> enumerator;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&enumerator));
    if (FAILED(hr))
    {
        logMic("[Mic] Failed to create IMMDeviceEnumerator");
        return false;
    }

    ComPtr<IMMDevice> device;
    if (!endpointId.empty())
    {
        hr = enumerator->GetDevice(endpointId.c_str(), &device);
        if (FAILED(hr))
        {
            logMic("[Mic] Failed to open requested endpoint; falling back to default");
        }
    }

    if (!device)
    {
        hr = enumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &device);
        if (FAILED(hr))
        {
            logMic("[Mic] Failed to access default capture endpoint");
            return false;
        }
    }

    hr = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(audioClient_.GetAddressOf()));
    if (FAILED(hr))
    {
        logMic("[Mic] Failed to activate IAudioClient");
        return false;
    }

    hr = audioClient_->GetMixFormat(&waveFormat_);
    if (FAILED(hr))
    {
        logMic("[Mic] GetMixFormat failed");
        releaseClient();
        return false;
    }

    bytesPerFrame_ = waveFormat_->nBlockAlign;

    REFERENCE_TIME defaultPeriod = 0;
    REFERENCE_TIME minimumPeriod = 0;
    if (FAILED(audioClient_->GetDevicePeriod(&defaultPeriod, &minimumPeriod)))
    {
        defaultPeriod = 10000000LL / 100; // 10 ms fallback
    }

    hr = audioClient_->Initialize(AUDCLNT_SHAREMODE_SHARED,
                                  AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                                  defaultPeriod,
                                  0,
                                  waveFormat_,
                                  nullptr);
    if (FAILED(hr))
    {
        logMic("[Mic] IAudioClient::Initialize failed");
        releaseClient();
        return false;
    }

    captureEvent_ = CreateEventW(nullptr, FALSE, FALSE, nullptr);
    if (!captureEvent_)
    {
        logMic("[Mic] Failed to create capture event");
        releaseClient();
        return false;
    }

    hr = audioClient_->SetEventHandle(captureEvent_);
    if (FAILED(hr))
    {
        logMic("[Mic] Failed to set event handle");
        releaseClient();
        return false;
    }

    hr = audioClient_->GetService(IID_PPV_ARGS(&captureClient_));
    if (FAILED(hr))
    {
        logMic("[Mic] Failed to access IAudioCaptureClient");
        releaseClient();
        return false;
    }

    hr = audioClient_->GetBufferSize(&bufferFrameCount_);
    if (FAILED(hr))
    {
        bufferFrameCount_ = 0;
    }

    logMic("[Mic] Microphone capture initialized");
    return true;
}

void MicrophoneCapture::releaseClient()
{
    std::lock_guard<std::mutex> lock(clientMutex_);

    if (captureClient_)
    {
        captureClient_.Reset();
    }
    if (audioClient_)
    {
        audioClient_.Reset();
    }
    if (waveFormat_)
    {
        CoTaskMemFree(waveFormat_);
        waveFormat_ = nullptr;
    }
    if (captureEvent_)
    {
        CloseHandle(captureEvent_);
        captureEvent_ = nullptr;
    }
    bufferFrameCount_ = 0;
    bytesPerFrame_ = 0;
}

void MicrophoneCapture::processAvailableAudio()
{
    if (!captureClient_ || !waveFormat_)
    {
        return;
    }

    bool unsupportedLogged = false;

    while (true)
    {
        BYTE* data = nullptr;
        UINT32 frames = 0;
        DWORD flags = 0;
        HRESULT hr = captureClient_->GetBuffer(&data, &frames, &flags, nullptr, nullptr);
        if (hr == AUDCLNT_S_BUFFER_EMPTY)
        {
            break;
        }
        if (FAILED(hr))
        {
            logMic("[Mic] GetBuffer failed");
            break;
        }

        if (frames == 0)
        {
            captureClient_->ReleaseBuffer(0);
            continue;
        }

        const bool silent = (flags & AUDCLNT_BUFFERFLAGS_SILENT) != 0;
        const bool pcm16 = isPcm16Format(waveFormat_);
        const bool float32 = isFloatFormat(waveFormat_);
        const std::uint32_t channels = waveFormat_->nChannels ? waveFormat_->nChannels : 1;
        std::vector<std::int16_t> samples16;

        if (silent)
        {
            samples16.assign(static_cast<std::size_t>(frames) * channels, 0);
        }
        else if (pcm16)
        {
            const std::size_t bytes = static_cast<std::size_t>(frames) * bytesPerFrame_;
            samples16.resize(static_cast<std::size_t>(frames) * channels);
            std::memcpy(samples16.data(), data, bytes);
        }
        else if (float32)
        {
            const std::size_t sampleCount = static_cast<std::size_t>(frames) * channels;
            samples16.resize(sampleCount);
            const float* samples = reinterpret_cast<const float*>(data);
            for (std::size_t i = 0; i < sampleCount; ++i)
            {
                const float value = std::clamp(samples[i], -1.0f, 1.0f);
                samples16[i] = static_cast<std::int16_t>(value * 32767.0f);
            }
        }
        else
        {
            if (!unsupportedLogged)
            {
                logMic("[Mic] Unsupported microphone sample format; dropping audio");
                unsupportedLogged = true;
            }
            captureClient_->ReleaseBuffer(frames);
            continue;
        }

        const std::uint16_t sampleRate = static_cast<std::uint16_t>(waveFormat_->nSamplesPerSec);
        if (sampleRate != kTargetSampleRate)
        {
            static bool warnedRate = false;
            if (!warnedRate)
            {
                logMic("[Mic] Warning: capture sample rate is " + std::to_string(sampleRate) + " Hz; expected 48000 Hz");
                warnedRate = true;
            }
        }

        std::vector<std::int16_t> mono;
        mono = mixDownToMono(samples16.data(), samples16.size() / (channels ? channels : 1), channels);

        std::vector<std::int16_t> finalSamples;
        finalSamples = resampleLinear(mono, sampleRate, kTargetSampleRate);

        if (!finalSamples.empty())
        {
            if (autoGainEnabled_)
            {
                int maxAbs = 0;
                for (std::int16_t sample : finalSamples)
                {
                    const int absVal = std::abs(static_cast<int>(sample));
                    if (absVal > maxAbs)
                    {
                        maxAbs = absVal;
                    }
                }

                if (maxAbs > 0)
                {
                    constexpr double desiredPeak = 24000.0;
                    double gain = desiredPeak / static_cast<double>(maxAbs);
                    gain = std::clamp(gain, 1.0, 4.0);
                    if (gain > 1.0)
                    {
                        for (auto& sample : finalSamples)
                        {
                            const double scaled = static_cast<double>(sample) * gain;
                            const double clamped = std::clamp(scaled, -32768.0, 32767.0);
                            sample = static_cast<std::int16_t>(clamped);
                        }
                    }
                }
            }

            const std::uint8_t* bytes = reinterpret_cast<const std::uint8_t*>(finalSamples.data());
            const std::size_t byteCount = finalSamples.size() * sizeof(std::int16_t);
            streamer_->publishMicrophoneSamples(bytes, byteCount);
        }

        captureClient_->ReleaseBuffer(frames);
    }
}
