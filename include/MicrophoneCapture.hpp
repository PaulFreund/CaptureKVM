#pragma once

#include <Windows.h>
#include <audioclient.h>
#include <wrl/client.h>

#include <atomic>
#include <cstdint>
#include <memory>
#include <mutex>
#include <string>
#include <thread>

class SerialStreamer;

class MicrophoneCapture {
public:
    MicrophoneCapture();
    ~MicrophoneCapture();

    void start(const std::string& endpointId, SerialStreamer& streamer, bool enableAutoGain);
    void stop();

    [[nodiscard]] bool isRunning() const noexcept { return running_.load(std::memory_order_acquire); }

private:
    void captureThread(std::wstring endpointId);
    bool initializeClient(const std::wstring& endpointId);
    void releaseClient();
    void processAvailableAudio();

    SerialStreamer* streamer_ = nullptr;
    std::thread worker_;
    std::atomic<bool> running_{false};
    std::atomic<bool> stopRequested_{false};
    bool autoGainEnabled_ = true;

    HANDLE captureEvent_ = nullptr;
    Microsoft::WRL::ComPtr<IAudioClient> audioClient_;
    Microsoft::WRL::ComPtr<IAudioCaptureClient> captureClient_;
    WAVEFORMATEX* waveFormat_ = nullptr;
    UINT32 bufferFrameCount_ = 0;
    UINT32 bytesPerFrame_ = 0;
    std::mutex clientMutex_;
};
