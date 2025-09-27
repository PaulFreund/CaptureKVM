#include "AudioPlayback.hpp"

#include <algorithm>
#include <fstream>

namespace
{
    using Microsoft::WRL::ComPtr;

    void logAudio(const std::string& message)
    {
        std::ofstream("pckvm.log", std::ios::app) << message << '\n';
    }
}

AudioPlayback::AudioPlayback() = default;
AudioPlayback::~AudioPlayback()
{
    stop();
}

std::wstring AudioPlayback::widen(const std::string& text)
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

std::string AudioPlayback::narrow(const std::wstring& text)
{
    if (text.empty())
    {
        return {};
    }
    const int required = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }
    std::string result(static_cast<std::size_t>(required), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    return result;
}

void AudioPlayback::start(const std::string& deviceMoniker)
{
    stop();

    std::lock_guard<std::mutex> lock(mutex_);

    requestedMoniker_ = widen(deviceMoniker);

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(hr) || hr == S_FALSE)
    {
        coInitialized_ = true;
    }
    else if (hr == RPC_E_CHANGED_MODE)
    {
        coInitialized_ = false;
        logAudio("[Audio] CoInitializeEx: COM already initialised with different model; continuing");
    }
    else
    {
        coInitialized_ = false;
        logAudio("[Audio] CoInitializeEx failed");
    }

    if (!selectDevice(requestedMoniker_))
    {
        releaseGraph();
        if (coInitialized_)
        {
            CoUninitialize();
            coInitialized_ = false;
        }
        return;
    }

    if (!buildGraph())
    {
        releaseGraph();
        if (coInitialized_)
        {
            CoUninitialize();
            coInitialized_ = false;
        }
        return;
    }

    if (control_ && SUCCEEDED(control_->Run()))
    {
        running_ = true;
        logAudio("[Audio] Audio playback started for '" + narrow(selectedFriendlyName_.empty() ? selectedDisplayName_ : selectedFriendlyName_) + "'");
    }
    else
    {
        logAudio("[Audio] Failed to start audio graph");
        releaseGraph();
        if (coInitialized_)
        {
            CoUninitialize();
            coInitialized_ = false;
        }
    }
}

void AudioPlayback::stop()
{
    std::lock_guard<std::mutex> lock(mutex_);

    if (running_ && control_)
    {
        control_->Stop();
    }
    running_ = false;

    releaseGraph();

    if (coInitialized_)
    {
        CoUninitialize();
        coInitialized_ = false;
    }
}

std::string AudioPlayback::currentDeviceFriendlyName() const
{
    std::lock_guard<std::mutex> lock(mutex_);
    const std::wstring& name = !selectedFriendlyName_.empty() ? selectedFriendlyName_ : selectedDisplayName_;
    return narrow(name);
}

bool AudioPlayback::selectDevice(const std::wstring& requestedMoniker)
{
    ComPtr<ICreateDevEnum> devEnum;
    HRESULT hr = CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&devEnum));
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to create device enumerator");
        return false;
    }

    ComPtr<IEnumMoniker> enumMoniker;
    hr = devEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, enumMoniker.GetAddressOf(), 0);
    if (hr != S_OK || !enumMoniker)
    {
        logAudio("[Audio] No audio capture devices were found");
        return false;
    }

    ComPtr<IMoniker> matched;
    std::wstring matchedFriendly;
    std::wstring matchedDisplay;

    ComPtr<IMoniker> fallback;
    std::wstring fallbackFriendly;
    std::wstring fallbackDisplay;

    ComPtr<IMoniker> current;
    ULONG fetched = 0;

    while (enumMoniker->Next(1, current.GetAddressOf(), &fetched) == S_OK)
    {
        std::wstring displayName;
        {
            LPOLESTR name = nullptr;
            if (SUCCEEDED(current->GetDisplayName(nullptr, nullptr, &name)) && name)
            {
                displayName.assign(name);
                CoTaskMemFree(name);
            }
        }

        std::wstring friendly;
        ComPtr<IPropertyBag> bag;
        if (SUCCEEDED(current->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&bag))) && bag)
        {
            VARIANT value;
            VariantInit(&value);
            if (SUCCEEDED(bag->Read(L"FriendlyName", &value, nullptr)) && value.vt == VT_BSTR)
            {
                friendly.assign(value.bstrVal, SysStringLen(value.bstrVal));
            }
            VariantClear(&value);
        }

        const bool matches = !requestedMoniker.empty() &&
            ((!displayName.empty() && displayName == requestedMoniker) || (!friendly.empty() && friendly == requestedMoniker));

        if (matches)
        {
            matched = current;
            matchedFriendly = friendly;
            matchedDisplay = displayName;
            break;
        }

        if (!fallback)
        {
            fallback = current;
            fallbackFriendly = friendly;
            fallbackDisplay = displayName;
        }

        current.Reset();
    }

    if (!matched)
    {
        matched = fallback;
        matchedFriendly = fallbackFriendly;
        matchedDisplay = fallbackDisplay;
    }

    if (!matched)
    {
        logAudio("[Audio] Unable to select an audio capture device");
        return false;
    }

    selectedMoniker_ = matched;
    selectedFriendlyName_ = matchedFriendly;
    selectedDisplayName_ = matchedDisplay;
    if (selectedFriendlyName_.empty())
    {
        selectedFriendlyName_ = selectedDisplayName_;
    }

    return true;
}

bool AudioPlayback::buildGraph()
{
    releaseGraph();

    HRESULT hr = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&graph_));
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to create FilterGraph");
        return false;
    }

    hr = CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&builder_));
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to create CaptureGraphBuilder2");
        return false;
    }

    hr = builder_->SetFiltergraph(graph_.Get());
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to set filter graph on builder");
        return false;
    }

    if (!selectedMoniker_)
    {
        logAudio("[Audio] No selected moniker to build graph");
        return false;
    }

    ComPtr<IBaseFilter> filter;
    hr = selectedMoniker_->BindToObject(nullptr, nullptr, IID_PPV_ARGS(&filter));
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to bind capture filter");
        return false;
    }

    const std::wstring filterName = !selectedFriendlyName_.empty() ? selectedFriendlyName_ : std::wstring(L"Audio Capture");
    hr = graph_->AddFilter(filter.Get(), filterName.c_str());
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to add capture filter to graph");
        return false;
    }

    sourceFilter_ = filter;

    hr = builder_->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, sourceFilter_.Get(), nullptr, nullptr);
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to render audio stream");
        return false;
    }

    hr = graph_->QueryInterface(IID_PPV_ARGS(&control_));
    if (FAILED(hr))
    {
        logAudio("[Audio] Failed to acquire IMediaControl");
        return false;
    }

    return true;
}

void AudioPlayback::releaseGraph()
{
    if (control_)
    {
        control_->Stop();
        control_.Reset();
    }
    if (sourceFilter_)
    {
        sourceFilter_.Reset();
    }
    if (builder_)
    {
        builder_.Reset();
    }
    if (graph_)
    {
        graph_.Reset();
    }
    selectedMoniker_.Reset();
    selectedFriendlyName_.clear();
    selectedDisplayName_.clear();
}
