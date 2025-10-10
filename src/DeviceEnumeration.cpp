#include "DeviceEnumeration.hpp"

#include <Windows.h>
#include <SetupAPI.h>
#include <devguid.h>
#include <dshow.h>
#include <mmdeviceapi.h>
#include <functiondiscoverykeys_devpkey.h>
#include <wrl/client.h>
#include <winreg.h>

#include <algorithm>
#include <array>
#include <cmath>
#include <map>
#include <string>
#include <vector>

#pragma comment(lib, "Setupapi.lib")

namespace
{
    using Microsoft::WRL::ComPtr;

    class ScopedCoInit
    {
    public:
        explicit ScopedCoInit(DWORD flags)
        {
            const HRESULT hr = CoInitializeEx(nullptr, flags);
            if (SUCCEEDED(hr))
            {
                shouldUninit_ = true;
            }
            else if (hr == RPC_E_CHANGED_MODE)
            {
                shouldUninit_ = false;
            }
        }

        ~ScopedCoInit()
        {
            if (shouldUninit_)
            {
                CoUninitialize();
            }
        }

    private:
        bool shouldUninit_ = false;
    };

    std::string wideToUtf8(const std::wstring& input)
    {
        if (input.empty())
        {
            return {};
        }
        const int sizeNeeded = WideCharToMultiByte(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), nullptr, 0, nullptr, nullptr);
        if (sizeNeeded <= 0)
        {
            return {};
        }
        std::string result(static_cast<std::size_t>(sizeNeeded), '\0');
        WideCharToMultiByte(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), result.data(), sizeNeeded, nullptr, nullptr);
        return result;
    }

    std::string bstrToUtf8(BSTR value)
    {
        if (!value)
        {
            return {};
        }
        return wideToUtf8(std::wstring(value, SysStringLen(value)));
    }

    std::wstring utf8ToWide(const std::string& input)
    {
        if (input.empty())
        {
            return {};
        }
        const int required = MultiByteToWideChar(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), nullptr, 0);
        if (required <= 0)
        {
            return {};
        }
        std::wstring result(static_cast<std::size_t>(required), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), result.data(), required);
        return result;
    }

    void freeMediaType(AM_MEDIA_TYPE& mt)
    {
        if (mt.cbFormat != 0 && mt.pbFormat)
        {
            CoTaskMemFree(mt.pbFormat);
            mt.cbFormat = 0;
            mt.pbFormat = nullptr;
        }
        if (mt.pUnk)
        {
            mt.pUnk->Release();
            mt.pUnk = nullptr;
        }
    }

    template <typename DeviceInfo>
    std::vector<DeviceInfo> enumerateCategory(REFCLSID category)
    {
        ScopedCoInit coInit(COINIT_APARTMENTTHREADED);

        std::vector<DeviceInfo> devices;

        ComPtr<ICreateDevEnum> devEnum;
        if (FAILED(CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&devEnum))))
        {
            return devices;
        }

        ComPtr<IEnumMoniker> enumMoniker;
        if (devEnum->CreateClassEnumerator(category, enumMoniker.GetAddressOf(), 0) != S_OK || !enumMoniker)
        {
            return devices;
        }

        ComPtr<IMoniker> moniker;
        ULONG fetched = 0;
        while (enumMoniker->Next(1, moniker.GetAddressOf(), &fetched) == S_OK)
        {
            DeviceInfo info;

            LPOLESTR displayName = nullptr;
            if (SUCCEEDED(moniker->GetDisplayName(nullptr, nullptr, &displayName)) && displayName)
            {
                info.monikerDisplayName = wideToUtf8(displayName);
                CoTaskMemFree(displayName);
            }

            ComPtr<IPropertyBag> props;
            if (SUCCEEDED(moniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&props))) && props)
            {
                VARIANT friendly;
                VariantInit(&friendly);
                if (SUCCEEDED(props->Read(L"FriendlyName", &friendly, nullptr)) && friendly.vt == VT_BSTR)
                {
                    info.friendlyName = bstrToUtf8(friendly.bstrVal);
                }
                VariantClear(&friendly);
            }

            devices.push_back(std::move(info));
            moniker.Reset();
        }

        return devices;
    }
}

std::vector<VideoDeviceInfo> enumerateVideoCaptureDevices()
{
    return enumerateCategory<VideoDeviceInfo>(CLSID_VideoInputDeviceCategory);
}

std::vector<AudioCaptureDeviceInfo> enumerateAudioCaptureDevices()
{
    return enumerateCategory<AudioCaptureDeviceInfo>(CLSID_AudioInputDeviceCategory);
}

std::vector<MicrophoneDeviceInfo> enumerateMicrophoneDevices()
{
    ScopedCoInit coInit(COINIT_MULTITHREADED);
    std::vector<MicrophoneDeviceInfo> devices;

    ComPtr<IMMDeviceEnumerator> enumerator;
    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&enumerator))))
    {
        return devices;
    }

    ComPtr<IMMDeviceCollection> collection;
    if (FAILED(enumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &collection)))
    {
        return devices;
    }

    UINT count = 0;
    if (FAILED(collection->GetCount(&count)))
    {
        return devices;
    }

    for (UINT i = 0; i < count; ++i)
    {
        ComPtr<IMMDevice> device;
        if (FAILED(collection->Item(i, &device)))
        {
            continue;
        }

        LPWSTR id = nullptr;
        if (FAILED(device->GetId(&id)))
        {
            continue;
        }

        MicrophoneDeviceInfo info;
        info.endpointId = wideToUtf8(id);
        CoTaskMemFree(id);

        ComPtr<IPropertyStore> props;
        if (SUCCEEDED(device->OpenPropertyStore(STGM_READ, &props)) && props)
        {
            PROPVARIANT value;
            PropVariantInit(&value);
            if (SUCCEEDED(props->GetValue(PKEY_Device_FriendlyName, &value)) && value.vt == VT_LPWSTR)
            {
                info.friendlyName = wideToUtf8(value.pwszVal);
            }
            PropVariantClear(&value);
        }

        devices.push_back(std::move(info));
    }

    return devices;
}

std::vector<SerialPortInfo> enumerateSerialPorts()
{
    std::vector<SerialPortInfo> ports;

    HDEVINFO deviceInfo = SetupDiGetClassDevsW(&GUID_DEVCLASS_PORTS, nullptr, nullptr, DIGCF_PRESENT);
    if (deviceInfo == INVALID_HANDLE_VALUE)
    {
        return ports;
    }

    SP_DEVINFO_DATA deviceData{};
    deviceData.cbSize = sizeof(deviceData);

    for (DWORD index = 0; SetupDiEnumDeviceInfo(deviceInfo, index, &deviceData); ++index)
    {
        SerialPortInfo info;
        std::array<WCHAR, 256> buffer{};

        if (SetupDiGetDeviceRegistryPropertyW(deviceInfo, &deviceData, SPDRP_FRIENDLYNAME, nullptr, reinterpret_cast<PBYTE>(buffer.data()), static_cast<DWORD>(buffer.size() * sizeof(WCHAR)), nullptr))
        {
            info.friendlyName = wideToUtf8(buffer.data());
        }

        if (SetupDiGetDeviceRegistryPropertyW(deviceInfo, &deviceData, SPDRP_DEVICEDESC, nullptr, reinterpret_cast<PBYTE>(buffer.data()), static_cast<DWORD>(buffer.size() * sizeof(WCHAR)), nullptr))
        {
            info.deviceDescription = wideToUtf8(buffer.data());
        }

        HKEY deviceKey = SetupDiOpenDevRegKey(deviceInfo, &deviceData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ);
        if (deviceKey != INVALID_HANDLE_VALUE)
        {
            std::array<WCHAR, 256> portBuffer{};
            DWORD bufferSize = static_cast<DWORD>(portBuffer.size() * sizeof(WCHAR));
            DWORD valueType = 0;
            if (RegQueryValueExW(deviceKey, L"PortName", nullptr, &valueType, reinterpret_cast<LPBYTE>(portBuffer.data()), &bufferSize) == ERROR_SUCCESS && valueType == REG_SZ)
            {
                portBuffer.back() = L'\0';
                info.portName = wideToUtf8(portBuffer.data());
            }
            RegCloseKey(deviceKey);
        }

        DWORD requiredSize = 0;
        SetupDiGetDeviceRegistryPropertyW(deviceInfo,
                                          &deviceData,
                                          SPDRP_HARDWAREID,
                                          nullptr,
                                          nullptr,
                                          0,
                                          &requiredSize);
        if (requiredSize > 0)
        {
            std::vector<WCHAR> hardwareIds((requiredSize / sizeof(WCHAR)) + 1);
            if (SetupDiGetDeviceRegistryPropertyW(deviceInfo,
                                                  &deviceData,
                                                  SPDRP_HARDWAREID,
                                                  nullptr,
                                                  reinterpret_cast<PBYTE>(hardwareIds.data()),
                                                  requiredSize,
                                                  nullptr))
            {
                for (const WCHAR* id = hardwareIds.data(); id && *id; id += wcslen(id) + 1)
                {
                    info.hardwareIds.push_back(wideToUtf8(id));
                }
            }
        }

        if (!info.portName.empty())
        {
            if (info.friendlyName.empty())
            {
                info.friendlyName = info.portName;
            }
            ports.push_back(std::move(info));
        }
    }

    SetupDiDestroyDeviceInfoList(deviceInfo);
    return ports;
}

std::vector<VideoModeInfo> enumerateVideoModes(const std::string& monikerDisplayName)
{
    std::vector<VideoModeInfo> modes;
    if (monikerDisplayName.empty())
    {
        return modes;
    }

    ScopedCoInit coInit(COINIT_MULTITHREADED);

    ComPtr<IGraphBuilder> graph;
    if (FAILED(CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&graph))))
    {
        return modes;
    }

    ComPtr<ICaptureGraphBuilder2> builder;
    if (FAILED(CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&builder))))
    {
        return modes;
    }

    if (FAILED(builder->SetFiltergraph(graph.Get())))
    {
        return modes;
    }

    const std::wstring monikerWide = utf8ToWide(monikerDisplayName);
    if (monikerWide.empty())
    {
        return modes;
    }

    ComPtr<IBindCtx> bindCtx;
    if (FAILED(CreateBindCtx(0, &bindCtx)))
    {
        return modes;
    }

    ULONG eaten = 0;
    ComPtr<IMoniker> moniker;
    if (FAILED(MkParseDisplayName(bindCtx.Get(), monikerWide.c_str(), &eaten, moniker.GetAddressOf())) || !moniker)
    {
        return modes;
    }

    ComPtr<IBaseFilter> captureFilter;
    if (FAILED(moniker->BindToObject(nullptr, nullptr, IID_PPV_ARGS(&captureFilter))) || !captureFilter)
    {
        return modes;
    }

    if (FAILED(graph->AddFilter(captureFilter.Get(), L"Source")))
    {
        return modes;
    }

    ComPtr<IAMStreamConfig> streamConfig;
    HRESULT hr = builder->FindInterface(&PIN_CATEGORY_CAPTURE,
                                        &MEDIATYPE_Video,
                                        captureFilter.Get(),
                                        IID_PPV_ARGS(streamConfig.GetAddressOf()));
    if (FAILED(hr) || !streamConfig)
    {
        hr = builder->FindInterface(&PIN_CATEGORY_PREVIEW,
                                     &MEDIATYPE_Video,
                                     captureFilter.Get(),
                                     IID_PPV_ARGS(streamConfig.GetAddressOf()));
    }

    if (FAILED(hr) || !streamConfig)
    {
        return modes;
    }

    int capabilityCount = 0;
    int capabilitySize = 0;
    if (FAILED(streamConfig->GetNumberOfCapabilities(&capabilityCount, &capabilitySize)) || capabilityCount <= 0 || capabilitySize <= 0)
    {
        return modes;
    }

    std::vector<std::uint8_t> capabilityBuffer(static_cast<std::size_t>(capabilitySize));
    std::map<std::pair<std::uint32_t, std::uint32_t>, double> uniqueModes;

    for (int i = 0; i < capabilityCount; ++i)
    {
        AM_MEDIA_TYPE* mediaType = nullptr;
        if (FAILED(streamConfig->GetStreamCaps(i, &mediaType, capabilityBuffer.data())) || !mediaType)
        {
            continue;
        }

        if (mediaType->formattype == FORMAT_VideoInfo && mediaType->cbFormat >= sizeof(VIDEOINFOHEADER) && mediaType->pbFormat)
        {
            const auto* vih = reinterpret_cast<const VIDEOINFOHEADER*>(mediaType->pbFormat);
            const std::uint32_t width = static_cast<std::uint32_t>(std::abs(vih->bmiHeader.biWidth));
            const std::uint32_t height = static_cast<std::uint32_t>(std::abs(vih->bmiHeader.biHeight));
            double frameRate = 0.0;
            if (vih->AvgTimePerFrame > 0)
            {
                frameRate = 10'000'000.0 / static_cast<double>(vih->AvgTimePerFrame);
            }

            auto key = std::make_pair(width, height);
            auto existing = uniqueModes.find(key);
            if (existing == uniqueModes.end() || frameRate > existing->second)
            {
                uniqueModes[key] = frameRate;
            }
        }

        freeMediaType(*mediaType);
        CoTaskMemFree(mediaType);
    }

    modes.reserve(uniqueModes.size());
    for (const auto& entry : uniqueModes)
    {
        VideoModeInfo mode;
        mode.width = entry.first.first;
        mode.height = entry.first.second;
        mode.frameRate = entry.second;
        modes.push_back(mode);
    }

    std::sort(modes.begin(), modes.end(), [](const VideoModeInfo& a, const VideoModeInfo& b) {
        if (a.width != b.width)
        {
            return a.width > b.width;
        }
        if (a.height != b.height)
        {
            return a.height > b.height;
        }
        return a.frameRate > b.frameRate;
    });

    return modes;
}
