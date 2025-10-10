#include "DeviceEnumeration.hpp"

#include <Windows.h>
#include <SetupAPI.h>
#include <devguid.h>
#include <dshow.h>
#include <mmdeviceapi.h>
#include <functiondiscoverykeys_devpkey.h>
#include <wrl/client.h>
#include <winreg.h>

#include <array>
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
