#pragma once

#include <cstdint>
#include <string>
#include <vector>

struct VideoDeviceInfo {
    std::string monikerDisplayName;
    std::string friendlyName;
};

struct AudioCaptureDeviceInfo {
    std::string monikerDisplayName;
    std::string friendlyName;
};

struct MicrophoneDeviceInfo {
    std::string endpointId;
    std::string friendlyName;
};

struct SerialPortInfo {
    std::string portName;
    std::string friendlyName;
    std::string deviceDescription;
    std::vector<std::string> hardwareIds;
};

struct VideoModeInfo {
    std::uint32_t width = 0;
    std::uint32_t height = 0;
    double frameRate = 0.0;
};

std::vector<VideoDeviceInfo> enumerateVideoCaptureDevices();
std::vector<AudioCaptureDeviceInfo> enumerateAudioCaptureDevices();
std::vector<MicrophoneDeviceInfo> enumerateMicrophoneDevices();
std::vector<SerialPortInfo> enumerateSerialPorts();
std::vector<VideoModeInfo> enumerateVideoModes(const std::string& monikerDisplayName);
