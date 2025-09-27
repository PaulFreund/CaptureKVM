#pragma once

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
};

std::vector<VideoDeviceInfo> enumerateVideoCaptureDevices();
std::vector<AudioCaptureDeviceInfo> enumerateAudioCaptureDevices();
std::vector<MicrophoneDeviceInfo> enumerateMicrophoneDevices();
std::vector<SerialPortInfo> enumerateSerialPorts();
