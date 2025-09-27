#pragma once

#include <string>
#include <filesystem>

struct HotkeyConfig {
    unsigned int virtualKey = 0;
    unsigned int chordVirtualKey = 0;
    bool requireCtrl = false;
    bool requireRightCtrl = false;
    bool requireShift = false;
    bool requireAlt = false;
    bool requireWin = false;
};

struct AppSettings {
    std::string videoDeviceMoniker;
    std::string audioDeviceMoniker;
    bool audioPlaybackEnabled = true;
    bool microphoneCaptureEnabled = false;
    std::string microphoneDeviceId;
    bool microphoneAutoGain = true;
    bool inputCaptureEnabled = true;
    bool mouseAbsoluteMode = true;
    std::string inputTargetDevice;
    unsigned int serialBaudRate = 921600;
    HotkeyConfig menuHotkey;
};

class SettingsManager {
public:
    SettingsManager();

    AppSettings load();
    void save(const AppSettings& settings) const;

    [[nodiscard]] const std::filesystem::path& settingsFile() const noexcept { return settingsFile_; }

    static HotkeyConfig defaultMenuHotkey();

private:
    std::filesystem::path settingsFile_;
    static std::filesystem::path determineSettingsPath();
};
