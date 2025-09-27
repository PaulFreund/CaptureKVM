#pragma once

#define UNICODE

#include "DirectShowCapture.hpp"
#include "D3DRenderer.hpp"
#include "Settings.hpp"
#include "SerialStreamer.hpp"
#include "InputCapture.hpp"
#include "MicrophoneCapture.hpp"
#include "AudioPlayback.hpp"
#include <unordered_map>
#include <vector>

#include <Windows.h>
#include <atomic>
#include <mutex>
#include <vector>

class Application {
public:
    Application();
    ~Application();

    int run();

private:
    struct CpuFrame {
        std::uint32_t width = 0;
        std::uint32_t height = 0;
        std::uint32_t stride = 0;
        std::uint64_t timestamp100ns = 0;
        std::vector<std::uint8_t> data;
    };

    static LRESULT CALLBACK windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    bool createWindow(int width, int height);
    void destroyWindow();
    void handleFrame(const DirectShowCapture::Frame& frame);
    void renderLoop();
    void parseCommandLine();
    void loadPersistentSettings();
    void savePersistentSettings();
    bool registerMenuHotkey();
    void unregisterMenuHotkey();
    void showSettingsMenu();
    void handleMenuSelection(unsigned int commandId);
    bool isMenuHotkeySatisfied() const;
    void applyAudioPlaybackSetting();
    void applyInputCaptureSetting();
    void applyMicrophoneCaptureSetting();
    void applySerialTargetSetting();
    void updateInputCaptureBounds();
    void restartVideoCapture();
    bool shouldUseVideoAudio() const;
    bool shouldEnableCaptureAudio() const;

    HWND hwnd_ = nullptr;
    D3DRenderer renderer_;
    DirectShowCapture directShowCapture_;

    std::mutex frameMutex_;
    CpuFrame frames_[2];
    std::atomic<std::uint64_t> frameCounter_{0};
    std::uint64_t lastPresentedFrame_ = 0;
    int frontBufferIndex_ = 0;
    bool running_ = false;
    bool classRegistered_ = false;
    bool audioEnabled_ = false;

    SerialStreamer serialStreamer_;
    InputCaptureManager inputCaptureManager_{serialStreamer_};
    MicrophoneCapture microphoneCapture_;
    AudioPlayback audioPlayback_;

    SettingsManager settingsManager_;
    AppSettings settings_{};
    unsigned int menuHotkeyId_ = 1;
    std::unordered_map<unsigned int, std::string> videoCommandMap_;
    std::unordered_map<unsigned int, std::string> audioCommandMap_;
    std::unordered_map<unsigned int, std::string> microphoneCommandMap_;
    DWORD ignoreMenuHotkeyUntil_ = 0;
};
