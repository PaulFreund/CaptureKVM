#pragma once

#define UNICODE

#include "DirectShowCapture.hpp"
#include "D3DRenderer.hpp"
#include "Settings.hpp"
#include "SerialStreamer.hpp"
#include "InputCapture.hpp"
#include "MicrophoneCapture.hpp"
#include "AudioPlayback.hpp"
#include "OverlayUI.hpp"
#include "DeviceEnumeration.hpp"

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
    friend class OverlayUI;
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
    bool isMenuHotkeySatisfied() const;
    void applyAudioPlaybackSetting();
    void applyInputCaptureSetting();
    void applyMicrophoneCaptureSetting();
    void applySerialTargetSetting();
    void updateInputCaptureBounds();
    void restartVideoCapture();
    bool shouldUseVideoAudio() const;
    bool shouldEnableCaptureAudio() const;
    void applySourceDimensions(std::uint32_t width, std::uint32_t height);
    bool resizeWindowToClient(int width, int height);
    void updateWindowResizeMode();
    bool applyLockedWindowSize(MINMAXINFO* info) const;
    RECT computeVideoViewport(const RECT& clientRect, bool& valid) const;
    bool uploadLatestFrame();
    void renderFrame(bool forcePresent);
    void setAudioPlaybackEnabled(bool enabled);
    void setMicrophoneCaptureEnabled(bool enabled);
    void setInputCaptureEnabled(bool enabled);
    void selectVideoDevice(const std::string& moniker);
    void selectAudioDevice(const std::string& moniker);
    void selectMicrophoneDevice(const std::string& endpointId);
    void setVideoResolution(std::uint32_t width, std::uint32_t height);
    void setVideoAllowResizing(bool enabled);
    void setVideoAspectMode(VideoAspectMode mode);
    void requestImmediateRender();
    void processPendingSourceDimensions();
    void selectBridgeDevice(const SerialPortInfo& info, bool autoSelect);
    bool classifyBridgeDevice(const SerialPortInfo& info, unsigned int* outBaud) const;
    static std::string toLowerCopy(const std::string& text);
    HWND hwnd() const { return hwnd_; }
    AppSettings& settings() { return settings_; }
    const AppSettings& settings() const { return settings_; }
    std::uint32_t currentCaptureWidth() const { return currentSourceWidth_.load(std::memory_order_acquire); }
    std::uint32_t currentCaptureHeight() const { return currentSourceHeight_.load(std::memory_order_acquire); }

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
    OverlayUI overlay_;

    SettingsManager settingsManager_;
    AppSettings settings_{};
    unsigned int menuHotkeyId_ = 1;
    DWORD ignoreMenuHotkeyUntil_ = 0;
    bool menuHotkeyRegistered_ = false;
    std::atomic<std::uint32_t> pendingSourceWidth_{0};
    std::atomic<std::uint32_t> pendingSourceHeight_{0};
    std::atomic<bool> sourceChangePending_{false};
    std::atomic<std::uint32_t> currentSourceWidth_{0};
    std::atomic<std::uint32_t> currentSourceHeight_{0};
    int lockedClientWidth_ = 0;
    int lockedClientHeight_ = 0;
    std::atomic<bool> forceRender_{false};
};
