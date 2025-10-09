#include "Application.hpp"
#include "DeviceEnumeration.hpp"

#ifndef MOD_NOREPEAT
#define MOD_NOREPEAT 0x4000
#endif


#include <algorithm>
#include <array>
#include <chrono>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <iomanip>
#include <sstream>
#include <string>
#include <string_view>
#include <stdexcept>
#include <thread>
#include <cmath>
#include <shellapi.h>

namespace
{
    constexpr wchar_t kWindowClassName[] = L"PCKVM.GC573.Window";
    constexpr int kDefaultWidth = 1920;
    constexpr int kDefaultHeight = 1080;

    constexpr UINT_PTR kTimerRenderDuringInteraction = 0x7101;
    const std::string kAudioSourceVideoSentinel = "@video";

    std::wstring utf8ToWide(const std::string& text)
    {
        if (text.empty())
        {
            return L"";
        }
        const int required = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), nullptr, 0);
        if (required <= 0)
        {
            return L"";
        }
        std::wstring result(static_cast<std::size_t>(required), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), result.data(), required);
        return result;
    }

    void logApp(const std::string& message)
    {
        std::ofstream("pckvm.log", std::ios::app) << message << '\n';
    }
}

Application::Application() = default;
Application::~Application()
{
    running_ = false;
    inputCaptureManager_.setEnabled(false);
    microphoneCapture_.stop();
    audioPlayback_.stop();
    serialStreamer_.stop();
    directShowCapture_.stop();
    renderer_.shutdown();
    unregisterMenuHotkey();
    destroyWindow();
}

int Application::run()
{
    {
        std::ofstream("pckvm.log", std::ios::trunc) << "[App] Launching viewer" << std::endl;
    }
    logApp("[App] Starting initialization");

    loadPersistentSettings();
    parseCommandLine();
    audioEnabled_ = shouldEnableCaptureAudio();
    logApp(std::string("[App] Audio capture ") + (audioEnabled_ ? "enabled" : "disabled"));

    if (!createWindow(kDefaultWidth, kDefaultHeight))
    {
        logApp("[App] Failed to create window");
        return EXIT_FAILURE;
    }

    if (!registerMenuHotkey())
    {
        logApp("[App] Failed to register menu hotkey");
    }

    if (!renderer_.initialize(hwnd_))
    {
        logApp("[App] Failed to initialize renderer");
        destroyWindow();
        return EXIT_FAILURE;
    }
    logApp("[App] Renderer initialized");

    if (!overlay_.initialize(hwnd_, renderer_))
    {
        logApp("[App] Failed to initialize ImGui overlay");
        // Continue without overlay
    }

    running_ = true;

    logApp("[App] Starting DirectShow capture");

    try
    {
        DirectShowCapture::Options captureOptions;
        captureOptions.deviceMoniker = settings_.videoDeviceMoniker;
        captureOptions.enableAudio = audioEnabled_;

        directShowCapture_.start([this](const DirectShowCapture::Frame& frame) {
            handleFrame(frame);
        }, captureOptions);
        logApp("[App] DirectShow capture started successfully");
    }
    catch (const std::exception& ex)
    {
        running_ = false;
        logApp(std::string("[App] DirectShow capture start failed: ") + ex.what());
        return EXIT_FAILURE;
    }
    catch (...)
    {
        running_ = false;
        logApp("[App] DirectShow capture start failed: unknown exception");
        return EXIT_FAILURE;
    }


    serialStreamer_.start();
    applySerialTargetSetting();
    applyInputCaptureSetting();
    applyMicrophoneCaptureSetting();
    applyAudioPlaybackSetting();

    logApp("[App] Entering render loop");
    renderLoop();
    logApp("[App] Render loop exited");

    inputCaptureManager_.setEnabled(false);
    microphoneCapture_.stop();
    audioPlayback_.stop();
    serialStreamer_.stop();

    directShowCapture_.stop();
    logApp("[App] DirectShow capture stopped");
    std::string captureError = directShowCapture_.consumeLastError();
    const bool anyFrames = frameCounter_.load(std::memory_order_acquire) > 0;

    overlay_.shutdown();
    renderer_.shutdown();
    logApp("[App] Renderer shutdown");

    if (captureError.empty() && !anyFrames)
    {
        const std::string deviceLabel = directShowCapture_.currentDeviceFriendlyName();
        captureError = "No video frames received from '" + (deviceLabel.empty() ? std::string("the selected capture device") : deviceLabel) + "'. Confirm a valid input signal and that no other application is using the device.";
    }

    if (!captureError.empty())
    {
        logApp(std::string("[App] Reporting error: ") + captureError);
    }

    unregisterMenuHotkey();
    destroyWindow();
    logApp("[App] Window destroyed");

    return captureError.empty() ? EXIT_SUCCESS : EXIT_FAILURE;
}

void Application::parseCommandLine()
{
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argv)
    {
        return;
    }

    bool enableAudio = settings_.audioPlaybackEnabled;
    for (int i = 1; i < argc; ++i)
    {
        std::wstring_view arg(argv[i]);
        if (arg == L"--enable-audio" || arg == L"--audio")
        {
            enableAudio = true;
        }
        else if (arg == L"--disable-audio" || arg == L"--no-audio")
        {
            enableAudio = false;
        }
    }

    if (enableAudio != settings_.audioPlaybackEnabled)
    {
        settings_.audioPlaybackEnabled = enableAudio;
        if (settings_.audioPlaybackEnabled && settings_.audioDeviceMoniker.empty())
        {
            settings_.audioDeviceMoniker = kAudioSourceVideoSentinel;
        }
        savePersistentSettings();
    }

    audioEnabled_ = shouldEnableCaptureAudio();
    LocalFree(argv);
}

LRESULT CALLBACK Application::windowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCT*>(lParam);
        auto* self = static_cast<Application*>(cs->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        return TRUE;
    }

    auto* self = reinterpret_cast<Application*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    if (!self)
    {
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    if (self->overlay_.processEvent(hwnd, msg, wParam, lParam))
    {
        return 1;
    }

    switch (msg)
    {
    case WM_SIZE:
    {
        const UINT width = LOWORD(lParam);
        const UINT height = HIWORD(lParam);
        self->renderer_.onResize(width, height);
        logApp("[App] WM_SIZE -> " + std::to_string(width) + "x" + std::to_string(height));
        self->updateInputCaptureBounds();
        return 0;
    }
    case WM_MOVE:
        self->updateInputCaptureBounds();
        return 0;
    case WM_ENTERSIZEMOVE:
        SetTimer(hwnd, kTimerRenderDuringInteraction, 16, nullptr);
        self->renderFrame(true);
        return 0;
    case WM_EXITSIZEMOVE:
        KillTimer(hwnd, kTimerRenderDuringInteraction);
        self->renderFrame(true);
        self->updateInputCaptureBounds();
        return 0;
    case WM_TIMER:
        if (wParam == kTimerRenderDuringInteraction)
        {
            self->renderFrame(true);
            return 0;
        }
        return 0;
    case WM_ACTIVATEAPP:
        if (wParam)
        {
            self->registerMenuHotkey();
        }
        else
        {
            self->unregisterMenuHotkey();
            self->inputCaptureManager_.clearModifierState();
        }
        self->updateInputCaptureBounds();
        return 0;
    case WM_SETFOCUS:
        self->registerMenuHotkey();
        self->updateInputCaptureBounds();
        return 0;
    case WM_KILLFOCUS:
        self->unregisterMenuHotkey();
        self->inputCaptureManager_.clearModifierState();
        self->updateInputCaptureBounds();
        return 0;
    case WM_SHOWWINDOW:
        if (!wParam)
        {
            self->inputCaptureManager_.clearModifierState();
        }
        self->updateInputCaptureBounds();
        return 0;
    case WM_ACTIVATE:
        if (LOWORD(wParam) == WA_INACTIVE)
        {
            self->inputCaptureManager_.clearModifierState();
        }
        self->updateInputCaptureBounds();
        return 0;
    case WM_KEYDOWN:
        if (wParam == 'G')
        {
            const bool newState = !self->renderer_.debugGradientEnabled();
            self->renderer_.setDebugGradient(newState);
            return 0;
        }
        break;
    case WM_GETMINMAXINFO:
        if (self->applyLockedWindowSize(reinterpret_cast<MINMAXINFO*>(lParam)))
        {
            return 0;
        }
        break;
    case WM_HOTKEY:
        if (self->ignoreMenuHotkeyUntil_ != 0)
        {
            const DWORD now = GetTickCount();
            if (now <= self->ignoreMenuHotkeyUntil_)
            {
                self->ignoreMenuHotkeyUntil_ = 0;
                return 0;
            }
            self->ignoreMenuHotkeyUntil_ = 0;
        }
        if (wParam == self->menuHotkeyId_ && self->isMenuHotkeySatisfied())
        {
            self->showSettingsMenu();
            return 0;
        }
        break;
    case WM_INPUT_CAPTURE_SHOW_MENU:
        self->ignoreMenuHotkeyUntil_ = GetTickCount() + 250;
        self->showSettingsMenu();
        return 0;
    case WM_INPUT_CAPTURE_UPDATE_CLIP:
        self->inputCaptureManager_.applyCursorClip(wParam != 0);
        return 0;
    case WM_CLOSE:
        logApp("[App] WM_CLOSE received");
        break;
    case WM_DESTROY:
        logApp("[App] WM_DESTROY received");
        PostQuitMessage(0);
        return 0;
    default:
        break;
    }

    return DefWindowProc(hwnd, msg, wParam, lParam);
}

bool Application::createWindow(int width, int height)
{
    logApp("[App] Registering window class");
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = &Application::windowProc;
    wc.hInstance = GetModuleHandle(nullptr);
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.lpszClassName = kWindowClassName;

    if (!RegisterClassExW(&wc))
    {
        logApp("[App] RegisterClassExW failed");
        return false;
    }
    classRegistered_ = true;

    DWORD style = WS_OVERLAPPEDWINDOW | WS_VISIBLE;
    RECT rect{0, 0, width, height};
    AdjustWindowRect(&rect, style, FALSE);

    hwnd_ = CreateWindowExW(
        WS_EX_APPWINDOW,
        kWindowClassName,
        L"CaptureKVM",
        style,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        rect.right - rect.left,
        rect.bottom - rect.top,
        nullptr,
        nullptr,
        GetModuleHandle(nullptr),
        this);

    if (!hwnd_)
    {
        logApp("[App] CreateWindowExW failed");
        return false;
    }

    if (!SetWindowTextW(hwnd_, L"CaptureKVM"))
    {
        logApp("[App] SetWindowTextW failed");
    }

    ShowWindow(hwnd_, SW_SHOW);
    UpdateWindow(hwnd_);
    inputCaptureManager_.setTargetWindow(hwnd_);
    updateInputCaptureBounds();
    RECT initialClient{};
    if (GetClientRect(hwnd_, &initialClient))
    {
        lockedClientWidth_ = initialClient.right - initialClient.left;
        lockedClientHeight_ = initialClient.bottom - initialClient.top;
    }
    updateWindowResizeMode();
    logApp("[App] Window created");
    return true;
}

void Application::destroyWindow()
{
    if (hwnd_)
    {
        inputCaptureManager_.setCaptureRegion(RECT{}, false);
        inputCaptureManager_.setTargetWindow(nullptr);
        DestroyWindow(hwnd_);
        hwnd_ = nullptr;
    }
    if (classRegistered_)
    {
        UnregisterClassW(kWindowClassName, GetModuleHandle(nullptr));
        classRegistered_ = false;
    }
}

void Application::handleFrame(const DirectShowCapture::Frame& frame)
{
    std::scoped_lock lock(frameMutex_);

    int backIndex = 1 - frontBufferIndex_;
    CpuFrame& dst = frames_[backIndex];

    dst.timestamp100ns = frame.timestamp100ns;
    dst.width = frame.width;
    dst.height = frame.height;

    const std::uint32_t frameWidth = frame.width;
    const std::uint32_t frameHeight = frame.height;
    const std::uint32_t knownWidth = currentSourceWidth_.load(std::memory_order_acquire);
    const std::uint32_t knownHeight = currentSourceHeight_.load(std::memory_order_acquire);
    if (frameWidth != knownWidth || frameHeight != knownHeight)
    {
        pendingSourceWidth_.store(frameWidth, std::memory_order_release);
        pendingSourceHeight_.store(frameHeight, std::memory_order_release);
        sourceChangePending_.store(true, std::memory_order_release);
    }

    inputCaptureManager_.setTargetResolution(static_cast<int>(frame.width), static_cast<int>(frame.height));

    const std::uint32_t stride = frame.stride != 0 ? frame.stride : frame.width * 4;
    const std::size_t requiredBytes = static_cast<std::size_t>(stride) * frame.height;
    if (frame.dataSize < requiredBytes)
    {
        logApp("[App] Warning: frame data shorter than expected (" + std::to_string(frame.dataSize) + " < " + std::to_string(requiredBytes) + ")");
    }
    dst.data.resize(requiredBytes);

    const bool bottomUp = frame.bottomUp;

    if (frame.data && frame.dataSize > 0)
    {
        const auto* srcRows = static_cast<const std::uint8_t*>(frame.data);
        auto* dstRows = dst.data.data();
        const std::size_t strideSize = stride;
        const std::size_t availableRows = strideSize != 0 ? std::min<std::size_t>(frame.height, frame.dataSize / strideSize) : 0;

        if (!bottomUp)
        {
            for (std::size_t y = 0; y < availableRows; ++y)
            {
                std::memcpy(dstRows + y * strideSize,
                            srcRows + y * strideSize,
                            strideSize);
            }
        }
        else
        {
            for (std::size_t y = 0; y < availableRows; ++y)
            {
                const std::size_t srcIndex = frame.height - 1 - y;
                if (srcIndex < availableRows)
                {
                    std::memcpy(dstRows + y * strideSize,
                                srcRows + srcIndex * strideSize,
                                strideSize);
                }
            }
        }

        const std::size_t copiedBytes = availableRows * strideSize;
        if (copiedBytes < requiredBytes)
        {
            std::memset(dst.data.data() + copiedBytes, 0, requiredBytes - copiedBytes);
        }
    }
    else
    {
        std::memset(dst.data.data(), 0, requiredBytes);
    }

    dst.stride = stride;

    static std::atomic<bool> loggedPixels{false};
    if (!loggedPixels.exchange(true))
    {
        logApp("[App] Stored frame size=" + std::to_string(dst.data.size()) + " stride=" + std::to_string(dst.stride));
        auto logPixel = [&](const char* label, std::size_t row, std::size_t col) {
            if (row < dst.height && col < dst.width)
            {
                const std::size_t offset = row * dst.stride + col * 4;
                if (offset + 3 < dst.data.size())
                {
                    const auto* px = dst.data.data() + offset;
                    std::ostringstream oss;
                    oss << "[App] Sample pixel " << label << " (row=" << row << ", col=" << col << ") = "
                        << std::hex << std::uppercase
                        << std::setw(2) << std::setfill('0') << static_cast<int>(px[0])
                        << std::setw(2) << static_cast<int>(px[1])
                        << std::setw(2) << static_cast<int>(px[2])
                        << std::setw(2) << static_cast<int>(px[3]);
                    logApp(oss.str());
                }
            }
        };
        logPixel("top-left", 0, 0);
        logPixel("center", dst.height / 2, dst.width / 2);
        logPixel("bottom-right", dst.height - 1, dst.width - 1);
    }

    frontBufferIndex_ = backIndex;
    frameCounter_.fetch_add(1, std::memory_order_acq_rel);

    static std::atomic<bool> logged{false};
    if (!logged.exchange(true))
    {
        logApp("[App] First frame received: " + std::to_string(dst.width) + "x" + std::to_string(dst.height) + " stride=" + std::to_string(dst.stride));
    }
}

void Application::renderLoop()
{
    MSG msg = {};

    while (running_)
    {
        while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
        {
            if (msg.message == WM_QUIT)
            {
                logApp("[App] WM_QUIT in render loop");
                running_ = false;
                break;
            }
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        if (sourceChangePending_.load(std::memory_order_acquire))
        {
            const std::uint32_t newWidth = pendingSourceWidth_.load(std::memory_order_acquire);
            const std::uint32_t newHeight = pendingSourceHeight_.load(std::memory_order_acquire);
            if (newWidth != 0 && newHeight != 0)
            {
                applySourceDimensions(newWidth, newHeight);
            }
            sourceChangePending_.store(false, std::memory_order_release);
        }

        renderFrame(false);
    }
}

void Application::loadPersistentSettings()
{
    settings_ = settingsManager_.load();
    settings_.inputTargetDevice.clear();
    if (settings_.menuHotkey.virtualKey == 0)
    {
        settings_.menuHotkey = SettingsManager::defaultMenuHotkey();
    }
    if (settings_.serialBaudRate == 0)
    {
        settings_.serialBaudRate = 921600;
    }
    if (settings_.audioPlaybackEnabled && settings_.audioDeviceMoniker.empty())
    {
        settings_.audioDeviceMoniker = kAudioSourceVideoSentinel;
    }
    settings_.mouseAbsoluteMode = true;
    inputCaptureManager_.setAbsoluteMode(settings_.mouseAbsoluteMode);
    audioEnabled_ = shouldEnableCaptureAudio();
}

void Application::savePersistentSettings()
{
    settingsManager_.save(settings_);
}

bool Application::registerMenuHotkey()
{
    if (!hwnd_)
    {
        menuHotkeyRegistered_ = false;
        inputCaptureManager_.setMenuChordEnabled(false);
        return false;
    }

    unregisterMenuHotkey();

    HotkeyConfig hotkey = settings_.menuHotkey;
    if (hotkey.virtualKey == 0)
    {
        hotkey = SettingsManager::defaultMenuHotkey();
        settings_.menuHotkey = hotkey;
        savePersistentSettings();
    }

    UINT modifiers = MOD_NOREPEAT;
    if (hotkey.requireCtrl || hotkey.requireRightCtrl)
    {
        modifiers |= MOD_CONTROL;
    }
    if (hotkey.requireShift)
    {
        modifiers |= MOD_SHIFT;
    }
    if (hotkey.requireAlt)
    {
        modifiers |= MOD_ALT;
    }
    if (hotkey.requireWin)
    {
        modifiers |= MOD_WIN;
    }

    inputCaptureManager_.setMenuChordEnabled(true);

    if (RegisterHotKey(hwnd_, static_cast<int>(menuHotkeyId_), modifiers, static_cast<UINT>(hotkey.virtualKey)))
    {
        menuHotkeyRegistered_ = true;
        return true;
    }

    menuHotkeyRegistered_ = false;
    return false;
}

void Application::unregisterMenuHotkey()
{
    if (hwnd_ && menuHotkeyRegistered_)
    {
        UnregisterHotKey(hwnd_, static_cast<int>(menuHotkeyId_));
    }
    menuHotkeyRegistered_ = false;
    inputCaptureManager_.setMenuChordEnabled(false);
}

void Application::showSettingsMenu()
{
    overlay_.toggleMenu(*this);
}

bool Application::isMenuHotkeySatisfied() const
{
    const HotkeyConfig& hotkey = settings_.menuHotkey.virtualKey != 0 ? settings_.menuHotkey : SettingsManager::defaultMenuHotkey();

    if (hotkey.chordVirtualKey != 0)
    {
        if ((GetAsyncKeyState(static_cast<int>(hotkey.chordVirtualKey)) & 0x8000) == 0)
        {
            return false;
        }
    }

    if (hotkey.requireRightCtrl)
    {
        if ((GetAsyncKeyState(VK_RCONTROL) & 0x8000) == 0)
        {
            return false;
        }
    }
    else if (hotkey.requireCtrl)
    {
        if ((GetAsyncKeyState(VK_CONTROL) & 0x8000) == 0)
        {
            return false;
        }
    }

    if (hotkey.requireShift && (GetAsyncKeyState(VK_SHIFT) & 0x8000) == 0)
    {
        return false;
    }
    if (hotkey.requireAlt && (GetAsyncKeyState(VK_MENU) & 0x8000) == 0)
    {
        return false;
    }
    if (hotkey.requireWin)
    {
        const bool leftWin = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0;
        const bool rightWin = (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;
        if (!(leftWin || rightWin))
        {
            return false;
        }
    }

    return true;
}

void Application::applyAudioPlaybackSetting()
{
    const bool useVideoAudio = shouldUseVideoAudio();
    if (settings_.audioPlaybackEnabled && useVideoAudio && settings_.audioDeviceMoniker != kAudioSourceVideoSentinel)
    {
        if (settings_.audioDeviceMoniker.empty() || settings_.audioDeviceMoniker == settings_.videoDeviceMoniker)
        {
            settings_.audioDeviceMoniker = kAudioSourceVideoSentinel;
            savePersistentSettings();
        }
    }

    const bool desiredCaptureAudio = shouldEnableCaptureAudio();

    if (desiredCaptureAudio != audioEnabled_)
    {
        audioEnabled_ = desiredCaptureAudio;
        restartVideoCapture();
    }

    if (settings_.audioPlaybackEnabled && !useVideoAudio)
    {
        if (!settings_.audioDeviceMoniker.empty())
        {
            audioPlayback_.start(settings_.audioDeviceMoniker);
        }
        else
        {
            audioPlayback_.stop();
        }
    }
    else
    {
        audioPlayback_.stop();
    }
}

void Application::applyInputCaptureSetting()
{
    settings_.mouseAbsoluteMode = true;
    inputCaptureManager_.setAbsoluteMode(true);
    inputCaptureManager_.setEnabled(settings_.inputCaptureEnabled);
    if (menuHotkeyRegistered_)
    {
        inputCaptureManager_.setMenuChordEnabled(true);
    }
    else
    {
        inputCaptureManager_.setMenuChordEnabled(false);
    }
}

void Application::applyMicrophoneCaptureSetting()
{
    if (settings_.microphoneCaptureEnabled)
    {
        microphoneCapture_.start(settings_.microphoneDeviceId, serialStreamer_, settings_.microphoneAutoGain);
    }
    else
    {
        microphoneCapture_.stop();
    }
}

void Application::applySerialTargetSetting()
{
    if (!serialStreamer_.isRunning())
    {
        serialStreamer_.start();
    }
    serialStreamer_.setBaudRate(settings_.serialBaudRate);
    serialStreamer_.requestReconnect();
}

void Application::setAudioPlaybackEnabled(bool enabled)
{
    if (settings_.audioPlaybackEnabled == enabled)
    {
        return;
    }

    settings_.audioPlaybackEnabled = enabled;
    if (settings_.audioPlaybackEnabled && settings_.audioDeviceMoniker.empty())
    {
        settings_.audioDeviceMoniker = kAudioSourceVideoSentinel;
    }
    savePersistentSettings();
    logApp(std::string("[App] Audio playback toggled -> ") + (settings_.audioPlaybackEnabled ? "enabled" : "disabled"));
    applyAudioPlaybackSetting();
}

void Application::setMicrophoneCaptureEnabled(bool enabled)
{
    if (settings_.microphoneCaptureEnabled == enabled)
    {
        return;
    }

    settings_.microphoneCaptureEnabled = enabled;
    savePersistentSettings();
    logApp(std::string("[App] Microphone capture toggled -> ") + (settings_.microphoneCaptureEnabled ? "enabled" : "disabled"));
    applyMicrophoneCaptureSetting();
}

void Application::setInputCaptureEnabled(bool enabled)
{
    if (settings_.inputCaptureEnabled == enabled)
    {
        return;
    }

    settings_.inputCaptureEnabled = enabled;
    savePersistentSettings();
    logApp(std::string("[App] Input capture toggled -> ") + (settings_.inputCaptureEnabled ? "enabled" : "disabled"));
    applyInputCaptureSetting();
}

void Application::selectVideoDevice(const std::string& moniker)
{
    if (settings_.videoDeviceMoniker == moniker)
    {
        return;
    }

    settings_.videoDeviceMoniker = moniker;
    savePersistentSettings();
    logApp(std::string("[App] Selected video capture device: ") + settings_.videoDeviceMoniker);
    restartVideoCapture();
    if (settings_.audioDeviceMoniker == kAudioSourceVideoSentinel && settings_.audioPlaybackEnabled)
    {
        applyAudioPlaybackSetting();
    }
    requestImmediateRender();
}

void Application::selectAudioDevice(const std::string& moniker)
{
    std::string newMoniker = moniker.empty() ? kAudioSourceVideoSentinel : moniker;
    if (settings_.audioDeviceMoniker == newMoniker)
    {
        return;
    }

    settings_.audioDeviceMoniker = newMoniker;
    savePersistentSettings();
    const std::string logLabel = (settings_.audioDeviceMoniker == kAudioSourceVideoSentinel) ? std::string("video source audio") : settings_.audioDeviceMoniker;
    logApp(std::string("[App] Selected audio capture device: ") + logLabel);
    applyAudioPlaybackSetting();
    requestImmediateRender();
}

void Application::selectMicrophoneDevice(const std::string& endpointId)
{
    if (settings_.microphoneDeviceId == endpointId)
    {
        return;
    }

    settings_.microphoneDeviceId = endpointId;
    savePersistentSettings();
    logApp(std::string("[App] Selected microphone device: ") + settings_.microphoneDeviceId);
    if (settings_.microphoneCaptureEnabled)
    {
        applyMicrophoneCaptureSetting();
    }
    requestImmediateRender();
}

void Application::setVideoAllowResizing(bool enabled)
{
    if (settings_.videoAllowResizing == enabled)
    {
        return;
    }

    settings_.videoAllowResizing = enabled;
    savePersistentSettings();
    logApp(std::string("[App] Video allow resizing -> ") + (settings_.videoAllowResizing ? "enabled" : "disabled"));
    updateWindowResizeMode();

    if (!settings_.videoAllowResizing)
    {
        const std::uint32_t srcW = currentSourceWidth_.load(std::memory_order_acquire);
        const std::uint32_t srcH = currentSourceHeight_.load(std::memory_order_acquire);
        if (srcW > 0 && srcH > 0)
        {
            lockedClientWidth_ = static_cast<int>(srcW);
            lockedClientHeight_ = static_cast<int>(srcH);
        }
        else if (hwnd_)
        {
            RECT client{};
            if (GetClientRect(hwnd_, &client))
            {
                lockedClientWidth_ = client.right - client.left;
                lockedClientHeight_ = client.bottom - client.top;
            }
        }

        if (lockedClientWidth_ > 0 && lockedClientHeight_ > 0)
        {
            resizeWindowToClient(lockedClientWidth_, lockedClientHeight_);
        }
    }

    updateInputCaptureBounds();
    requestImmediateRender();
}

void Application::setVideoAspectMode(VideoAspectMode mode)
{
    if (settings_.videoAspectMode == mode)
    {
        return;
    }

    settings_.videoAspectMode = mode;
    savePersistentSettings();
    logApp(std::string("[App] Video aspect mode -> ") + std::to_string(static_cast<unsigned int>(mode)));
    updateInputCaptureBounds();
    requestImmediateRender();
}

void Application::requestImmediateRender()
{
    forceRender_.store(true, std::memory_order_release);
}

bool Application::uploadLatestFrame()
{
    std::unique_lock<std::mutex> lock(frameMutex_, std::try_to_lock);
    if (!lock.owns_lock())
    {
        return false;
    }

    const std::uint64_t latest = frameCounter_.load(std::memory_order_acquire);
    if (latest == lastPresentedFrame_)
    {
        return false;
    }

    const CpuFrame& src = frames_[frontBufferIndex_];
    if (src.data.empty() || src.width == 0 || src.height == 0)
    {
        return false;
    }

    renderer_.uploadFrame(src.data.data(), src.stride, src.width, src.height);
    lastPresentedFrame_ = latest;
    return true;
}

void Application::renderFrame(bool forcePresent)
{
    overlay_.newFrame();
    overlay_.buildUI(*this);
    overlay_.endFrame();

    const bool uploaded = uploadLatestFrame();
    const bool forced = forcePresent || forceRender_.exchange(false, std::memory_order_acq_rel);
    const bool overlayHasDraw = overlay_.hasDrawData();
    const bool hasFrame = (lastPresentedFrame_ != 0);

    if (uploaded || forced || overlayHasDraw || (forcePresent && hasFrame))
    {
        renderer_.render([&](ID3D12GraphicsCommandList* cmdList) {
            overlay_.render(cmdList);
        });
    }
    else if (!forcePresent)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
}

void Application::applySourceDimensions(std::uint32_t width, std::uint32_t height)
{
    if (width == 0 || height == 0)
    {
        return;
    }

    currentSourceWidth_.store(width, std::memory_order_release);
    currentSourceHeight_.store(height, std::memory_order_release);

    lockedClientWidth_ = static_cast<int>(width);
    lockedClientHeight_ = static_cast<int>(height);

    if (hwnd_)
    {
        resizeWindowToClient(static_cast<int>(width), static_cast<int>(height));
        updateWindowResizeMode();
        updateInputCaptureBounds();
    }
}

bool Application::resizeWindowToClient(int width, int height)
{
    if (!hwnd_ || width <= 0 || height <= 0)
    {
        return false;
    }

    RECT current{};
    if (!GetClientRect(hwnd_, &current))
    {
        return false;
    }

    const int currentWidth = current.right - current.left;
    const int currentHeight = current.bottom - current.top;
    if (currentWidth == width && currentHeight == height)
    {
        return false;
    }

    RECT desired{0, 0, width, height};
    DWORD style = static_cast<DWORD>(GetWindowLongPtr(hwnd_, GWL_STYLE));
    DWORD exStyle = static_cast<DWORD>(GetWindowLongPtr(hwnd_, GWL_EXSTYLE));
    if (!AdjustWindowRectEx(&desired, style, FALSE, exStyle))
    {
        return false;
    }

    const int windowWidth = desired.right - desired.left;
    const int windowHeight = desired.bottom - desired.top;

    if (!SetWindowPos(hwnd_, nullptr, 0, 0, windowWidth, windowHeight,
                      SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOSENDCHANGING))
    {
        return false;
    }

    return true;
}

void Application::updateWindowResizeMode()
{
    if (!hwnd_)
    {
        return;
    }

    LONG_PTR style = GetWindowLongPtr(hwnd_, GWL_STYLE);
    if (style == 0)
    {
        return;
    }

    const LONG_PTR desiredStyle = settings_.videoAllowResizing
        ? (style | WS_THICKFRAME | WS_MAXIMIZEBOX)
        : (style & ~(WS_THICKFRAME | WS_MAXIMIZEBOX));

    if (desiredStyle != style)
    {
        SetWindowLongPtr(hwnd_, GWL_STYLE, desiredStyle);
        SetWindowPos(hwnd_, nullptr, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    }
}

bool Application::applyLockedWindowSize(MINMAXINFO* info) const
{
    if (!info || settings_.videoAllowResizing || !hwnd_ || lockedClientWidth_ <= 0 || lockedClientHeight_ <= 0)
    {
        return false;
    }

    RECT rect{0, 0, lockedClientWidth_, lockedClientHeight_};
    DWORD style = static_cast<DWORD>(GetWindowLongPtr(hwnd_, GWL_STYLE));
    DWORD exStyle = static_cast<DWORD>(GetWindowLongPtr(hwnd_, GWL_EXSTYLE));
    if (!AdjustWindowRectEx(&rect, style, FALSE, exStyle))
    {
        return false;
    }

    const int width = rect.right - rect.left;
    const int height = rect.bottom - rect.top;
    info->ptMinTrackSize.x = width;
    info->ptMinTrackSize.y = height;
    info->ptMaxTrackSize.x = width;
    info->ptMaxTrackSize.y = height;
    return true;
}

RECT Application::computeVideoViewport(const RECT& clientRect, bool& valid) const
{
    valid = false;
    RECT viewport{0, 0, 0, 0};

    const LONG clientWidth = clientRect.right - clientRect.left;
    const LONG clientHeight = clientRect.bottom - clientRect.top;
    if (clientWidth <= 0 || clientHeight <= 0)
    {
        return viewport;
    }

    const std::uint32_t srcWidth = currentSourceWidth_.load(std::memory_order_acquire);
    const std::uint32_t srcHeight = currentSourceHeight_.load(std::memory_order_acquire);
    if (srcWidth == 0 || srcHeight == 0)
    {
        return viewport;
    }

    switch (settings_.videoAspectMode)
    {
    case VideoAspectMode::Stretch:
        viewport.left = 0;
        viewport.top = 0;
        viewport.right = clientWidth;
        viewport.bottom = clientHeight;
        valid = true;
        return viewport;

    case VideoAspectMode::Maintain:
    {
        const double srcAspect = static_cast<double>(srcWidth) / static_cast<double>(srcHeight);
        const double clientAspect = static_cast<double>(clientWidth) / static_cast<double>(clientHeight);
        constexpr double epsilon = 1e-4;

        int viewportWidth = static_cast<int>(clientWidth);
        int viewportHeight = static_cast<int>(clientHeight);
        if (std::abs(clientAspect - srcAspect) > epsilon)
        {
            if (clientAspect > srcAspect)
            {
                viewportHeight = static_cast<int>(clientHeight);
                viewportWidth = static_cast<int>(std::round(static_cast<double>(viewportHeight) * srcAspect));
            }
            else
            {
                viewportWidth = static_cast<int>(clientWidth);
                viewportHeight = static_cast<int>(std::round(static_cast<double>(viewportWidth) / srcAspect));
            }
        }

        viewportWidth = std::max(1, std::min(viewportWidth, static_cast<int>(clientWidth)));
        viewportHeight = std::max(1, std::min(viewportHeight, static_cast<int>(clientHeight)));

        const int offsetX = (static_cast<int>(clientWidth) - viewportWidth) / 2;
        const int offsetY = (static_cast<int>(clientHeight) - viewportHeight) / 2;

        viewport.left = offsetX;
        viewport.top = offsetY;
        viewport.right = offsetX + viewportWidth;
        viewport.bottom = offsetY + viewportHeight;
        valid = true;
        return viewport;
    }

    case VideoAspectMode::Capture:
    {
        double scale = std::min<double>({static_cast<double>(clientWidth) / static_cast<double>(srcWidth),
                                         static_cast<double>(clientHeight) / static_cast<double>(srcHeight),
                                         1.0});
        int viewportWidth = static_cast<int>(std::round(static_cast<double>(srcWidth) * scale));
        int viewportHeight = static_cast<int>(std::round(static_cast<double>(srcHeight) * scale));
        viewportWidth = std::max(1, std::min(viewportWidth, static_cast<int>(clientWidth)));
        viewportHeight = std::max(1, std::min(viewportHeight, static_cast<int>(clientHeight)));
        const int offsetX = (static_cast<int>(clientWidth) - viewportWidth) / 2;
        const int offsetY = (static_cast<int>(clientHeight) - viewportHeight) / 2;
        viewport.left = offsetX;
        viewport.top = offsetY;
        viewport.right = offsetX + viewportWidth;
        viewport.bottom = offsetY + viewportHeight;
        valid = true;
        return viewport;
    }
    }

    return viewport;
}

void Application::updateInputCaptureBounds()
{
    if (!hwnd_ || !IsWindowVisible(hwnd_))
    {
        inputCaptureManager_.setCaptureRegion(RECT{}, false);
        inputCaptureManager_.setVideoViewport(RECT{}, false);
        renderer_.setViewportRect(0.0f, 0.0f, 0.0f, 0.0f);
        return;
    }

    RECT client{};
    if (!GetClientRect(hwnd_, &client))
    {
        inputCaptureManager_.setCaptureRegion(RECT{}, false);
        inputCaptureManager_.setVideoViewport(RECT{}, false);
        return;
    }

    POINT topLeft{client.left, client.top};
    POINT bottomRight{client.right, client.bottom};
    ClientToScreen(hwnd_, &topLeft);
    ClientToScreen(hwnd_, &bottomRight);

    RECT screenRect{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};

    const bool windowHasArea = (screenRect.right > screenRect.left) && (screenRect.bottom > screenRect.top);
    const bool windowActive = !IsIconic(hwnd_) && GetForegroundWindow() == hwnd_ && windowHasArea;
    inputCaptureManager_.setCaptureRegion(screenRect, windowActive);

    bool viewportValid = false;
    RECT viewportClient = computeVideoViewport(client, viewportValid);

    if (!viewportValid)
    {
        inputCaptureManager_.setVideoViewport(RECT{}, false);
        const LONG clientWidth = client.right - client.left;
        const LONG clientHeight = client.bottom - client.top;
        renderer_.setViewportRect(0.0f,
                                  0.0f,
                                  static_cast<float>(std::max<LONG>(clientWidth, 0)),
                                  static_cast<float>(std::max<LONG>(clientHeight, 0)));
        return;
    }

    RECT viewportScreen{
        topLeft.x + viewportClient.left,
        topLeft.y + viewportClient.top,
        topLeft.x + viewportClient.right,
        topLeft.y + viewportClient.bottom
    };

    inputCaptureManager_.setVideoViewport(viewportScreen, viewportValid && windowActive);

    renderer_.setViewportRect(static_cast<float>(viewportClient.left),
                              static_cast<float>(viewportClient.top),
                              static_cast<float>(viewportClient.right - viewportClient.left),
                              static_cast<float>(viewportClient.bottom - viewportClient.top));
}

bool Application::shouldUseVideoAudio() const
{
    if (settings_.audioDeviceMoniker.empty())
    {
        return true;
    }
    if (settings_.audioDeviceMoniker == kAudioSourceVideoSentinel)
    {
        return true;
    }
    if (!settings_.videoDeviceMoniker.empty() && settings_.audioDeviceMoniker == settings_.videoDeviceMoniker)
    {
        return true;
    }
    return false;
}

bool Application::shouldEnableCaptureAudio() const
{
    return settings_.audioPlaybackEnabled && shouldUseVideoAudio();
}

void Application::restartVideoCapture()
{
    if (!running_)
    {
        return;
    }

    logApp("[App] Restarting video capture with updated settings");
    directShowCapture_.stop();

    {
        std::lock_guard<std::mutex> lock(frameMutex_);
        frames_[0] = CpuFrame{};
        frames_[1] = CpuFrame{};
    }
    frameCounter_.store(0, std::memory_order_release);
    lastPresentedFrame_ = 0;

    try
    {
        DirectShowCapture::Options options;
        options.deviceMoniker = settings_.videoDeviceMoniker;
        options.enableAudio = audioEnabled_;
        directShowCapture_.start([this](const DirectShowCapture::Frame& frame) {
            handleFrame(frame);
        }, options);
        logApp("[App] Video capture restarted successfully");
    }
    catch (const std::exception& ex)
    {
        logApp(std::string("[App] Failed to restart capture: ") + ex.what());
    }
    catch (...)
    {
        logApp("[App] Failed to restart capture: unknown error");
    }
}
