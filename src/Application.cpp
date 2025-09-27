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
#include <shellapi.h>

namespace
{
    constexpr wchar_t kWindowClassName[] = L"PCKVM.GC573.Window";
    constexpr int kDefaultWidth = 1920;
    constexpr int kDefaultHeight = 1080;

    constexpr unsigned int kCommandToggleAudioPlayback = 0x0100;
    constexpr unsigned int kCommandToggleMicrophoneCapture = 0x0101;
    constexpr unsigned int kCommandToggleInputCapture = 0x0102;
    constexpr unsigned int kCommandVideoDeviceBase = 0x1000;
    constexpr unsigned int kCommandAudioUseVideoSource = 0x10FE;
    constexpr unsigned int kCommandAudioDeviceBase = 0x1100;
    constexpr unsigned int kCommandMicrophoneDeviceBase = 0x1200;
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
    case WM_SHOWWINDOW:
    case WM_ACTIVATE:
    case WM_ACTIVATEAPP:
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
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

        bool presented = false;
        {
            std::unique_lock<std::mutex> lock(frameMutex_, std::try_to_lock);
            if (lock.owns_lock())
            {
                const std::uint64_t latest = frameCounter_.load(std::memory_order_acquire);
                if (latest != lastPresentedFrame_)
                {
                    const CpuFrame& src = frames_[frontBufferIndex_];
                    renderer_.uploadFrame(src.data.data(), src.stride, src.width, src.height);
                    lastPresentedFrame_ = latest;
                    presented = true;
                }
            }
        }

        if (presented)
        {
            renderer_.render();
        }
        else
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
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

    if (RegisterHotKey(hwnd_, static_cast<int>(menuHotkeyId_), modifiers, static_cast<UINT>(hotkey.virtualKey)))
    {
        return true;
    }

    return false;
}

void Application::unregisterMenuHotkey()
{
    if (hwnd_)
    {
        UnregisterHotKey(hwnd_, static_cast<int>(menuHotkeyId_));
    }
}

void Application::showSettingsMenu()
{
    videoCommandMap_.clear();
    audioCommandMap_.clear();
    microphoneCommandMap_.clear();

    HMENU rootMenu = CreatePopupMenu();
    if (!rootMenu)
    {
        return;
    }

    auto videoDevices = enumerateVideoCaptureDevices();
    std::wstring selectedVideoFriendly;
    HMENU videoMenu = CreatePopupMenu();
    if (videoMenu)
    {
        if (videoDevices.empty())
        {
            AppendMenuW(videoMenu, MF_STRING | MF_DISABLED | MF_GRAYED, 0, L"No video capture devices");
        }
        else
        {
            for (std::size_t i = 0; i < videoDevices.size(); ++i)
            {
                const auto& device = videoDevices[i];
                const UINT commandId = kCommandVideoDeviceBase + static_cast<UINT>(i);
                videoCommandMap_[commandId] = device.monikerDisplayName;
                std::wstring label = utf8ToWide(device.friendlyName.empty() ? device.monikerDisplayName : device.friendlyName);
                UINT flags = MF_STRING;
                if (!settings_.videoDeviceMoniker.empty() && settings_.videoDeviceMoniker == device.monikerDisplayName)
                {
                    flags |= MF_CHECKED;
                    selectedVideoFriendly = label;
                }
                AppendMenuW(videoMenu, flags, commandId, label.c_str());
            }
        }
        AppendMenuW(rootMenu, MF_POPUP | (videoDevices.empty() ? MF_DISABLED | MF_GRAYED : 0u), reinterpret_cast<UINT_PTR>(videoMenu), L"Video Capture Source");
    }

    auto audioDevices = enumerateAudioCaptureDevices();
    HMENU audioMenu = CreatePopupMenu();
    if (audioMenu)
    {
        bool anyAudioEntries = false;
        if (!settings_.videoDeviceMoniker.empty())
        {
            std::wstring label = L"Use Video Source Audio";
            if (!selectedVideoFriendly.empty())
            {
                label += L" (" + selectedVideoFriendly + L")";
            }
            UINT flags = MF_STRING;
            if (shouldUseVideoAudio())
            {
                flags |= MF_CHECKED;
            }
            AppendMenuW(audioMenu, flags, kCommandAudioUseVideoSource, label.c_str());
            audioCommandMap_[kCommandAudioUseVideoSource] = kAudioSourceVideoSentinel;
            anyAudioEntries = true;
        }

        if (!audioDevices.empty())
        {
            if (anyAudioEntries)
            {
                AppendMenuW(audioMenu, MF_SEPARATOR, 0, nullptr);
            }
            for (std::size_t i = 0; i < audioDevices.size(); ++i)
            {
                const auto& device = audioDevices[i];
                const UINT commandId = kCommandAudioDeviceBase + static_cast<UINT>(i);
                audioCommandMap_[commandId] = device.monikerDisplayName;
                std::wstring label = utf8ToWide(device.friendlyName.empty() ? device.monikerDisplayName : device.friendlyName);
                UINT flags = MF_STRING;
                if (!shouldUseVideoAudio() && !settings_.audioDeviceMoniker.empty() && settings_.audioDeviceMoniker == device.monikerDisplayName)
                {
                    flags |= MF_CHECKED;
                }
                AppendMenuW(audioMenu, flags, commandId, label.c_str());
            }
            anyAudioEntries = true;
        }

        if (!anyAudioEntries)
        {
            AppendMenuW(audioMenu, MF_STRING | MF_DISABLED | MF_GRAYED, 0, L"No audio capture devices");
        }

        AppendMenuW(rootMenu, MF_POPUP | (!anyAudioEntries ? MF_DISABLED | MF_GRAYED : 0u), reinterpret_cast<UINT_PTR>(audioMenu), L"Audio Capture Source");
    }

    auto microphones = enumerateMicrophoneDevices();
    HMENU microphoneMenu = CreatePopupMenu();
    if (microphoneMenu)
    {
        if (microphones.empty())
        {
            AppendMenuW(microphoneMenu, MF_STRING | MF_DISABLED | MF_GRAYED, 0, L"No microphone devices");
        }
        else
        {
            for (std::size_t i = 0; i < microphones.size(); ++i)
            {
                const auto& device = microphones[i];
                const UINT commandId = kCommandMicrophoneDeviceBase + static_cast<UINT>(i);
                microphoneCommandMap_[commandId] = device.endpointId;
                std::wstring label = utf8ToWide(device.friendlyName.empty() ? device.endpointId : device.friendlyName);
                UINT flags = MF_STRING;
                if (!settings_.microphoneDeviceId.empty() && settings_.microphoneDeviceId == device.endpointId)
                {
                    flags |= MF_CHECKED;
                }
                AppendMenuW(microphoneMenu, flags, commandId, label.c_str());
            }
        }
        AppendMenuW(rootMenu, MF_POPUP | (microphones.empty() ? MF_DISABLED | MF_GRAYED : 0u), reinterpret_cast<UINT_PTR>(microphoneMenu), L"Microphone Device");
    }

    AppendMenuW(rootMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(rootMenu, MF_STRING | (settings_.audioPlaybackEnabled ? MF_CHECKED : 0u), kCommandToggleAudioPlayback, L"Enable Audio Playback");
    AppendMenuW(rootMenu, MF_STRING | (settings_.microphoneCaptureEnabled ? MF_CHECKED : 0u), kCommandToggleMicrophoneCapture, L"Enable Microphone Capture");
    AppendMenuW(rootMenu, MF_STRING | (settings_.inputCaptureEnabled ? MF_CHECKED : 0u), kCommandToggleInputCapture, L"Enable Keyboard && Mouse Capture");

    POINT cursor{};
    GetCursorPos(&cursor);
    SetForegroundWindow(hwnd_);
    const UINT command = TrackPopupMenu(rootMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, cursor.x, cursor.y, 0, hwnd_, nullptr);
    DestroyMenu(rootMenu);

    if (command != 0)
    {
        handleMenuSelection(command);
    }
}

void Application::handleMenuSelection(unsigned int commandId)
{
    if (commandId == 0)
    {
        return;
    }

    if (commandId == kCommandToggleAudioPlayback)
    {
        settings_.audioPlaybackEnabled = !settings_.audioPlaybackEnabled;
        if (settings_.audioPlaybackEnabled && settings_.audioDeviceMoniker.empty())
        {
            settings_.audioDeviceMoniker = kAudioSourceVideoSentinel;
        }
        savePersistentSettings();
        logApp(std::string("[App] Audio playback toggled -> ") + (settings_.audioPlaybackEnabled ? "enabled" : "disabled"));
        applyAudioPlaybackSetting();
        return;
    }
    if (commandId == kCommandToggleMicrophoneCapture)
    {
        settings_.microphoneCaptureEnabled = !settings_.microphoneCaptureEnabled;
        savePersistentSettings();
        logApp(std::string("[App] Microphone capture toggled -> ") + (settings_.microphoneCaptureEnabled ? "enabled" : "disabled"));
        applyMicrophoneCaptureSetting();
        return;
    }
    if (commandId == kCommandToggleInputCapture)
    {
        settings_.inputCaptureEnabled = !settings_.inputCaptureEnabled;
        savePersistentSettings();
        logApp(std::string("[App] Input capture toggled -> ") + (settings_.inputCaptureEnabled ? "enabled" : "disabled"));
        applyInputCaptureSetting();
        return;
    }
    if (auto it = videoCommandMap_.find(commandId); it != videoCommandMap_.end())
    {
        if (settings_.videoDeviceMoniker != it->second)
        {
            settings_.videoDeviceMoniker = it->second;
            savePersistentSettings();
            logApp(std::string("[App] Selected video capture device: ") + settings_.videoDeviceMoniker);
            restartVideoCapture();
            if (settings_.audioDeviceMoniker == kAudioSourceVideoSentinel && settings_.audioPlaybackEnabled)
            {
                applyAudioPlaybackSetting();
            }
        }
        return;
    }

    if (auto it = audioCommandMap_.find(commandId); it != audioCommandMap_.end())
    {
        if (settings_.audioDeviceMoniker != it->second)
        {
            settings_.audioDeviceMoniker = it->second;
            savePersistentSettings();
        }
        const std::string logLabel = settings_.audioDeviceMoniker == kAudioSourceVideoSentinel ? std::string("video source audio") : settings_.audioDeviceMoniker;
        logApp(std::string("[App] Selected audio capture device: ") + logLabel);
        applyAudioPlaybackSetting();
        return;
    }

    if (auto it = microphoneCommandMap_.find(commandId); it != microphoneCommandMap_.end())
    {
        if (settings_.microphoneDeviceId != it->second)
        {
            settings_.microphoneDeviceId = it->second;
            savePersistentSettings();
            logApp(std::string("[App] Selected microphone device: ") + settings_.microphoneDeviceId);
            if (settings_.microphoneCaptureEnabled)
            {
                applyMicrophoneCaptureSetting();
            }
        }
        return;
    }

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

void Application::updateInputCaptureBounds()
{
    if (!hwnd_ || !IsWindowVisible(hwnd_))
    {
        inputCaptureManager_.setCaptureRegion(RECT{}, false);
        return;
    }

    RECT client{};
    if (!GetClientRect(hwnd_, &client))
    {
        inputCaptureManager_.setCaptureRegion(RECT{}, false);
        return;
    }

    POINT topLeft{client.left, client.top};
    POINT bottomRight{client.right, client.bottom};
    ClientToScreen(hwnd_, &topLeft);
    ClientToScreen(hwnd_, &bottomRight);

    RECT screenRect{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};

    if (IsIconic(hwnd_) || GetForegroundWindow() != hwnd_)
    {
        inputCaptureManager_.setCaptureRegion(screenRect, false);
    }
    else
    {
        inputCaptureManager_.setCaptureRegion(screenRect, screenRect.right > screenRect.left && screenRect.bottom > screenRect.top);
    }
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
