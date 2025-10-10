#include "OverlayUI.hpp"

#include "Application.hpp"
#include "D3DRenderer.hpp"
#include "Settings.hpp"

#include "imgui.h"
#include "backends/imgui_impl_dx12.h"
#define IMGUI_IMPL_WIN32_DISABLE_GAMEPAD
#include "backends/imgui_impl_win32.h"

#include <algorithm>
#include <cmath>
#include <string_view>

extern IMGUI_IMPL_API LRESULT ImGui_ImplWin32_WndProcHandler(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

OverlayUI::~OverlayUI()
{
    shutdown();
}

bool OverlayUI::initialize(HWND hwnd, D3DRenderer& renderer)
{
    if (initialized_)
    {
        return true;
    }

    hwnd_ = hwnd;
    renderer_ = &renderer;
    srvHeap_ = renderer.srvHeap();
    fontCpuHandle_ = renderer.imguiSrvCpuHandle();
    fontGpuHandle_ = renderer.imguiSrvGpuHandle();

    IMGUI_CHECKVERSION();
    ImGui::CreateContext();
    ImGuiIO& io = ImGui::GetIO();
    io.ConfigFlags |= ImGuiConfigFlags_NavEnableKeyboard;

    ImGui::StyleColorsDark();
    ImGuiStyle& style = ImGui::GetStyle();
    style.WindowRounding = 6.0f;
    style.FrameRounding = 4.0f;
    style.GrabRounding = 4.0f;

    ImGui_ImplWin32_Init(hwnd_);
    ImGui_ImplDX12_Init(renderer.device(), static_cast<int>(renderer.frameCount()), renderer.renderTargetFormat(),
                        srvHeap_, fontCpuHandle_, fontGpuHandle_);

    initialized_ = true;
    return true;
}

void OverlayUI::shutdown()
{
    if (!initialized_)
    {
        return;
    }

    ImGui_ImplDX12_Shutdown();
    ImGui_ImplWin32_Shutdown();
    ImGui::DestroyContext();

    initialized_ = false;
    menuVisible_ = false;
    srvHeap_ = nullptr;
    renderer_ = nullptr;
    drawData_ = nullptr;
    drawDataValid_ = false;
}

void OverlayUI::newFrame()
{
    if (!initialized_)
    {
        return;
    }

    ImGui_ImplDX12_NewFrame();
    ImGui_ImplWin32_NewFrame();
    ImGui::NewFrame();
    ImGuiIO& io = ImGui::GetIO();
    io.MouseDrawCursor = menuVisible_;

}

void OverlayUI::buildUI(Application& app)
{
    if (!initialized_ || !menuVisible_)
    {
        return;
    }

    drawMenuWindow(app);
}

void OverlayUI::endFrame()
{
    if (!initialized_)
    {
        return;
    }

    ImGui::Render();
    drawData_ = ImGui::GetDrawData();
    drawDataValid_ = (drawData_ != nullptr) && (drawData_->CmdListsCount > 0);
}

void OverlayUI::render(ID3D12GraphicsCommandList* commandList)
{
    if (!initialized_ || !drawData_)
    {
        drawDataValid_ = false;
        return;
    }

    if (drawData_->CmdListsCount > 0)
    {
        ID3D12DescriptorHeap* heaps[] = {srvHeap_};
        commandList->SetDescriptorHeaps(1, heaps);
        ImGui_ImplDX12_RenderDrawData(drawData_, commandList);
    }

    drawData_ = nullptr;
    drawDataValid_ = false;
}

bool OverlayUI::processEvent(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (!initialized_)
    {
        return false;
    }
    bool handled = ImGui_ImplWin32_WndProcHandler(hwnd, msg, wParam, lParam);
    return menuVisible_ && handled;
}

void OverlayUI::toggleMenu(Application& app)
{
    if (!initialized_)
    {
        return;
    }
    if (menuVisible_)
    {
        hideMenu(app);
    }
    else
    {
        showMenu(app);
    }
}

void OverlayUI::hideMenu(Application& app)
{
    if (!initialized_)
    {
        return;
    }
    if (!menuVisible_)
    {
        return;
    }
    menuVisible_ = false;
    drawDataValid_ = false;
    PostMessage(hwnd_, WM_INPUT_CAPTURE_UPDATE_CLIP, 1, 0);
    ImGui::GetIO().MouseDrawCursor = false;
    app.requestImmediateRender();
}

void OverlayUI::showMenu(Application& app)
{
    if (!initialized_)
    {
        return;
    }
    menuVisible_ = true;
    refreshDeviceLists(app);
    PostMessage(hwnd_, WM_INPUT_CAPTURE_UPDATE_CLIP, 0, 0);
    ImGui::GetIO().MouseDrawCursor = true;
    app.requestImmediateRender();
}

void OverlayUI::refreshDeviceLists(Application& app)
{
    videoDevices_ = enumerateVideoCaptureDevices();
    audioDevices_ = enumerateAudioCaptureDevices();
    microphoneDevices_ = enumerateMicrophoneDevices();

    bridgeDevices_.clear();
    const auto serialPorts = enumerateSerialPorts();
    for (const auto& port : serialPorts)
    {
        unsigned int suggestedBaud = 0;
        if (app.classifyBridgeDevice(port, &suggestedBaud))
        {
            BridgeOption option;
            option.port = port;
            option.suggestedBaud = suggestedBaud;
            bridgeDevices_.push_back(std::move(option));
        }
    }

    if (bridgeDevices_.size() == 1)
    {
        const BridgeOption& option = bridgeDevices_.front();
        if (app.settings().inputTargetDevice != option.port.portName)
        {
            app.selectBridgeDevice(option.port, true);
        }
    }
}

void OverlayUI::drawMenuWindow(Application& app)
{
    ImGuiIO& io = ImGui::GetIO();

    ImGui::PushStyleVar(ImGuiStyleVar_WindowRounding, 0.0f);
    ImGui::PushStyleVar(ImGuiStyleVar_WindowPadding, ImVec2(0.0f, 0.0f));
    ImGui::PushStyleColor(ImGuiCol_WindowBg, ImVec4(0.0f, 0.0f, 0.0f, 0.55f));
    ImGui::SetNextWindowPos(ImVec2(0.0f, 0.0f));
    ImGui::SetNextWindowSize(io.DisplaySize);
    ImGui::Begin("##overlay_bg", nullptr,
                 ImGuiWindowFlags_NoDecoration |
                 ImGuiWindowFlags_NoInputs |
                 ImGuiWindowFlags_NoSavedSettings |
                 ImGuiWindowFlags_NoMove |
                 ImGuiWindowFlags_NoBringToFrontOnFocus);
    ImGui::End();
    ImGui::PopStyleColor();
    ImGui::PopStyleVar(2);

    const float panelWidth = 460.0f;
    const float panelHeight = 520.0f;
    ImVec2 panelPos((io.DisplaySize.x - panelWidth) * 0.5f, (io.DisplaySize.y - panelHeight) * 0.5f);
    ImGui::SetNextWindowPos(panelPos, ImGuiCond_Always);
    ImGui::SetNextWindowSize(ImVec2(panelWidth, panelHeight));
    ImGui::PushStyleColor(ImGuiCol_WindowBg, ImVec4(0.08f, 0.08f, 0.08f, 0.94f));
    ImGui::PushStyleVar(ImGuiStyleVar_WindowRounding, 10.0f);
    ImGuiWindowFlags windowFlags = ImGuiWindowFlags_NoCollapse | ImGuiWindowFlags_NoResize | ImGuiWindowFlags_NoSavedSettings;
    if (!ImGui::Begin("CaptureKVM Settings", &menuVisible_, windowFlags))
    {
        ImGui::End();
        ImGui::PopStyleVar();
        ImGui::PopStyleColor();
        if (!menuVisible_)
        {
            hideMenu(app);
        }
        return;
    }

    ImGui::TextUnformatted("General");
    ImGui::Separator();

    bool audioPlayback = app.settings().audioPlaybackEnabled;
    if (ImGui::Checkbox("Enable Audio Playback", &audioPlayback))
    {
        app.setAudioPlaybackEnabled(audioPlayback);
    }

    bool microphoneCapture = app.settings().microphoneCaptureEnabled;
    if (ImGui::Checkbox("Enable Microphone Capture", &microphoneCapture))
    {
        app.setMicrophoneCaptureEnabled(microphoneCapture);
    }

    bool inputCapture = app.settings().inputCaptureEnabled;
    if (ImGui::Checkbox("Enable Keyboard && Mouse Capture", &inputCapture))
    {
        app.setInputCaptureEnabled(inputCapture);
    }

    ImGui::Spacing();

    ImGui::TextUnformatted("Bridge Device");
    ImGui::Separator();
    if (bridgeDevices_.empty())
    {
        ImGui::TextDisabled("No supported bridge devices detected");
    }
    else
    {
        ImGui::BeginChild("BridgeDevices", ImVec2(0.0f, 110.0f), true);
        const std::string& currentBridge = app.settings().inputTargetDevice;
        for (const auto& option : bridgeDevices_)
        {
            std::string label = option.port.friendlyName.empty() ? option.port.portName : option.port.friendlyName;
            if (label.empty())
            {
                label = option.port.portName;
            }
            if (!option.port.portName.empty())
            {
                label += " (" + option.port.portName + ")";
            }
            label += "##bridge" + option.port.portName;
            bool selected = (!currentBridge.empty() && currentBridge == option.port.portName);
            if (ImGui::Selectable(label.c_str(), selected))
            {
                app.selectBridgeDevice(option.port, false);
            }
            if (ImGui::IsItemHovered())
            {
                std::string tooltip;
                tooltip.reserve(128);
                tooltip += "Port: " + option.port.portName;
                if (!option.port.deviceDescription.empty())
                {
                    tooltip += "\nDescription: " + option.port.deviceDescription;
                }
                tooltip += "\nSuggested baud: " + std::to_string(option.suggestedBaud);
                if (!option.port.hardwareIds.empty())
                {
                    tooltip += "\nHardware IDs:";
                    for (const auto& id : option.port.hardwareIds)
                    {
                        tooltip += "\n  " + id;
                    }
                }
                ImGui::SetTooltip("%s", tooltip.c_str());
            }
        }
        ImGui::EndChild();
    }

    ImGui::Spacing();

    ImGui::TextUnformatted("Video Settings");
    ImGui::Separator();
    bool allowResizing = app.settings().videoAllowResizing;
    if (ImGui::Checkbox("Allow Resizing", &allowResizing))
    {
        app.setVideoAllowResizing(allowResizing);
    }

    static const char* aspectOptions[] = {"Stretch", "Force Aspect Ratio", "Force Capture Resolution"};
    int currentAspect = static_cast<int>(app.settings().videoAspectMode);
    if (ImGui::Combo("Aspect Mode", &currentAspect, aspectOptions, IM_ARRAYSIZE(aspectOptions)))
    {
        currentAspect = std::clamp(currentAspect, 0, 2);
        app.setVideoAspectMode(static_cast<VideoAspectMode>(currentAspect));
    }

    ImGui::Spacing();

    if (ImGui::Button("Refresh Devices"))
    {
        refreshDeviceLists(app);
    }

    ImGui::Spacing();

    const float listHeight = 130.0f;

    ImGui::TextUnformatted("Video Capture Devices");
    ImGui::BeginChild("VideoDevices", ImVec2(0.0f, listHeight), true);
    const std::string& currentVideo = app.settings().videoDeviceMoniker;
    if (videoDevices_.empty())
    {
        ImGui::TextDisabled("No video capture devices detected");
    }
    else
    {
        for (const auto& device : videoDevices_)
        {
            std::string label = !device.friendlyName.empty() ? device.friendlyName : device.monikerDisplayName;
            bool selected = (!currentVideo.empty() && currentVideo == device.monikerDisplayName);
            if (ImGui::Selectable(label.c_str(), selected))
            {
                app.selectVideoDevice(device.monikerDisplayName);
            }
        }
    }
    ImGui::EndChild();

    ImGui::Spacing();

    ImGui::TextUnformatted("Audio Capture Devices");
    ImGui::BeginChild("AudioDevices", ImVec2(0.0f, listHeight), true);
    const std::string& currentAudio = app.settings().audioDeviceMoniker;
    bool useVideoAudio = currentAudio == "@video" || currentAudio.empty();
    if (ImGui::Selectable("Use Video Source Audio", useVideoAudio))
    {
        app.selectAudioDevice("@video");
    }
    if (audioDevices_.empty())
    {
        ImGui::TextDisabled("No dedicated audio capture devices detected");
    }
    else
    {
        for (const auto& device : audioDevices_)
        {
            std::string label = !device.friendlyName.empty() ? device.friendlyName : device.monikerDisplayName;
            bool selected = (!currentAudio.empty() && currentAudio == device.monikerDisplayName);
            if (ImGui::Selectable(label.c_str(), selected))
            {
                app.selectAudioDevice(device.monikerDisplayName);
            }
        }
    }
    ImGui::EndChild();

    ImGui::Spacing();

    ImGui::TextUnformatted("Microphone Devices");
    ImGui::BeginChild("MicrophoneDevices", ImVec2(0.0f, listHeight), true);
    const std::string& currentMic = app.settings().microphoneDeviceId;
    if (microphoneDevices_.empty())
    {
        ImGui::TextDisabled("No microphone devices detected");
    }
    else
    {
        for (const auto& device : microphoneDevices_)
        {
            std::string label = !device.friendlyName.empty() ? device.friendlyName : device.endpointId;
            bool selected = (!currentMic.empty() && currentMic == device.endpointId);
            if (ImGui::Selectable(label.c_str(), selected))
            {
                app.selectMicrophoneDevice(device.endpointId);
            }
        }
    }
    ImGui::EndChild();

    if (ImGui::IsKeyReleased(ImGuiKey_Escape))
    {
        hideMenu(app);
    }

    ImGui::End();
    ImGui::PopStyleVar();
    ImGui::PopStyleColor();

    if (!menuVisible_)
    {
        hideMenu(app);
    }
}
