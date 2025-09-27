#pragma once

#include <Windows.h>

#include <cstdint>

#include <d3d12.h>
#include <dxgi1_6.h>
#include <wrl/client.h>

class D3DRenderer {
public:
    D3DRenderer() = default;
    ~D3DRenderer();

    bool initialize(HWND hwnd, bool enableDebug = false);
    void shutdown();

    void onResize(UINT width, UINT height);

    void uploadFrame(const void* data,
                     std::uint32_t stride,
                     std::uint32_t width,
                     std::uint32_t height);

    void render();

    void setDebugGradient(bool enable);
    [[nodiscard]] bool debugGradientEnabled() const { return debugGradient_; }

private:
    struct FrameContext {
        Microsoft::WRL::ComPtr<ID3D12CommandAllocator> commandAllocator;
        std::uint64_t fenceValue = 0;
    };

    bool createDevice(HWND hwnd, bool enableDebug);
    bool createSwapChain(HWND hwnd);
    bool createPipelineResources();
    bool createRenderTargets();
    void destroyRenderTarget();
    bool ensureFrameResources(std::uint32_t width,
                              std::uint32_t height,
                              std::uint32_t stride);
    void destroyFrameResources();
    void waitForFrame(FrameContext& frameContext);
    void waitForGpu();

    static constexpr std::uint32_t kFrameCount = 2;

    Microsoft::WRL::ComPtr<ID3D12Device> device_;
    Microsoft::WRL::ComPtr<ID3D12CommandQueue> commandQueue_;
    Microsoft::WRL::ComPtr<IDXGISwapChain4> swapChain_;

    Microsoft::WRL::ComPtr<ID3D12DescriptorHeap> rtvHeap_;
    Microsoft::WRL::ComPtr<ID3D12DescriptorHeap> srvHeap_;
    Microsoft::WRL::ComPtr<ID3D12DescriptorHeap> samplerHeap_;

    Microsoft::WRL::ComPtr<ID3D12Resource> renderTargets_[kFrameCount];
    FrameContext frameContexts_[kFrameCount];

    Microsoft::WRL::ComPtr<ID3D12GraphicsCommandList> commandList_;
    Microsoft::WRL::ComPtr<ID3D12Fence> fence_;
    HANDLE fenceEvent_ = nullptr;
    std::uint64_t fenceValue_ = 1;

    Microsoft::WRL::ComPtr<ID3D12RootSignature> rootSignature_;
    Microsoft::WRL::ComPtr<ID3D12PipelineState> pipelineState_;
    Microsoft::WRL::ComPtr<ID3D12PipelineState> pipelineStateGradient_;

    Microsoft::WRL::ComPtr<ID3D12Resource> vertexBuffer_;
    Microsoft::WRL::ComPtr<ID3D12Resource> indexBuffer_;
    D3D12_VERTEX_BUFFER_VIEW vertexBufferView_{};
    D3D12_INDEX_BUFFER_VIEW indexBufferView_{};

    struct UploadResource {
        Microsoft::WRL::ComPtr<ID3D12Resource> resource;
        D3D12_PLACED_SUBRESOURCE_FOOTPRINT layout{};
        std::uint64_t sizeBytes = 0;
        std::uint8_t* cpuAddress = nullptr;
    };

    Microsoft::WRL::ComPtr<ID3D12Resource> frameTexture_;
    UploadResource frameUploads_[kFrameCount];
    bool pendingUpload_[kFrameCount] = {};

    D3D12_CPU_DESCRIPTOR_HANDLE rtvHandleStart_{};
    D3D12_CPU_DESCRIPTOR_HANDLE srvHandleCpu_{};
    D3D12_GPU_DESCRIPTOR_HANDLE srvHandleGpu_{};
    D3D12_GPU_DESCRIPTOR_HANDLE samplerHandleGpu_{};

    UINT rtvDescriptorSize_ = 0;
    D3D12_VIEWPORT viewport_{};
    D3D12_RECT scissorRect_{};

    UINT frameWidth_ = 0;
    UINT frameHeight_ = 0;
    UINT frameStride_ = 0;
    UINT backBufferWidth_ = 0;
    UINT backBufferHeight_ = 0;

    HANDLE frameLatencyWaitableObject_ = nullptr;
    bool allowTearing_ = false;
    bool debugGradient_ = false;
    bool loggedGpuPixels_ = false;
    bool debugLayerEnabled_ = false;

    void updateViewport(UINT width, UINT height);
};
