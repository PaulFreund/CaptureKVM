#pragma once

#include <cstdint>
#include <functional>
#include <memory>
#include <string>

struct DirectShowCaptureImpl;

class DirectShowCapture {
public:
    enum class PixelFormat {
        BGRA8,
    };

    struct Frame {
        std::uint32_t width{};
        std::uint32_t height{};
        std::uint32_t stride{};
        std::uint64_t timestamp100ns{};
        const std::uint8_t* data{};
        std::size_t dataSize{};
        bool bottomUp{};
    };

    using FrameHandler = std::function<void(const Frame&)>;

    struct Options {
        std::string deviceMoniker;
        bool enableAudio = false;
    };

    DirectShowCapture();
    ~DirectShowCapture();

    void start(FrameHandler handler, const Options& options = {});
    void stop();

    [[nodiscard]] std::string consumeLastError();
    [[nodiscard]] std::string currentDeviceFriendlyName() const;

    DirectShowCapture(const DirectShowCapture&) = delete;
    DirectShowCapture& operator=(const DirectShowCapture&) = delete;

private:
    std::unique_ptr<DirectShowCaptureImpl> impl_;
};
