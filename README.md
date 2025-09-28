# CaptureKVM [![Implemented with Codex](https://img.shields.io/badge/Implemented%20with-Codex-6A5ACD?logo=openai&logoColor=white)](https://github.com/openai/codex)

Low-latency Windows viewer that pulls raw frames from an AVerMedia Live Gamer 4K (GC573) capture card using DirectShow and presents them with Direct3D 12. Any HDMI capture device that exposes a DirectShow video stream should work, but the workflow has been verified explicitly with the GC573.

```
+-----------+     USB    +---------------------------+     USB    +------------------------+
| Keyboard  |----------->|                           |----------->|    ESP32-P4 Bridge     |
+-----------+            |         Host PC           |            +-----------+------------+
                         |   +-------------------+   |                        |
+-----------+     USB    |   | CaptureKVM App    |   |                        |          USB 
| Mouse     |----------->|   +-------------------+   |                        | (Mouse/Keyboard/Mic)
+-----------+            |                           |                        |
                         |                           |            +------------------------+
+-----------+     USB    |                           |            |      Target Laptop     |
| Mic       |----------->|                           |            +-----------+------------+
+-----------+            +------------+--------------+                        |
                                      | PCIe                                  |
                                      v                                       |
                          +-------------------------------+         HDMI      |
                          |      Capture Card (GC573)     |<------------------+
                          +-------------------------------+    (Video/Audio)
```

## Demo

[![YouTube demo video](https://img.youtube.com/vi/n5pnRM8FXPs/0.jpg)](https://www.youtube.com/watch?v=n5pnRM8FXPs)

## Prerequisites

- Windows 10/11 with the GC573 installed and configured.
- CMake 3.20 or newer.
- Microsoft Visual Studio 2019/2022 with the MSVC x64 toolchain.
- Official AVerMedia GC573 driver. For best presentation latency keep the window on the NVIDIA RTX 4070 driving your display.
- A [CaptureKVMBridge](https://github.com/PaulFreund/CaptureKVMBridge) device

## Building

1. Configure and build (Visual Studio generator shown, adjust to taste):
   ```powershell
   cmake -S . -B build -G "Visual Studio 17 2022" -A x64
   cmake --build build --config Release
   ```
2. Run the viewer from the generated `Release` (or `Debug`) output directory.
   - Audio playback can be toggled at runtime from the settings menu (Page Up + Home). The legacy `--enable-audio` flag still forces audio on at launch if you prefer.

## Runtime behaviour

- The app enumerates the GC573 through DirectShow, builds a graph with the Sample Grabber filter, and streams 32-bit BGRA frames into the renderer without extra buffering.
- Frames are uploaded into a D3D12 texture and drawn over a flip-model swapchain to minimise the presentation queue.
- Press Page Up + Home at any time to open an in-window settings menu. Device choices and feature toggles persist in `settings.json` beside the executable.
- Optional audio playback, microphone capture, and keyboard/mouse streaming can all be toggled live without breaking the video pipeline.
- TLV reports are pushed over the auto-detected COM port exposed by the `USB JTAG/serial debug unit` bridge (VID 303A, PID 1001), so the viewer immediately reconnects whenever the adapter is attached.
- CPU-side double buffering keeps the capture callback decoupled from the render loop while maintaining low latency.
- Keyboard, mouse, and microphone data are streamed as TLV packets over the configured serial link by a dedicated worker thread so the video path stays contention-free.

## Notes

- The project links against `d3d11`, `dxgi`, `d3dcompiler`, `quartz`, `strmiids`, `ole32`, and `oleaut32`; make sure those libraries are available in your Visual Studio environment.
- If you need to support additional pixel formats, adjust the `SampleGrabber` configuration in `src/DirectShowCapture.cpp` and the upload path in `src/D3DRenderer.cpp` accordingly.
- Close any other capture applications (e.g. RECentral, OBS) before launching the viewer to avoid exclusive-device conflicts.

## Serial TLV Protocol

Input, pointer, and microphone data are serialized onto the configured COM port using a compact type-length-value envelope. The length field is big-endian so the receiver can determine the payload size without peeking ahead.

- **Envelope**: 2 byte sync word (`0xD5 0xAA`), 1 byte `type`, 2 bytes `length` (big-endian payload size), followed by `length` payload bytes. The sync word uses an alternating bit pattern that is unlikely to appear in audio or HID payloads, which simplifies resynchronisation on noisy links.
- **Type 0x01 – Keyboard Report**
  - 8-byte HID boot keyboard report
    - Byte 0: modifier bits (`0x01` LCtrl, `0x02` LShift, `0x04` LAlt, `0x08` LGUI, `0x10` RCtrl, `0x20` RShift, `0x40` RAlt, `0x80` RGUI)
    - Byte 1: reserved (always 0)
    - Bytes 2-7: up to six HID usage codes for the currently pressed keys (unused slots are 0)
    - If more than six keys are held, each key slot is set to `0x01` (error roll-over)
- **Type 0x02 – Mouse Report**
  - 5 bytes: `buttons`, `dx`, `dy`, `wheel`, `pan`
    - `buttons`: bitfield (`0x01` L, `0x02` R, `0x04` M, `0x08` X1, `0x10` X2)
    - `dx`, `dy`: signed 8-bit relative motion (clamped to ±127)
    - `wheel`: signed 8-bit vertical wheel steps (multiples of 1 detent)
    - `pan`: signed 8-bit horizontal wheel steps
- **Type 0x04 – Mouse Absolute Report**
  - 7 bytes: `buttons`, `x_hi`, `x_lo`, `y_hi`, `y_lo`, `wheel`, `pan`
    - `buttons`: bitfield (`0x01` L, `0x02` R, `0x04` M, `0x08` X1, `0x10` X2)
    - `x`, `y`: unsigned 16-bit absolute coordinates spanning the full virtual desktop
    - `wheel`, `pan`: signed 8-bit wheel deltas matching the relative report
- **Type 0x03 – Microphone Samples**
  - Little-endian 16-bit PCM samples captured at 48 kHz. Each packet carries up to 65,535 bytes of contiguous sample data; larger buffers are automatically split across multiple TLVs.

Packets are queued and written from a dedicated worker thread so the TLV stream never blocks the capture/render loop. Receivers should scan for the `0xD5 0xAA` sync word, read the following type/length header, and then consume the payload according to the type.

## License

CaptureKVM is released under the [MIT License](LICENSE).
