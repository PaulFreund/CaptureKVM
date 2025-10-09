#include "Application.hpp"

#include <Windows.h>

namespace
{
    void enableHighDpiAwareness()
    {
#ifdef DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        const auto perMonitorV2 = DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2;
#else
        const auto perMonitorV2 = reinterpret_cast<HANDLE>(static_cast<INT_PTR>(-4));
#endif

        HMODULE user32 = GetModuleHandleW(L"user32.dll");
        if (user32)
        {
            using SetProcessDpiAwarenessContextFn = BOOL(WINAPI*)(HANDLE);
            auto setContext = reinterpret_cast<SetProcessDpiAwarenessContextFn>(
                GetProcAddress(user32, "SetProcessDpiAwarenessContext"));
            if (setContext && setContext(perMonitorV2))
            {
                return;
            }
        }

        HMODULE shcore = LoadLibraryW(L"Shcore.dll");
        if (shcore)
        {
            using SetProcessDpiAwarenessFn = HRESULT(WINAPI*)(int);
            auto setAwareness = reinterpret_cast<SetProcessDpiAwarenessFn>(
                GetProcAddress(shcore, "SetProcessDpiAwareness"));
            if (setAwareness)
            {
                constexpr int kProcessPerMonitorDpiAware = 2; // PROCESS_PER_MONITOR_DPI_AWARE
                const HRESULT result = setAwareness(kProcessPerMonitorDpiAware);
                if (SUCCEEDED(result) || result == E_ACCESSDENIED)
                {
                    FreeLibrary(shcore);
                    return;
                }
            }
            FreeLibrary(shcore);
        }

        if (user32)
        {
            using SetProcessDPIAwareFn = BOOL(WINAPI*)(void);
            auto setLegacy = reinterpret_cast<SetProcessDPIAwareFn>(
                GetProcAddress(user32, "SetProcessDPIAware"));
            if (setLegacy)
            {
                setLegacy();
            }
        }
    }
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    enableHighDpiAwareness(); // Ensure absolute mouse coordinates remain accurate (e.g. over RDP)

    Application app;
    return app.run();
}
