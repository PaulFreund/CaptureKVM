#include "Application.hpp"

#include <Windows.h>

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    Application app;
    return app.run();
}
