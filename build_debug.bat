@REM cmake -S . -B build_debug -DSERIAL_STREAMER_DEBUG=1
@REM cmake --build build_debug --config Release

cmake -S . -B build_debug -G "Visual Studio 17 2022" -A x64 -DSERIAL_STREAMER_DEBUG=1
cmake --build build_debug --config Release