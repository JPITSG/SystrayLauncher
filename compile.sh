#!/bin/bash

echo "Compiling resources..."
x86_64-w64-mingw32-windres resource.rc -o resource.o
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to compile resources"
    exit 1
fi
echo "Compiling NextcloudLauncher.c..."
x86_64-w64-mingw32-gcc -c NextcloudLauncher.c -o main.o -mwindows
if [ $? -ne 0 ]; then
    echo "ERROR: Compilation failed"
    exit 1
fi
echo "Linking executable..."
x86_64-w64-mingw32-gcc -o NextcloudLauncher.exe main.o resource.o -mwindows \
    WebView2Loader.dll.lib -lole32 -lshell32 -lshlwapi -luuid -luser32 -lgdi32 -lcomctl32
if [ $? -ne 0 ]; then
    echo "ERROR: Linking failed"
    exit 1
fi

rm -f main.o resource.o