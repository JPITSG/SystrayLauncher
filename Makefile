# NextcloudLauncher Makefile
# Cross-compile for Windows using MinGW-w64

CC = x86_64-w64-mingw32-gcc
WINDRES = x86_64-w64-mingw32-windres

TARGET = NextcloudLauncher.exe
SOURCES = NextcloudLauncher.c
RESOURCES = resource.rc

OBJ = main.o resource.o

CFLAGS = -mwindows
LDFLAGS = -mwindows
LIBS = WebView2Loader.dll.lib -lole32 -lshell32 -lshlwapi -luuid -luser32 -lgdi32 -lcomctl32

.PHONY: all clean

all: $(TARGET)

$(TARGET): $(OBJ)
	@echo "Linking executable..."
	$(CC) -o $@ $(OBJ) $(LDFLAGS) $(LIBS)

main.o: $(SOURCES)
	@echo "Compiling $(SOURCES)..."
	$(CC) -c $< -o $@ $(CFLAGS)

resource.o: $(RESOURCES)
	@echo "Compiling resources..."
	$(WINDRES) $< -o $@

clean:
	rm -f $(OBJ) $(TARGET)
