# NextcloudLauncher Makefile
# Cross-compile for Windows using MinGW-w64

CC = x86_64-w64-mingw32-gcc
WINDRES = x86_64-w64-mingw32-windres

TARGET = NextcloudLauncher.exe
SOURCES = NextcloudLauncher.c
RESOURCES = resource.rc

OBJ = main.o resource.o

# WebView2 SDK paths
SDK_DIR = webview2_sdk
SDK_URL = https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2
SDK_INCLUDE = $(SDK_DIR)/build/native/include
SDK_DLL = $(SDK_DIR)/build/native/x64/WebView2Loader.dll
SDK_LIB = $(SDK_DIR)/build/native/x64/WebView2Loader.dll.lib

# Icon sizes for system tray (16=100% DPI, 24=150%, 32=200%, 48=large, 256=hi-res)
ICON_SIZES = 16 24 32 48 256

CFLAGS = -mwindows -isystem $(SDK_INCLUDE) -I.
LDFLAGS = -mwindows
LIBS = $(SDK_LIB) -lole32 -lshell32 -lshlwapi -luuid -luser32 -lgdi32 -lcomctl32

.PHONY: all clean deps check-deps dist icon

all: check-deps $(TARGET)

$(TARGET): $(OBJ)
	@echo "Linking executable..."
	$(CC) -o $@ $(OBJ) $(LDFLAGS) $(LIBS)
	@rm -f $(OBJ)
	@cp $(SDK_DLL) .
	@echo "Build complete: $(TARGET) + WebView2Loader.dll"

main.o: $(SOURCES)
	@echo "Compiling $(SOURCES)..."
	$(CC) -c $< -o $@ $(CFLAGS)

resource.o: $(RESOURCES) icon.ico
	@echo "Compiling resources..."
	$(WINDRES) $< -o $@

# Generate multi-sized icon.ico from icon.svg using ImageMagick
icon:
	@if [ -f "icon.svg" ]; then \
		echo "Generating icon.ico from icon.svg..."; \
		convert -background transparent icon.svg \
			$(foreach size,$(ICON_SIZES),\( -clone 0 -resize $(size)x$(size) \)) \
			-delete 0 -alpha on icon.ico; \
		echo "Created icon.ico with sizes: $(ICON_SIZES)"; \
	else \
		echo "No icon.svg found, skipping icon generation."; \
	fi

# Download and extract WebView2 SDK
deps: webview2.nupkg
	@echo "Extracting WebView2 SDK..."
	unzip -o webview2.nupkg -d $(SDK_DIR)
	@echo "WebView2 SDK ready."

webview2.nupkg:
	@echo "Downloading WebView2 SDK..."
	curl -L -o $@ $(SDK_URL)

# Check if dependencies exist before building
check-deps:
	@if [ ! -f "$(SDK_LIB)" ] || [ ! -d "$(SDK_INCLUDE)" ]; then \
		echo "Error: WebView2 SDK not found. Run 'make deps' first."; \
		exit 1; \
	fi

# Create distribution folder with exe and required DLL
dist: all
	@echo "Creating distribution..."
	mkdir -p dist
	cp $(TARGET) dist/
	cp $(SDK_DLL) dist/
	@echo "Distribution ready in dist/"

clean:
	rm -f $(OBJ) $(TARGET) WebView2Loader.dll

clean-all: clean
	rm -rf $(SDK_DIR) webview2.nupkg dist
