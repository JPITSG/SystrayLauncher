# SystrayLauncher Makefile
# Cross-compile for Windows using MinGW-w64

CC = x86_64-w64-mingw32-gcc
WINDRES = x86_64-w64-mingw32-windres

RELEASE_DIR = release
TARGET = $(RELEASE_DIR)/SystrayLauncher.exe
SOURCES = SystrayLauncher.c
RESOURCES = resource.rc

OBJ = main.o resource.o

# WebView2 SDK paths
SDK_DIR = webview2_sdk
SDK_URL = https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2
SDK_INCLUDE = $(SDK_DIR)/build/native/include

CFLAGS = -mwindows -isystem $(SDK_INCLUDE) -I.
LDFLAGS = -mwindows
LIBS = -lole32 -lshell32 -lshlwapi -luuid -luser32 -lgdi32 -lcomctl32

.PHONY: all clean deps check-deps

all: check-deps $(TARGET)

$(TARGET): $(OBJ) | $(RELEASE_DIR)
	@echo "Linking executable..."
	$(CC) -o $@ $(OBJ) $(LDFLAGS) $(LIBS)
	@rm -f $(OBJ)
	@echo "Build complete: $(TARGET)"

$(RELEASE_DIR):
	@mkdir -p $(RELEASE_DIR)

main.o: $(SOURCES)
	@echo "Compiling $(SOURCES)..."
	$(CC) -c $< -o $@ $(CFLAGS)

resource.o: $(RESOURCES) resource.h assets/icon.ico assets/dist/index.html assets/WebView2Loader.dll
	@echo "Compiling resources..."
	$(WINDRES) $< -o $@

# Build frontend assets
assets/dist/index.html: $(wildcard assets/src/**/*.tsx assets/src/**/*.ts assets/src/**/*.css assets/index.html)
	@echo "Building frontend assets..."
	cd assets && npm install && npm run build

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
	@if [ ! -d "$(SDK_INCLUDE)" ]; then \
		echo "Error: WebView2 SDK not found. Run 'make deps' first."; \
		exit 1; \
	fi

clean:
	rm -f $(OBJ) $(TARGET)
	rm -rf assets/dist assets/node_modules

clean-release:
	rm -rf $(RELEASE_DIR)

clean-all: clean clean-release
	rm -rf $(SDK_DIR) webview2.nupkg
