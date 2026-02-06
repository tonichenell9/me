#!/bin/bash
# =============================================================================
# Report Viewer iPad App - Xcode Project Setup Script
# =============================================================================
# This script creates an Xcode project for the Report Viewer iPad app.
# It uses xcodegen (if available) or provides manual setup instructions.
#
# Prerequisites:
#   - macOS with Xcode installed
#   - Optional: xcodegen (brew install xcodegen)
#
# Usage:
#   chmod +x setup_xcode_project.sh
#   ./setup_xcode_project.sh
# =============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")/ReportViewer"
XCODEGEN_SPEC="$SCRIPT_DIR/project.yml"

echo "============================================"
echo " Report Viewer - iPad App Setup"
echo "============================================"
echo ""
echo "Project directory: $PROJECT_DIR"
echo ""

# Check if we're on macOS
if [[ "$(uname)" != "Darwin" ]]; then
    echo "WARNING: This script is designed for macOS."
    echo "You can still use the Swift source files on any platform,"
    echo "but Xcode project generation requires macOS."
    echo ""
    echo "To use on macOS:"
    echo "  1. Copy the iPadApp/ directory to your Mac"
    echo "  2. Run this script from Terminal"
    echo ""
fi

# Check for xcodegen
if command -v xcodegen &> /dev/null; then
    echo "Found xcodegen. Generating Xcode project..."
    echo ""

    # Create xcodegen spec if it doesn't exist
    if [ ! -f "$XCODEGEN_SPEC" ]; then
        cat > "$XCODEGEN_SPEC" << 'XCODEGEN_EOF'
name: ReportViewer
options:
  bundleIdPrefix: com.reportviewer
  deploymentTarget:
    iPad: "16.0"
  xcodeVersion: "15.0"
  generateEmptyDirectories: true
  groupSortPosition: top

settings:
  base:
    SWIFT_VERSION: "5.9"
    TARGETED_DEVICE_FAMILY: "2"
    SUPPORTS_MAC_DESIGNED_FOR_IPHONE_IPAD: true
    INFOPLIST_FILE: ReportViewer/Info.plist
    ASSETCATALOG_COMPILER_APPICON_NAME: AppIcon
    ASSETCATALOG_COMPILER_GLOBAL_ACCENT_COLOR_NAME: AccentColor

targets:
  ReportViewer:
    type: application
    platform: iOS
    deploymentTarget: "16.0"
    sources:
      - path: ReportViewer/Sources
      - path: ReportViewer/Assets.xcassets
      - path: ReportViewer/Info.plist
        buildPhase: none
    settings:
      base:
        PRODUCT_BUNDLE_IDENTIFIER: com.reportviewer.app
        PRODUCT_NAME: "Report Viewer"
        MARKETING_VERSION: "1.0.0"
        CURRENT_PROJECT_VERSION: "1"
        DEVELOPMENT_TEAM: ""
        CODE_SIGN_STYLE: Automatic
        TARGETED_DEVICE_FAMILY: "2"
    info:
      path: ReportViewer/Info.plist
      properties:
        UILaunchScreen:
          UIColorName: AccentColor
        UIApplicationSceneManifest:
          UIApplicationSupportsMultipleScenes: true
        UISupportedInterfaceOrientations~ipad:
          - UIInterfaceOrientationPortrait
          - UIInterfaceOrientationPortraitUpsideDown
          - UIInterfaceOrientationLandscapeLeft
          - UIInterfaceOrientationLandscapeRight
XCODEGEN_EOF
        echo "Created xcodegen spec at: $XCODEGEN_SPEC"
    fi

    cd "$(dirname "$PROJECT_DIR")"
    xcodegen generate --spec "$XCODEGEN_SPEC"
    echo ""
    echo "Xcode project generated successfully!"
    echo ""

    # Open in Xcode
    read -p "Open in Xcode? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        open "ReportViewer.xcodeproj"
    fi
else
    echo "xcodegen not found. Providing manual setup instructions."
    echo ""
    echo "============================================"
    echo " Manual Xcode Project Setup"
    echo "============================================"
    echo ""
    echo "Option 1: Install xcodegen and re-run this script"
    echo "  brew install xcodegen"
    echo "  ./setup_xcode_project.sh"
    echo ""
    echo "Option 2: Create the project manually in Xcode"
    echo ""
    echo "  1. Open Xcode"
    echo "  2. File > New > Project"
    echo "  3. Choose 'App' under iOS"
    echo "  4. Configure:"
    echo "     - Product Name: ReportViewer"
    echo "     - Interface: SwiftUI"
    echo "     - Language: Swift"
    echo "     - Target: iPad"
    echo "  5. Save the project"
    echo "  6. Delete the auto-generated ContentView.swift"
    echo "  7. Drag the following folders into the project:"
    echo "     - Sources/ (with all subfolders)"
    echo "     - Assets.xcassets"
    echo "     - Preview Content/"
    echo "  8. Replace the auto-generated Info.plist with the one provided"
    echo "  9. Build and run on iPad simulator"
    echo ""
    echo "Option 3: Use Swift Playgrounds on iPad"
    echo ""
    echo "  1. Open Swift Playgrounds on your iPad"
    echo "  2. Create a new App project"
    echo "  3. Copy the Swift files from Sources/ into the project"
    echo "  4. The app will build and run directly on your iPad"
    echo ""

    # Create xcodegen spec for future use
    cat > "$XCODEGEN_SPEC" << 'XCODEGEN_EOF'
name: ReportViewer
options:
  bundleIdPrefix: com.reportviewer
  deploymentTarget:
    iPad: "16.0"
  xcodeVersion: "15.0"
  generateEmptyDirectories: true
  groupSortPosition: top

settings:
  base:
    SWIFT_VERSION: "5.9"
    TARGETED_DEVICE_FAMILY: "2"
    SUPPORTS_MAC_DESIGNED_FOR_IPHONE_IPAD: true
    INFOPLIST_FILE: ReportViewer/Info.plist
    ASSETCATALOG_COMPILER_APPICON_NAME: AppIcon
    ASSETCATALOG_COMPILER_GLOBAL_ACCENT_COLOR_NAME: AccentColor

targets:
  ReportViewer:
    type: application
    platform: iOS
    deploymentTarget: "16.0"
    sources:
      - path: ReportViewer/Sources
      - path: ReportViewer/Assets.xcassets
      - path: ReportViewer/Info.plist
        buildPhase: none
    settings:
      base:
        PRODUCT_BUNDLE_IDENTIFIER: com.reportviewer.app
        PRODUCT_NAME: "Report Viewer"
        MARKETING_VERSION: "1.0.0"
        CURRENT_PROJECT_VERSION: "1"
        DEVELOPMENT_TEAM: ""
        CODE_SIGN_STYLE: Automatic
        TARGETED_DEVICE_FAMILY: "2"
    info:
      path: ReportViewer/Info.plist
      properties:
        UILaunchScreen:
          UIColorName: AccentColor
        UIApplicationSceneManifest:
          UIApplicationSupportsMultipleScenes: true
        UISupportedInterfaceOrientations~ipad:
          - UIInterfaceOrientationPortrait
          - UIInterfaceOrientationPortraitUpsideDown
          - UIInterfaceOrientationLandscapeLeft
          - UIInterfaceOrientationLandscapeRight
XCODEGEN_EOF
    echo "Saved xcodegen spec to: $XCODEGEN_SPEC"
    echo "(Install xcodegen later with: brew install xcodegen)"
fi

echo ""
echo "============================================"
echo " Setup Complete"
echo "============================================"
echo ""
echo "Source files are located at:"
echo "  $PROJECT_DIR/Sources/"
echo ""
echo "Structure:"
echo "  Sources/"
echo "    App/          - App entry point and state management"
echo "    Models/       - Data models (Report, ServerConfig)"
echo "    Views/        - SwiftUI views (optimized for iPad)"
echo "    ViewModels/   - View models for data binding"
echo "    Services/     - Network and notification services"
echo "    Extensions/   - Swift type extensions"
echo ""
