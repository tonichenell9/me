# Report Viewer - iPad App

A SwiftUI iPad app that serves as a companion to the **Large Deal Report Automation** system. View, manage, and share your daily reports directly on your iPad with a modern, touch-optimized interface.

## Features

- **iPad-Optimized Layout**: Three-column NavigationSplitView with sidebar, list, and detail panels
- **Report Browsing**: Browse reports by category (Today, This Week, Favorites, All)
- **Report Details**: View detailed report metadata, worksheet information, and file properties
- **Search & Filter**: Full-text search across report titles and summaries
- **Pull to Refresh**: Swipe down to sync with the automation server
- **Favorites**: Star important reports for quick access
- **Settings**: Configure server connection, notifications, and preferences
- **Dark Mode**: Full support for light and dark appearance
- **Notifications**: Get notified when new reports are ready

## Architecture

The app follows the **MVVM (Model-View-ViewModel)** pattern with SwiftUI:

```
iPadApp/
├── Scripts/
│   ├── setup_xcode_project.sh    # Xcode project setup script
│   └── project.yml               # xcodegen specification
├── ReportViewer/
│   ├── Sources/
│   │   ├── App/
│   │   │   ├── ReportViewerApp.swift   # App entry point (@main)
│   │   │   └── AppState.swift          # Global app state (ObservableObject)
│   │   ├── Models/
│   │   │   ├── Report.swift            # Report data model
│   │   │   └── ServerConfig.swift      # Server configuration model
│   │   ├── Views/
│   │   │   ├── ContentView.swift       # Root NavigationSplitView
│   │   │   ├── SidebarView.swift       # Sidebar with categories & stats
│   │   │   ├── ReportListView.swift    # Report list with search/filter
│   │   │   ├── ReportRowView.swift     # Individual report row
│   │   │   ├── ReportDetailView.swift  # Full report detail view
│   │   │   ├── EmptyDetailView.swift   # Placeholder when no report selected
│   │   │   └── SettingsView.swift      # App settings & configuration
│   │   ├── ViewModels/
│   │   │   └── ReportListViewModel.swift  # List sorting & filtering logic
│   │   ├── Services/
│   │   │   ├── ReportService.swift        # Network & data service
│   │   │   └── NotificationService.swift  # Local notification management
│   │   └── Extensions/
│   │       ├── Date+Extensions.swift      # Date utility extensions
│   │       └── Color+Extensions.swift     # Brand color extensions
│   ├── Assets.xcassets/           # App icons and colors
│   ├── Preview Content/           # SwiftUI preview assets
│   └── Info.plist                 # App configuration
└── README.md                      # This file
```

## Requirements

- **macOS** 13+ with **Xcode 15+** (for building)
- **iPadOS 16.0+** (deployment target)
- **Swift 5.9+**

## Quick Start

### Option 1: Using xcodegen (Recommended)

1. Install xcodegen:
   ```bash
   brew install xcodegen
   ```

2. Run the setup script:
   ```bash
   cd iPadApp
   chmod +x Scripts/setup_xcode_project.sh
   ./Scripts/setup_xcode_project.sh
   ```

3. Open the generated `ReportViewer.xcodeproj` in Xcode

4. Select an iPad simulator and press **Cmd+R** to build and run

### Option 2: Manual Xcode Setup

1. Open Xcode
2. **File > New > Project**
3. Select **App** under iOS
4. Configure:
   - **Product Name**: ReportViewer
   - **Interface**: SwiftUI
   - **Language**: Swift
   - Check **iPad** as the target device
5. Delete the auto-generated `ContentView.swift` and `ReportViewerApp.swift`
6. Drag the `ReportViewer/Sources/` folder into the Xcode project navigator
7. Drag the `ReportViewer/Assets.xcassets` folder to replace the default one
8. Copy `ReportViewer/Info.plist` into the project
9. Build and run (**Cmd+R**)

### Option 3: Swift Playgrounds on iPad

You can build and run this app directly on an iPad using **Swift Playgrounds 4+**:

1. Open **Swift Playgrounds** on your iPad
2. Tap **"+"** to create a new **App** project
3. Copy the Swift files from `Sources/` into the project
4. Tap **Run** to build and test directly on your iPad

## Server Connection

The app connects to the same infrastructure as the Python automation system. By default, it runs with sample data for demonstration.

### Connecting to Your Server

1. Open the app on your iPad
2. Navigate to **Settings** in the sidebar
3. Enter your **Server URL** (the machine running the Python automation)
4. The app expects a REST API at `{serverURL}/api/reports`

### API Endpoints Expected

The app looks for these endpoints on your server:

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/reports` | GET | List all reports |
| `/api/reports/{id}/download` | GET | Download a report file |

### Configuration Compatibility

The `ServerConfig` model in the app mirrors the Python `config.json` format, making it easy to share configuration between the automation system and the iPad app.

## Key Features Explained

### Three-Column iPad Layout

The app uses `NavigationSplitView` to provide a native iPad experience:
- **Sidebar**: Categories, quick stats, and settings access
- **Content**: Searchable, filterable report list with swipe actions
- **Detail**: Rich report information with worksheets and action buttons

### Report Management

- **Search**: Type to search across report titles and summaries
- **Filter by Category**: Today, This Week, Favorites, or All Reports
- **Swipe Actions**: Swipe right to favorite, swipe left to delete
- **Context Menus**: Long-press for additional options
- **Pull to Refresh**: Pull down the list to sync with the server

### Notifications

- **Report Ready**: Get notified when a new report is generated
- **Daily Reminder**: Optional daily reminder to check for reports
- Configure notification preferences in Settings

## Customization

### Adding New Views

Create new SwiftUI views in `Sources/Views/` following the existing pattern:

```swift
import SwiftUI

struct MyNewView: View {
    @EnvironmentObject var appState: AppState

    var body: some View {
        // Your view content
    }
}
```

### Adding New Data Models

Add models to `Sources/Models/`:

```swift
struct MyModel: Identifiable, Codable {
    let id: UUID
    var name: String
    // ...
}
```

### Extending the API

Add new methods to `ReportService` in `Sources/Services/`:

```swift
func fetchCustomData() async throws -> [CustomData] {
    // Network request implementation
}
```

## Design Decisions

- **SwiftUI-first**: Built entirely with SwiftUI for modern, declarative UI
- **iPad-optimized**: Uses NavigationSplitView, large touch targets, and multi-column layouts
- **Offline-capable**: Works with sample data when no server is configured
- **Modular**: Clean separation of concerns with MVVM architecture
- **Accessible**: Supports Dynamic Type, VoiceOver, and system appearance settings

## Troubleshooting

### Build Errors

- Ensure your Xcode version is 15.0 or later
- Clean the build folder: **Product > Clean Build Folder** (Cmd+Shift+K)
- Delete derived data if needed

### No Reports Loading

- Check your server URL in Settings
- Ensure the server is running and accessible
- The app falls back to sample data if the server is unreachable

### Preview Not Working

- Ensure preview assets are included in the project
- Try **Editor > Canvas > Refresh All Previews**

## License

This iPad app is provided as a companion to the Large Deal Report Automation system for internal use.
