# iPad Notes App - SwiftUI Project

A complete, modern iPad app built with SwiftUI, featuring a three-column navigation layout, full note management, and iPad-optimized interactions.

## Overview

This project demonstrates how to build a production-quality iPad app using SwiftUI. It implements a **Notes** application with features that showcase iPad-specific design patterns and capabilities.

## Features

- **Three-Column Navigation** - Uses `NavigationSplitView` with sidebar, list, and detail columns optimized for iPad
- **Category Organization** - Notes can be organized into categories (General, Work, Personal, Ideas, Archive)
- **Favorites** - Quick-access favorites filter with star toggle
- **Full-Text Search** - Search across titles, content, and tags using `.searchable()`
- **Auto-Save** - Notes save automatically with a debounce timer (0.5s delay)
- **JSON Persistence** - Notes are saved to the Documents directory as JSON
- **Keyboard Shortcuts** - Full support for iPad external keyboards:
  - `Cmd+N` - New note
  - `Cmd+Shift+F` - Toggle favorite
  - `Cmd+Delete` - Delete note
- **Swipe Actions** - Swipe right to favorite, swipe left to delete
- **Context Menus** - Long-press notes for quick actions
- **Note Inspector** - Detailed metadata view (word count, character count, dates)
- **Tag System** - Comma-separated tags for flexible organization
- **Sort Options** - Multiple sort orders (date modified, date created, title)
- **Adaptive Layout** - Works in all iPad orientations and multitasking modes
- **Dark Mode** - Full dark mode support with custom accent color

## Project Structure

```
iPadApp/
├── README.md                              # This file
└── NotesApp/
    ├── Sources/
    │   ├── App/
    │   │   └── NotesApp.swift             # App entry point, keyboard commands
    │   ├── Models/
    │   │   ├── Note.swift                 # Note data model with sample data
    │   │   └── Category.swift             # Category enum and SidebarFilter
    │   ├── Views/
    │   │   ├── ContentView.swift          # Root NavigationSplitView
    │   │   ├── SidebarView.swift          # Category sidebar with badges
    │   │   ├── NoteListView.swift         # Filtered note list with swipe/context
    │   │   └── NoteEditorView.swift       # Full note editor with inspector
    │   ├── ViewModels/
    │   │   └── NotesViewModel.swift       # Central state management & persistence
    │   └── Utilities/
    │       ├── FlowLayout.swift           # Custom SwiftUI Layout for tags
    │       └── Extensions.swift           # Color, View, Date, String helpers
    ├── Resources/
    │   └── Assets.xcassets/               # App icon, accent color
    └── Preview Content/                   # SwiftUI preview assets
```

## Requirements

- **macOS** 13.0+ (for development in Xcode)
- **Xcode** 15.0+
- **iOS/iPadOS** 17.0+ deployment target
- **Swift** 5.9+

## Getting Started

### Step 1: Create the Xcode Project

1. Open **Xcode** and select **File > New > Project**
2. Choose **iOS > App**
3. Configure the project:
   - **Product Name**: `NotesApp`
   - **Team**: Select your development team
   - **Organization Identifier**: e.g., `com.yourname`
   - **Interface**: SwiftUI
   - **Language**: Swift
   - **Storage**: None
4. Click **Create** and choose a save location

### Step 2: Replace the Generated Files

1. Delete the auto-generated `ContentView.swift` and `NotesAppApp.swift` files in Xcode
2. Copy the source files from this project into your Xcode project:
   - Drag the contents of `NotesApp/Sources/` into your Xcode project navigator
   - Drag the contents of `NotesApp/Resources/Assets.xcassets/` into your existing Assets catalog
3. Make sure all `.swift` files are added to the app target

### Step 3: Configure the Project

1. Select the project in the navigator
2. Under **General > Supported Destinations**, ensure **iPad** is checked
3. Set **Minimum Deployments** to **iOS 17.0**
4. Under **Info**, set:
   - `UILaunchScreen` (empty dictionary for default launch screen)
   - `UISupportedInterfaceOrientations~ipad` to support all orientations

### Step 4: Build and Run

1. Select an iPad simulator (e.g., iPad Pro 13-inch)
2. Press `Cmd+R` to build and run
3. The app launches with sample notes pre-loaded

## Architecture

### MVVM Pattern

The app follows the **Model-View-ViewModel** architecture:

- **Models** (`Note`, `Category`) - Pure data structures conforming to `Codable`
- **Views** (`ContentView`, `SidebarView`, etc.) - SwiftUI views that render UI
- **ViewModel** (`NotesViewModel`) - `ObservableObject` managing state, business logic, and persistence

### Data Flow

```
User Action → View → ViewModel → Model → JSON File
                ↑                           |
                └───────── @Published ──────┘
```

- Views observe the `NotesViewModel` via `@EnvironmentObject`
- The ViewModel publishes changes that automatically update views
- Data is persisted to JSON in the app's Documents directory

### Key SwiftUI Patterns Used

| Pattern | Usage |
|---------|-------|
| `NavigationSplitView` | Three-column iPad layout |
| `@EnvironmentObject` | Shared ViewModel across views |
| `@StateObject` | ViewModel lifecycle management |
| `.searchable()` | Built-in search UI |
| `.commands {}` | Keyboard shortcut menus |
| `SwipeActions` | Swipe-to-delete and swipe-to-favorite |
| `contextMenu` | Long-press context menus |
| Custom `Layout` | `FlowLayout` for tag display |
| `Timer` (debounce) | Auto-save without excessive writes |

## Customization

### Adding a New Category

Edit `Category.swift` and add a new case:

```swift
enum Category: String, CaseIterable, Codable, Identifiable {
    case general = "General"
    case work = "Work"
    case personal = "Personal"
    case ideas = "Ideas"
    case archive = "Archive"
    case recipes = "Recipes"  // New category

    var iconName: String {
        switch self {
        // ... existing cases ...
        case .recipes:
            return "fork.knife"
        }
    }

    var color: Color {
        switch self {
        // ... existing cases ...
        case .recipes:
            return .pink
        }
    }
}
```

### Adding Core Data Persistence

To replace JSON persistence with Core Data:

1. Add a Core Data model (`.xcdatamodeld`) to the project
2. Create a `PersistenceController` class
3. Update `NotesViewModel` to use `NSManagedObjectContext`
4. Replace `loadNotes()`/`saveNotes()` with Core Data fetch/save operations

### Adding iCloud Sync

To sync notes across devices:

1. Enable the **iCloud** capability in Xcode
2. Add a `CloudKit` container
3. Replace JSON persistence with `NSPersistentCloudKitContainer`
4. Notes will automatically sync across all devices signed into the same iCloud account

## Keyboard Shortcuts Reference

| Shortcut | Action |
|----------|--------|
| `Cmd + N` | Create new note |
| `Cmd + Shift + F` | Toggle favorite on selected note |
| `Cmd + Delete` | Delete selected note |

## License

This project is provided as-is for educational and personal use. Feel free to modify and adapt it for your own projects.
