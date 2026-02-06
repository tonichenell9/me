import SwiftUI

/// Main entry point for the Notes iPad App.
/// Uses the @main attribute to mark this as the app's entry point.
@main
struct NotesApp: App {
    /// Shared view model injected into the environment for all views.
    @StateObject private var viewModel = NotesViewModel()

    var body: some Scene {
        WindowGroup {
            ContentView()
                .environmentObject(viewModel)
        }
        .commands {
            // Custom keyboard shortcuts for iPad with external keyboard
            CommandGroup(replacing: .newItem) {
                Button("New Note") {
                    viewModel.createNote()
                }
                .keyboardShortcut("n", modifiers: .command)
            }

            CommandMenu("Notes") {
                Button("Toggle Favorites") {
                    if let note = viewModel.selectedNote {
                        viewModel.toggleFavorite(note)
                    }
                }
                .keyboardShortcut("f", modifiers: [.command, .shift])

                Divider()

                Button("Delete Note") {
                    if let note = viewModel.selectedNote {
                        viewModel.deleteNote(note)
                    }
                }
                .keyboardShortcut(.delete, modifiers: .command)
            }
        }
    }
}
