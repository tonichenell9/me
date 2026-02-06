import SwiftUI

/// The root view of the app, providing a three-column NavigationSplitView
/// optimized for iPad's large screen.
struct ContentView: View {
    @EnvironmentObject var viewModel: NotesViewModel

    /// Controls the visibility of the navigation columns.
    @State private var columnVisibility: NavigationSplitViewVisibility = .all

    var body: some View {
        NavigationSplitView(columnVisibility: $columnVisibility) {
            SidebarView()
        } content: {
            NoteListView()
        } detail: {
            if let note = viewModel.selectedNote {
                NoteEditorView(note: note)
            } else {
                EmptyStateView()
            }
        }
        .navigationSplitViewStyle(.balanced)
        .searchable(
            text: $viewModel.searchText,
            placement: .sidebar,
            prompt: "Search notes..."
        )
    }
}

/// A placeholder view shown when no note is selected.
struct EmptyStateView: View {
    var body: some View {
        VStack(spacing: 16) {
            Image(systemName: "note.text")
                .font(.system(size: 64))
                .foregroundStyle(.tertiary)

            Text("No Note Selected")
                .font(.title2)
                .fontWeight(.medium)
                .foregroundStyle(.secondary)

            Text("Select a note from the list or create a new one.")
                .font(.subheadline)
                .foregroundStyle(.tertiary)
                .multilineTextAlignment(.center)
                .frame(maxWidth: 300)
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity)
        .background(Color(.systemGroupedBackground))
    }
}

// MARK: - Preview

#Preview {
    ContentView()
        .environmentObject(NotesViewModel())
}
