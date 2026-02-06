import SwiftUI

/// Displays the filtered list of notes in the content column.
/// Supports swipe actions, context menus, and selection.
struct NoteListView: View {
    @EnvironmentObject var viewModel: NotesViewModel

    var body: some View {
        Group {
            if viewModel.filteredNotes.isEmpty {
                emptyFilterView
            } else {
                notesList
            }
        }
        .navigationTitle(viewModel.selectedFilter.title)
        .toolbar {
            ToolbarItem(placement: .primaryAction) {
                Button {
                    viewModel.createNote()
                } label: {
                    Image(systemName: "plus")
                }
                .help("New Note")
            }
        }
    }

    // MARK: - Notes List

    private var notesList: some View {
        List(viewModel.filteredNotes, selection: Binding(
            get: { viewModel.selectedNote },
            set: { viewModel.selectedNote = $0 }
        )) { note in
            NoteRowView(note: note)
                .tag(note)
                .swipeActions(edge: .trailing, allowsFullSwipe: true) {
                    Button(role: .destructive) {
                        withAnimation {
                            viewModel.deleteNote(note)
                        }
                    } label: {
                        Label("Delete", systemImage: "trash")
                    }
                }
                .swipeActions(edge: .leading, allowsFullSwipe: true) {
                    Button {
                        withAnimation {
                            viewModel.toggleFavorite(note)
                        }
                    } label: {
                        Label(
                            note.isFavorite ? "Unfavorite" : "Favorite",
                            systemImage: note.isFavorite ? "star.slash" : "star.fill"
                        )
                    }
                    .tint(.yellow)
                }
                .contextMenu {
                    contextMenu(for: note)
                }
        }
        .listStyle(.insetGrouped)
    }

    // MARK: - Context Menu

    @ViewBuilder
    private func contextMenu(for note: Note) -> some View {
        Button {
            viewModel.toggleFavorite(note)
        } label: {
            Label(
                note.isFavorite ? "Remove from Favorites" : "Add to Favorites",
                systemImage: note.isFavorite ? "star.slash" : "star.fill"
            )
        }

        Divider()

        Menu("Move to Category") {
            ForEach(Category.allCases) { category in
                Button {
                    viewModel.moveNote(note, to: category)
                } label: {
                    Label(category.rawValue, systemImage: category.iconName)
                    if note.category == category {
                        Image(systemName: "checkmark")
                    }
                }
            }
        }

        Divider()

        Button(role: .destructive) {
            viewModel.deleteNote(note)
        } label: {
            Label("Delete", systemImage: "trash")
        }
    }

    // MARK: - Empty State

    private var emptyFilterView: some View {
        VStack(spacing: 12) {
            Image(systemName: "doc.text.magnifyingglass")
                .font(.system(size: 48))
                .foregroundStyle(.tertiary)

            Text("No Notes Found")
                .font(.title3)
                .fontWeight(.medium)
                .foregroundStyle(.secondary)

            if !viewModel.searchText.isEmpty {
                Text("Try a different search term.")
                    .font(.subheadline)
                    .foregroundStyle(.tertiary)
            } else {
                Button("Create Note") {
                    viewModel.createNote()
                }
                .buttonStyle(.borderedProminent)
            }
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity)
    }
}

// MARK: - Note Row View

/// A single row in the notes list showing title, preview, date, and metadata.
struct NoteRowView: View {
    let note: Note

    var body: some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack {
                Text(note.title.isEmpty ? "Untitled Note" : note.title)
                    .font(.headline)
                    .lineLimit(1)

                Spacer()

                if note.isFavorite {
                    Image(systemName: "star.fill")
                        .font(.caption)
                        .foregroundStyle(.yellow)
                }
            }

            Text(note.preview)
                .font(.subheadline)
                .foregroundStyle(.secondary)
                .lineLimit(2)

            HStack(spacing: 8) {
                Text(note.formattedDate)
                    .font(.caption)
                    .foregroundStyle(.tertiary)

                Label(note.category.rawValue, systemImage: note.category.iconName)
                    .font(.caption)
                    .foregroundStyle(note.category.color)
                    .padding(.horizontal, 6)
                    .padding(.vertical, 2)
                    .background(
                        Capsule()
                            .fill(note.category.color.opacity(0.1))
                    )

                if !note.tags.isEmpty {
                    Image(systemName: "tag")
                        .font(.caption2)
                        .foregroundStyle(.tertiary)
                    Text(note.tags.prefix(2).joined(separator: ", "))
                        .font(.caption2)
                        .foregroundStyle(.tertiary)
                        .lineLimit(1)
                }
            }
        }
        .padding(.vertical, 4)
    }
}

// MARK: - Preview

#Preview {
    NavigationSplitView {
        Text("Sidebar")
    } content: {
        NoteListView()
            .environmentObject(NotesViewModel())
    } detail: {
        Text("Detail")
    }
}
