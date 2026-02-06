import SwiftUI

/// The note editor view shown in the detail column.
/// Provides full editing capabilities with title, content, category, tags, and metadata.
struct NoteEditorView: View {
    @EnvironmentObject var viewModel: NotesViewModel

    /// The note being edited, passed from the parent.
    let note: Note

    /// Local state for the note title (debounced save).
    @State private var title: String = ""

    /// Local state for the note content (debounced save).
    @State private var content: String = ""

    /// Local state for the selected category.
    @State private var category: Category = .general

    /// Local state for the tags text field.
    @State private var tagsText: String = ""

    /// Whether the metadata inspector is shown.
    @State private var showInspector: Bool = false

    /// Debounce timer for auto-saving.
    @State private var saveTimer: Timer?

    var body: some View {
        VStack(spacing: 0) {
            // Title Field
            TextField("Note Title", text: $title)
                .font(.largeTitle)
                .fontWeight(.bold)
                .textFieldStyle(.plain)
                .padding(.horizontal, 20)
                .padding(.top, 20)
                .padding(.bottom, 8)
                .onChange(of: title) { _, newValue in
                    scheduleSave()
                }

            Divider()
                .padding(.horizontal, 20)

            // Content Editor
            TextEditor(text: $content)
                .font(.body)
                .scrollContentBackground(.hidden)
                .padding(.horizontal, 16)
                .padding(.top, 8)
                .onChange(of: content) { _, newValue in
                    scheduleSave()
                }

            // Bottom Bar with metadata
            bottomBar
        }
        .background(Color(.systemBackground))
        .navigationBarTitleDisplayMode(.inline)
        .toolbar {
            ToolbarItemGroup(placement: .primaryAction) {
                Button {
                    viewModel.toggleFavorite(note)
                } label: {
                    Image(systemName: note.isFavorite ? "star.fill" : "star")
                        .foregroundStyle(note.isFavorite ? .yellow : .secondary)
                }
                .help("Toggle Favorite")

                Button {
                    showInspector.toggle()
                } label: {
                    Image(systemName: "info.circle")
                }
                .help("Note Info")

                Menu {
                    shareMenu
                } label: {
                    Image(systemName: "ellipsis.circle")
                }
            }
        }
        .sheet(isPresented: $showInspector) {
            NoteInspectorView(note: note)
                .environmentObject(viewModel)
        }
        .onAppear {
            loadNote()
        }
        .onChange(of: note.id) { _, _ in
            loadNote()
        }
    }

    // MARK: - Bottom Bar

    private var bottomBar: some View {
        HStack(spacing: 12) {
            // Category Picker
            Menu {
                ForEach(Category.allCases) { cat in
                    Button {
                        category = cat
                        saveNote()
                    } label: {
                        Label(cat.rawValue, systemImage: cat.iconName)
                        if category == cat {
                            Image(systemName: "checkmark")
                        }
                    }
                }
            } label: {
                Label(category.rawValue, systemImage: category.iconName)
                    .font(.caption)
                    .foregroundStyle(category.color)
                    .padding(.horizontal, 10)
                    .padding(.vertical, 6)
                    .background(
                        Capsule()
                            .fill(category.color.opacity(0.1))
                    )
            }

            // Tags
            Image(systemName: "tag")
                .font(.caption)
                .foregroundStyle(.secondary)

            TextField("Add tags (comma separated)", text: $tagsText)
                .font(.caption)
                .textFieldStyle(.plain)
                .onChange(of: tagsText) { _, _ in
                    scheduleSave()
                }

            Spacer()

            // Word count
            Text(wordCount)
                .font(.caption2)
                .foregroundStyle(.tertiary)

            // Last modified
            Text(note.formattedDate)
                .font(.caption2)
                .foregroundStyle(.tertiary)
        }
        .padding(.horizontal, 20)
        .padding(.vertical, 10)
        .background(Color(.secondarySystemBackground))
    }

    // MARK: - Share Menu

    @ViewBuilder
    private var shareMenu: some View {
        Button {
            copyToClipboard()
        } label: {
            Label("Copy Text", systemImage: "doc.on.doc")
        }

        Button(role: .destructive) {
            viewModel.deleteNote(note)
        } label: {
            Label("Delete Note", systemImage: "trash")
        }
    }

    // MARK: - Helpers

    private var wordCount: String {
        let words = content.split(separator: " ").count
        return "\(words) word\(words == 1 ? "" : "s")"
    }

    private func loadNote() {
        title = note.title
        content = note.content
        category = note.category
        tagsText = note.tags.joined(separator: ", ")
    }

    private func scheduleSave() {
        saveTimer?.invalidate()
        saveTimer = Timer.scheduledTimer(withTimeInterval: 0.5, repeats: false) { _ in
            saveNote()
        }
    }

    private func saveNote() {
        var updated = note
        updated.title = title
        updated.content = content
        updated.category = category
        updated.tags = tagsText
            .split(separator: ",")
            .map { $0.trimmingCharacters(in: .whitespaces) }
            .filter { !$0.isEmpty }
        viewModel.updateNote(updated)
    }

    private func copyToClipboard() {
        #if canImport(UIKit)
        UIPasteboard.general.string = "\(title)\n\n\(content)"
        #endif
    }
}

// MARK: - Note Inspector View

/// A sheet showing detailed metadata about a note.
struct NoteInspectorView: View {
    @EnvironmentObject var viewModel: NotesViewModel
    @Environment(\.dismiss) private var dismiss

    let note: Note

    var body: some View {
        NavigationStack {
            List {
                Section("Details") {
                    LabeledContent("Title", value: note.title.isEmpty ? "Untitled" : note.title)
                    LabeledContent("Category", value: note.category.rawValue)
                    LabeledContent("Favorite", value: note.isFavorite ? "Yes" : "No")
                }

                Section("Statistics") {
                    LabeledContent("Characters", value: "\(note.content.count)")
                    LabeledContent("Words", value: "\(note.content.split(separator: " ").count)")
                    LabeledContent("Lines", value: "\(note.content.components(separatedBy: .newlines).count)")
                }

                Section("Dates") {
                    LabeledContent("Created") {
                        Text(note.createdAt, style: .date)
                    }
                    LabeledContent("Modified") {
                        Text(note.modifiedAt, style: .relative)
                    }
                }

                if !note.tags.isEmpty {
                    Section("Tags") {
                        FlowLayout(spacing: 8) {
                            ForEach(note.tags, id: \.self) { tag in
                                Text(tag)
                                    .font(.caption)
                                    .padding(.horizontal, 10)
                                    .padding(.vertical, 4)
                                    .background(
                                        Capsule()
                                            .fill(Color(.systemGray5))
                                    )
                            }
                        }
                    }
                }

                Section {
                    Button(role: .destructive) {
                        viewModel.deleteNote(note)
                        dismiss()
                    } label: {
                        Label("Delete Note", systemImage: "trash")
                            .foregroundStyle(.red)
                    }
                }
            }
            .navigationTitle("Note Info")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .confirmationAction) {
                    Button("Done") {
                        dismiss()
                    }
                }
            }
        }
    }
}

// MARK: - Preview

#Preview {
    NavigationSplitView {
        Text("Sidebar")
    } content: {
        Text("List")
    } detail: {
        NoteEditorView(note: Note.sampleNotes[0])
            .environmentObject(NotesViewModel())
    }
}
