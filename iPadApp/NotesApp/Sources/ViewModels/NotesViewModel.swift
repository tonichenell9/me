import Foundation
import Combine

/// Main view model for the Notes app.
/// Manages the collection of notes, filtering, search, and persistence.
class NotesViewModel: ObservableObject {

    // MARK: - Published Properties

    /// All notes in the app.
    @Published var notes: [Note] = []

    /// The currently selected sidebar filter.
    @Published var selectedFilter: SidebarFilter = .all

    /// The currently selected note (shown in the editor).
    @Published var selectedNote: Note?

    /// The current search query text.
    @Published var searchText: String = ""

    /// The sort order for notes.
    @Published var sortOrder: SortOrder = .modifiedDescending

    /// Whether the app is currently loading data.
    @Published var isLoading: Bool = false

    /// An error message to display, if any.
    @Published var errorMessage: String?

    // MARK: - Sort Order

    enum SortOrder: String, CaseIterable {
        case modifiedDescending = "Recently Modified"
        case modifiedAscending = "Oldest Modified"
        case titleAscending = "Title (A-Z)"
        case titleDescending = "Title (Z-A)"
        case createdDescending = "Recently Created"
        case createdAscending = "Oldest Created"
    }

    // MARK: - Computed Properties

    /// Notes filtered by the current sidebar filter and search text.
    var filteredNotes: [Note] {
        var result = notes

        // Apply sidebar filter
        switch selectedFilter {
        case .all:
            break
        case .favorites:
            result = result.filter { $0.isFavorite }
        case .category(let category):
            result = result.filter { $0.category == category }
        }

        // Apply search filter
        if !searchText.isEmpty {
            let query = searchText.lowercased()
            result = result.filter { note in
                note.title.lowercased().contains(query) ||
                note.content.lowercased().contains(query) ||
                note.tags.contains(where: { $0.lowercased().contains(query) })
            }
        }

        // Apply sort
        result.sort { a, b in
            switch sortOrder {
            case .modifiedDescending:
                return a.modifiedAt > b.modifiedAt
            case .modifiedAscending:
                return a.modifiedAt < b.modifiedAt
            case .titleAscending:
                return a.title.localizedCaseInsensitiveCompare(b.title) == .orderedAscending
            case .titleDescending:
                return a.title.localizedCaseInsensitiveCompare(b.title) == .orderedDescending
            case .createdDescending:
                return a.createdAt > b.createdAt
            case .createdAscending:
                return a.createdAt < b.createdAt
            }
        }

        return result
    }

    /// Count of notes per category for sidebar badges.
    var noteCounts: [SidebarFilter: Int] {
        var counts: [SidebarFilter: Int] = [
            .all: notes.count,
            .favorites: notes.filter(\.isFavorite).count
        ]
        for category in Category.allCases {
            counts[.category(category)] = notes.filter { $0.category == category }.count
        }
        return counts
    }

    // MARK: - Initialization

    init() {
        loadNotes()
    }

    // MARK: - CRUD Operations

    /// Creates a new note and selects it for editing.
    /// - Parameter category: The category for the new note (defaults based on current filter).
    @discardableResult
    func createNote(in category: Category? = nil) -> Note {
        let noteCategory: Category
        if let category = category {
            noteCategory = category
        } else if case .category(let cat) = selectedFilter {
            noteCategory = cat
        } else {
            noteCategory = .general
        }

        let note = Note(
            title: "",
            content: "",
            category: noteCategory
        )

        notes.insert(note, at: 0)
        selectedNote = note
        saveNotes()
        return note
    }

    /// Updates a note with new values.
    /// - Parameter note: The note with updated values.
    func updateNote(_ note: Note) {
        guard let index = notes.firstIndex(where: { $0.id == note.id }) else { return }
        var updated = note
        updated.modifiedAt = Date()
        notes[index] = updated

        if selectedNote?.id == note.id {
            selectedNote = updated
        }
        saveNotes()
    }

    /// Deletes a note.
    /// - Parameter note: The note to delete.
    func deleteNote(_ note: Note) {
        notes.removeAll { $0.id == note.id }
        if selectedNote?.id == note.id {
            selectedNote = filteredNotes.first
        }
        saveNotes()
    }

    /// Toggles the favorite status of a note.
    /// - Parameter note: The note to toggle.
    func toggleFavorite(_ note: Note) {
        guard let index = notes.firstIndex(where: { $0.id == note.id }) else { return }
        notes[index].isFavorite.toggle()
        notes[index].modifiedAt = Date()

        if selectedNote?.id == note.id {
            selectedNote = notes[index]
        }
        saveNotes()
    }

    /// Moves a note to a different category.
    /// - Parameters:
    ///   - note: The note to move.
    ///   - category: The destination category.
    func moveNote(_ note: Note, to category: Category) {
        guard let index = notes.firstIndex(where: { $0.id == note.id }) else { return }
        notes[index].category = category
        notes[index].modifiedAt = Date()

        if selectedNote?.id == note.id {
            selectedNote = notes[index]
        }
        saveNotes()
    }

    // MARK: - Persistence

    /// The file URL where notes are stored.
    private var notesFileURL: URL {
        let documentsDirectory = FileManager.default.urls(
            for: .documentDirectory,
            in: .userDomainMask
        ).first!
        return documentsDirectory.appendingPathComponent("notes.json")
    }

    /// Loads notes from disk. Falls back to sample data if no saved notes exist.
    func loadNotes() {
        isLoading = true
        defer { isLoading = false }

        do {
            let data = try Data(contentsOf: notesFileURL)
            let decoder = JSONDecoder()
            decoder.dateDecodingStrategy = .iso8601
            notes = try decoder.decode([Note].self, from: data)

            if let first = filteredNotes.first {
                selectedNote = first
            }
        } catch {
            // No saved data — use sample notes for first launch
            notes = Note.sampleNotes
            if let first = filteredNotes.first {
                selectedNote = first
            }
            saveNotes()
        }
    }

    /// Saves all notes to disk as JSON.
    func saveNotes() {
        do {
            let encoder = JSONEncoder()
            encoder.dateEncodingStrategy = .iso8601
            encoder.outputFormatting = .prettyPrinted
            let data = try encoder.encode(notes)
            try data.write(to: notesFileURL, options: .atomic)
        } catch {
            errorMessage = "Failed to save notes: \(error.localizedDescription)"
        }
    }
}
