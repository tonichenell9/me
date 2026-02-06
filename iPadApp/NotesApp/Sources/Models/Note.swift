import Foundation

/// Represents a single note in the app.
/// Conforms to Identifiable, Codable, and Hashable for use in SwiftUI lists and persistence.
struct Note: Identifiable, Codable, Hashable {
    /// Unique identifier for the note.
    let id: UUID

    /// The title of the note.
    var title: String

    /// The body content of the note (supports rich text as plain string).
    var content: String

    /// The category this note belongs to.
    var category: Category

    /// Whether this note has been marked as a favorite.
    var isFavorite: Bool

    /// The date this note was created.
    let createdAt: Date

    /// The date this note was last modified.
    var modifiedAt: Date

    /// Tags associated with this note for filtering and search.
    var tags: [String]

    /// Creates a new note with default values.
    /// - Parameters:
    ///   - title: The note title (default: "Untitled Note")
    ///   - content: The note content (default: empty string)
    ///   - category: The category (default: .general)
    init(
        id: UUID = UUID(),
        title: String = "Untitled Note",
        content: String = "",
        category: Category = .general,
        isFavorite: Bool = false,
        tags: [String] = []
    ) {
        self.id = id
        self.title = title
        self.content = content
        self.category = category
        self.isFavorite = isFavorite
        self.createdAt = Date()
        self.modifiedAt = Date()
        self.tags = tags
    }

    /// A short preview of the note content, truncated to 100 characters.
    var preview: String {
        if content.isEmpty {
            return "No content"
        }
        let trimmed = content.trimmingCharacters(in: .whitespacesAndNewlines)
        if trimmed.count <= 100 {
            return trimmed
        }
        return String(trimmed.prefix(100)) + "..."
    }

    /// A formatted string of the modification date.
    var formattedDate: String {
        let formatter = DateFormatter()
        let calendar = Calendar.current

        if calendar.isDateInToday(modifiedAt) {
            formatter.dateFormat = "h:mm a"
            return "Today, " + formatter.string(from: modifiedAt)
        } else if calendar.isDateInYesterday(modifiedAt) {
            formatter.dateFormat = "h:mm a"
            return "Yesterday, " + formatter.string(from: modifiedAt)
        } else {
            formatter.dateStyle = .medium
            formatter.timeStyle = .short
            return formatter.string(from: modifiedAt)
        }
    }
}

// MARK: - Sample Data

extension Note {
    /// Sample notes used for previews and initial app state.
    static let sampleNotes: [Note] = [
        Note(
            title: "Welcome to Notes",
            content: """
            Welcome to your new iPad Notes app! This app is designed to help you \
            organize your thoughts, ideas, and tasks efficiently.

            Features:
            • Create and organize notes by category
            • Mark notes as favorites for quick access
            • Tag notes for powerful filtering
            • Full-text search across all notes
            • iPad-optimized split view navigation
            • Keyboard shortcuts for power users

            Start by creating a new note using the + button or ⌘N.
            """,
            category: .general,
            isFavorite: true,
            tags: ["welcome", "getting-started"]
        ),
        Note(
            title: "Project Ideas",
            content: """
            1. Build a weather dashboard app
            2. Create a recipe organizer with image support
            3. Develop a habit tracker with charts
            4. Design a collaborative whiteboard tool
            5. Implement a markdown editor with live preview
            """,
            category: .ideas,
            isFavorite: false,
            tags: ["projects", "brainstorm"]
        ),
        Note(
            title: "Meeting Notes - Sprint Planning",
            content: """
            Sprint 14 Planning - Feb 2026

            Attendees: Team Alpha
            Duration: 1 hour

            Key Decisions:
            - Prioritize performance improvements
            - Defer new feature X to Sprint 15
            - Add accessibility audit to backlog

            Action Items:
            - [ ] Update Jira board with new priorities
            - [ ] Schedule design review for Wednesday
            - [ ] Share updated timeline with stakeholders
            """,
            category: .work,
            isFavorite: true,
            tags: ["meeting", "sprint", "planning"]
        ),
        Note(
            title: "SwiftUI Tips",
            content: """
            Useful SwiftUI patterns for iPad development:

            1. NavigationSplitView for sidebar layouts
            2. @Environment for dependency injection
            3. .searchable() modifier for built-in search
            4. .toolbar() for iPad-optimized toolbars
            5. GeometryReader for adaptive layouts
            6. .onDrag/.onDrop for drag and drop
            7. .focusedSceneValue for keyboard shortcuts
            """,
            category: .ideas,
            isFavorite: false,
            tags: ["swift", "swiftui", "coding"]
        ),
        Note(
            title: "Grocery List",
            content: """
            - Organic eggs
            - Whole wheat bread
            - Avocados (3)
            - Cherry tomatoes
            - Fresh basil
            - Olive oil
            - Greek yogurt
            - Almonds
            - Sparkling water
            """,
            category: .personal,
            isFavorite: false,
            tags: ["shopping", "groceries"]
        )
    ]
}
