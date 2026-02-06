import SwiftUI

/// Categories for organizing notes.
/// Each category has an associated color and SF Symbol icon.
enum Category: String, CaseIterable, Codable, Identifiable {
    case general = "General"
    case work = "Work"
    case personal = "Personal"
    case ideas = "Ideas"
    case archive = "Archive"

    var id: String { rawValue }

    /// The SF Symbol name associated with this category.
    var iconName: String {
        switch self {
        case .general:
            return "note.text"
        case .work:
            return "briefcase.fill"
        case .personal:
            return "person.fill"
        case .ideas:
            return "lightbulb.fill"
        case .archive:
            return "archivebox.fill"
        }
    }

    /// The color associated with this category.
    var color: Color {
        switch self {
        case .general:
            return .blue
        case .work:
            return .orange
        case .personal:
            return .green
        case .ideas:
            return .purple
        case .archive:
            return .gray
        }
    }

    /// A short description of the category.
    var description: String {
        switch self {
        case .general:
            return "General purpose notes"
        case .work:
            return "Work-related notes and tasks"
        case .personal:
            return "Personal notes and reminders"
        case .ideas:
            return "Ideas and inspiration"
        case .archive:
            return "Archived notes"
        }
    }
}

/// Represents a filter option in the sidebar.
/// Can be a specific category, all notes, or favorites.
enum SidebarFilter: Hashable {
    case all
    case favorites
    case category(Category)

    var title: String {
        switch self {
        case .all:
            return "All Notes"
        case .favorites:
            return "Favorites"
        case .category(let category):
            return category.rawValue
        }
    }

    var iconName: String {
        switch self {
        case .all:
            return "tray.full.fill"
        case .favorites:
            return "star.fill"
        case .category(let category):
            return category.iconName
        }
    }

    var color: Color {
        switch self {
        case .all:
            return .blue
        case .favorites:
            return .yellow
        case .category(let category):
            return category.color
        }
    }
}
