import SwiftUI
import Combine

/// View model for the report list, handling sorting, filtering, and batch operations.
@MainActor
final class ReportListViewModel: ObservableObject {
    // MARK: - Published Properties

    @Published var sortOrder: SortOrder = .dateDescending
    @Published var filterStatus: ReportStatus?
    @Published var isSelectionMode: Bool = false
    @Published var selectedReportIDs: Set<UUID> = []

    // MARK: - Sort Order

    enum SortOrder: String, CaseIterable {
        case dateDescending = "Newest First"
        case dateAscending = "Oldest First"
        case nameAscending = "Name (A-Z)"
        case nameDescending = "Name (Z-A)"
        case status = "By Status"

        var iconName: String {
            switch self {
            case .dateDescending: return "arrow.down"
            case .dateAscending: return "arrow.up"
            case .nameAscending: return "textformat"
            case .nameDescending: return "textformat"
            case .status: return "circle.grid.2x1"
            }
        }
    }

    // MARK: - Methods

    func sortedReports(_ reports: [Report]) -> [Report] {
        var filtered = reports

        // Apply status filter
        if let status = filterStatus {
            filtered = filtered.filter { $0.status == status }
        }

        // Apply sort order
        switch sortOrder {
        case .dateDescending:
            return filtered.sorted { $0.date > $1.date }
        case .dateAscending:
            return filtered.sorted { $0.date < $1.date }
        case .nameAscending:
            return filtered.sorted { $0.title.localizedCompare($1.title) == .orderedAscending }
        case .nameDescending:
            return filtered.sorted { $0.title.localizedCompare($1.title) == .orderedDescending }
        case .status:
            return filtered.sorted { $0.status.rawValue < $1.status.rawValue }
        }
    }

    func toggleSelection(for reportID: UUID) {
        if selectedReportIDs.contains(reportID) {
            selectedReportIDs.remove(reportID)
        } else {
            selectedReportIDs.insert(reportID)
        }
    }

    func selectAll(from reports: [Report]) {
        selectedReportIDs = Set(reports.map { $0.id })
    }

    func deselectAll() {
        selectedReportIDs.removeAll()
    }
}
