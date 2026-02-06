import SwiftUI
import Combine

/// Centralized application state that manages global settings, authentication,
/// and shared data across the app.
final class AppState: ObservableObject {
    // MARK: - Published Properties

    @Published var isAuthenticated: Bool = false
    @Published var serverURL: String = ""
    @Published var lastSyncDate: Date?
    @Published var reports: [Report] = []
    @Published var isLoading: Bool = false
    @Published var errorMessage: String?
    @Published var selectedReport: Report?
    @Published var searchText: String = ""

    // MARK: - Services

    private let reportService = ReportService()
    private var cancellables = Set<AnyCancellable>()

    // MARK: - Computed Properties

    var filteredReports: [Report] {
        if searchText.isEmpty {
            return reports
        }
        return reports.filter { report in
            report.title.localizedCaseInsensitiveContains(searchText) ||
            report.summary.localizedCaseInsensitiveContains(searchText)
        }
    }

    var recentReports: [Report] {
        Array(reports.sorted { $0.date > $1.date }.prefix(5))
    }

    var reportsByMonth: [String: [Report]] {
        Dictionary(grouping: reports) { report in
            let formatter = DateFormatter()
            formatter.dateFormat = "MMMM yyyy"
            return formatter.string(from: report.date)
        }
    }

    // MARK: - Initialization

    init() {
        loadSavedSettings()
        loadSampleData()
    }

    // MARK: - Methods

    func loadSavedSettings() {
        if let savedURL = UserDefaults.standard.string(forKey: "serverURL") {
            serverURL = savedURL
        }
        isAuthenticated = UserDefaults.standard.bool(forKey: "isAuthenticated")
        if let lastSync = UserDefaults.standard.object(forKey: "lastSyncDate") as? Date {
            lastSyncDate = lastSync
        }
    }

    func saveSettings() {
        UserDefaults.standard.set(serverURL, forKey: "serverURL")
        UserDefaults.standard.set(isAuthenticated, forKey: "isAuthenticated")
        if let lastSync = lastSyncDate {
            UserDefaults.standard.set(lastSync, forKey: "lastSyncDate")
        }
    }

    func refreshReports() async {
        await MainActor.run {
            isLoading = true
            errorMessage = nil
        }

        do {
            let fetchedReports = try await reportService.fetchReports(from: serverURL)
            await MainActor.run {
                reports = fetchedReports
                lastSyncDate = Date()
                isLoading = false
                saveSettings()
            }
        } catch {
            await MainActor.run {
                errorMessage = error.localizedDescription
                isLoading = false
            }
        }
    }

    func deleteReport(_ report: Report) {
        reports.removeAll { $0.id == report.id }
        if selectedReport?.id == report.id {
            selectedReport = nil
        }
    }

    func toggleFavorite(_ report: Report) {
        if let index = reports.firstIndex(where: { $0.id == report.id }) {
            reports[index].isFavorite.toggle()
        }
    }

    /// Load sample data for preview and initial demonstration
    private func loadSampleData() {
        let calendar = Calendar.current
        let today = Date()

        reports = [
            Report(
                title: "Large Deal Report",
                date: today,
                summary: "Daily large deal report with updated pipeline data and forecasts.",
                status: .completed,
                fileSize: "2.4 MB",
                worksheetCount: 3,
                isFavorite: true
            ),
            Report(
                title: "Large Deal Report",
                date: calendar.date(byAdding: .day, value: -1, to: today) ?? today,
                summary: "Yesterday's large deal report with Q4 pipeline updates.",
                status: .completed,
                fileSize: "2.1 MB",
                worksheetCount: 3,
                isFavorite: false
            ),
            Report(
                title: "Large Deal Report",
                date: calendar.date(byAdding: .day, value: -2, to: today) ?? today,
                summary: "Large deal report with revised forecasting models.",
                status: .completed,
                fileSize: "2.3 MB",
                worksheetCount: 3,
                isFavorite: false
            ),
            Report(
                title: "Large Deal Report",
                date: calendar.date(byAdding: .day, value: -3, to: today) ?? today,
                summary: "Mid-week large deal report with new account additions.",
                status: .completed,
                fileSize: "2.5 MB",
                worksheetCount: 3,
                isFavorite: true
            ),
            Report(
                title: "Large Deal Report",
                date: calendar.date(byAdding: .day, value: -4, to: today) ?? today,
                summary: "Weekly kickoff large deal report.",
                status: .completed,
                fileSize: "2.2 MB",
                worksheetCount: 3,
                isFavorite: false
            ),
            Report(
                title: "Large Deal Report",
                date: calendar.date(byAdding: .day, value: -7, to: today) ?? today,
                summary: "Previous week's final large deal report.",
                status: .completed,
                fileSize: "2.6 MB",
                worksheetCount: 3,
                isFavorite: false
            ),
        ]
    }
}
