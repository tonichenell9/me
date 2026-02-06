import SwiftUI

/// Root content view implementing an iPad-optimized NavigationSplitView layout.
/// Provides a sidebar with report categories and a detail view for selected reports.
struct ContentView: View {
    @EnvironmentObject var appState: AppState
    @State private var selectedCategory: SidebarCategory = .allReports
    @State private var columnVisibility: NavigationSplitViewVisibility = .all

    var body: some View {
        NavigationSplitView(columnVisibility: $columnVisibility) {
            SidebarView(selectedCategory: $selectedCategory)
        } content: {
            ReportListView(category: selectedCategory)
        } detail: {
            if let report = appState.selectedReport {
                ReportDetailView(report: report)
            } else {
                EmptyDetailView()
            }
        }
        .navigationSplitViewStyle(.balanced)
    }
}

/// Categories displayed in the sidebar
enum SidebarCategory: String, CaseIterable, Identifiable {
    case allReports = "All Reports"
    case today = "Today"
    case thisWeek = "This Week"
    case favorites = "Favorites"
    case settings = "Settings"

    var id: String { rawValue }

    var iconName: String {
        switch self {
        case .allReports: return "doc.text.fill"
        case .today: return "calendar"
        case .thisWeek: return "calendar.badge.clock"
        case .favorites: return "star.fill"
        case .settings: return "gear"
        }
    }

    var color: Color {
        switch self {
        case .allReports: return .blue
        case .today: return .orange
        case .thisWeek: return .purple
        case .favorites: return .yellow
        case .settings: return .gray
        }
    }
}

#Preview {
    ContentView()
        .environmentObject(AppState())
}
