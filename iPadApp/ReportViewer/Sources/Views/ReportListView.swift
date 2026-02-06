import SwiftUI

/// Displays a filterable, searchable list of reports based on the selected sidebar category.
/// Optimized for iPad with support for swipe actions, pull-to-refresh, and contextual menus.
struct ReportListView: View {
    @EnvironmentObject var appState: AppState
    let category: SidebarCategory

    var body: some View {
        Group {
            if category == .settings {
                SettingsView()
            } else {
                reportList
            }
        }
    }

    private var reportList: some View {
        List(selection: $appState.selectedReport) {
            if filteredReports.isEmpty {
                emptyState
            } else {
                ForEach(filteredReports) { report in
                    ReportRowView(report: report)
                        .tag(report)
                        .swipeActions(edge: .trailing) {
                            Button(role: .destructive) {
                                appState.deleteReport(report)
                            } label: {
                                Label("Delete", systemImage: "trash")
                            }
                        }
                        .swipeActions(edge: .leading) {
                            Button {
                                appState.toggleFavorite(report)
                            } label: {
                                Label(
                                    report.isFavorite ? "Unfavorite" : "Favorite",
                                    systemImage: report.isFavorite ? "star.slash" : "star.fill"
                                )
                            }
                            .tint(.yellow)
                        }
                        .contextMenu {
                            Button {
                                appState.selectedReport = report
                            } label: {
                                Label("View Details", systemImage: "doc.text.magnifyingglass")
                            }
                            Button {
                                appState.toggleFavorite(report)
                            } label: {
                                Label(
                                    report.isFavorite ? "Remove from Favorites" : "Add to Favorites",
                                    systemImage: report.isFavorite ? "star.slash" : "star"
                                )
                            }
                            Divider()
                            Button {
                                // Share action
                            } label: {
                                Label("Share", systemImage: "square.and.arrow.up")
                            }
                            Divider()
                            Button(role: .destructive) {
                                appState.deleteReport(report)
                            } label: {
                                Label("Delete", systemImage: "trash")
                            }
                        }
                }
            }
        }
        .listStyle(.insetGrouped)
        .searchable(text: $appState.searchText, prompt: "Search reports...")
        .refreshable {
            await appState.refreshReports()
        }
        .navigationTitle(category.rawValue)
        .toolbar {
            ToolbarItem(placement: .primaryAction) {
                Menu {
                    Button {
                        // Sort by date
                    } label: {
                        Label("Sort by Date", systemImage: "calendar")
                    }
                    Button {
                        // Sort by name
                    } label: {
                        Label("Sort by Name", systemImage: "textformat")
                    }
                    Button {
                        // Sort by status
                    } label: {
                        Label("Sort by Status", systemImage: "arrow.up.arrow.down")
                    }
                } label: {
                    Label("Sort", systemImage: "arrow.up.arrow.down.circle")
                }
            }
        }
        .overlay {
            if appState.isLoading {
                ProgressView("Loading reports...")
                    .padding()
                    .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 12))
            }
        }
    }

    private var emptyState: some View {
        ContentUnavailableView {
            Label("No Reports", systemImage: "doc.text")
        } description: {
            Text("No reports found for this category.")
        } actions: {
            Button("Refresh") {
                Task {
                    await appState.refreshReports()
                }
            }
            .buttonStyle(.borderedProminent)
        }
    }

    private var filteredReports: [Report] {
        let baseReports = appState.filteredReports

        switch category {
        case .allReports:
            return baseReports
        case .today:
            return baseReports.filter { $0.isToday }
        case .thisWeek:
            return baseReports.filter { $0.isThisWeek }
        case .favorites:
            return baseReports.filter { $0.isFavorite }
        case .settings:
            return []
        }
    }
}

#Preview {
    NavigationStack {
        ReportListView(category: .allReports)
            .environmentObject(AppState())
    }
}
