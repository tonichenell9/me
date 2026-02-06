import SwiftUI

/// iPad sidebar navigation with report categories, quick stats, and settings access.
struct SidebarView: View {
    @EnvironmentObject var appState: AppState
    @Binding var selectedCategory: SidebarCategory

    var body: some View {
        List(selection: $selectedCategory) {
            Section("Reports") {
                ForEach(SidebarCategory.allCases.filter { $0 != .settings }) { category in
                    Label {
                        HStack {
                            Text(category.rawValue)
                            Spacer()
                            Text("\(countForCategory(category))")
                                .font(.caption)
                                .foregroundStyle(.secondary)
                                .padding(.horizontal, 8)
                                .padding(.vertical, 2)
                                .background(Color(.systemGray5))
                                .clipShape(Capsule())
                        }
                    } icon: {
                        Image(systemName: category.iconName)
                            .foregroundStyle(category.color)
                    }
                    .tag(category)
                }
            }

            Section("Quick Stats") {
                VStack(alignment: .leading, spacing: 12) {
                    StatRow(
                        title: "Total Reports",
                        value: "\(appState.reports.count)",
                        icon: "doc.text",
                        color: .blue
                    )
                    StatRow(
                        title: "Completed",
                        value: "\(appState.reports.filter { $0.status == .completed }.count)",
                        icon: "checkmark.circle",
                        color: .green
                    )
                    if let lastSync = appState.lastSyncDate {
                        StatRow(
                            title: "Last Sync",
                            value: lastSync.formatted(.relative(presentation: .named)),
                            icon: "arrow.triangle.2.circlepath",
                            color: .orange
                        )
                    }
                }
                .padding(.vertical, 4)
            }

            Section {
                Label {
                    Text(SidebarCategory.settings.rawValue)
                } icon: {
                    Image(systemName: SidebarCategory.settings.iconName)
                        .foregroundStyle(SidebarCategory.settings.color)
                }
                .tag(SidebarCategory.settings)
            }
        }
        .listStyle(.sidebar)
        .navigationTitle("Report Viewer")
        .toolbar {
            ToolbarItem(placement: .primaryAction) {
                Button {
                    Task {
                        await appState.refreshReports()
                    }
                } label: {
                    Label("Refresh", systemImage: "arrow.clockwise")
                }
                .disabled(appState.isLoading)
            }
        }
    }

    private func countForCategory(_ category: SidebarCategory) -> Int {
        switch category {
        case .allReports:
            return appState.reports.count
        case .today:
            return appState.reports.filter { $0.isToday }.count
        case .thisWeek:
            return appState.reports.filter { $0.isThisWeek }.count
        case .favorites:
            return appState.reports.filter { $0.isFavorite }.count
        case .settings:
            return 0
        }
    }
}

/// A single stat row displayed in the sidebar
struct StatRow: View {
    let title: String
    let value: String
    let icon: String
    let color: Color

    var body: some View {
        HStack(spacing: 10) {
            Image(systemName: icon)
                .foregroundStyle(color)
                .frame(width: 20)
            VStack(alignment: .leading, spacing: 2) {
                Text(title)
                    .font(.caption)
                    .foregroundStyle(.secondary)
                Text(value)
                    .font(.subheadline)
                    .fontWeight(.medium)
            }
        }
    }
}

#Preview {
    NavigationSplitView {
        SidebarView(selectedCategory: .constant(.allReports))
            .environmentObject(AppState())
    } detail: {
        Text("Detail")
    }
}
