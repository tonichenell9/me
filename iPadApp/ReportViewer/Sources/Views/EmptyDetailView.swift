import SwiftUI

/// Placeholder view shown when no report is selected in the detail pane.
/// Provides a visually appealing empty state with guidance for the user.
struct EmptyDetailView: View {
    @EnvironmentObject var appState: AppState

    var body: some View {
        VStack(spacing: 24) {
            Spacer()

            Image(systemName: "doc.text.magnifyingglass")
                .font(.system(size: 64))
                .foregroundStyle(.secondary)
                .symbolEffect(.pulse, options: .repeating)

            VStack(spacing: 8) {
                Text("Select a Report")
                    .font(.title2)
                    .fontWeight(.semibold)

                Text("Choose a report from the list to view its details,\nworksheets, and available actions.")
                    .font(.body)
                    .foregroundStyle(.secondary)
                    .multilineTextAlignment(.center)
            }

            if !appState.reports.isEmpty {
                VStack(spacing: 8) {
                    Text("Quick Access")
                        .font(.caption)
                        .foregroundStyle(.tertiary)
                        .textCase(.uppercase)

                    HStack(spacing: 12) {
                        ForEach(appState.recentReports.prefix(3)) { report in
                            Button {
                                appState.selectedReport = report
                            } label: {
                                VStack(spacing: 4) {
                                    Image(systemName: "doc.text.fill")
                                        .font(.title3)
                                    Text(report.shortDate)
                                        .font(.caption)
                                        .fontWeight(.medium)
                                }
                                .frame(width: 70, height: 70)
                                .background(Color.blue.opacity(0.1))
                                .foregroundStyle(.blue)
                                .clipShape(RoundedRectangle(cornerRadius: 12))
                            }
                            .buttonStyle(.plain)
                        }
                    }
                }
            }

            Spacer()
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity)
        .background(Color(.systemGroupedBackground))
    }
}

#Preview {
    EmptyDetailView()
        .environmentObject(AppState())
}
