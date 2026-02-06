import SwiftUI

/// A single row in the report list, showing key report details.
/// Designed for readability on iPad with appropriate touch targets.
struct ReportRowView: View {
    let report: Report

    var body: some View {
        HStack(spacing: 16) {
            // Status icon
            ZStack {
                Circle()
                    .fill(statusColor.opacity(0.15))
                    .frame(width: 44, height: 44)
                Image(systemName: report.status.iconName)
                    .foregroundStyle(statusColor)
                    .font(.system(size: 18))
            }

            // Report details
            VStack(alignment: .leading, spacing: 4) {
                HStack {
                    Text(report.title)
                        .font(.headline)
                    if report.isFavorite {
                        Image(systemName: "star.fill")
                            .font(.caption)
                            .foregroundStyle(.yellow)
                    }
                    if report.isToday {
                        Text("Today")
                            .font(.caption2)
                            .fontWeight(.semibold)
                            .padding(.horizontal, 6)
                            .padding(.vertical, 2)
                            .background(Color.blue.opacity(0.15))
                            .foregroundStyle(.blue)
                            .clipShape(Capsule())
                    }
                }

                Text(report.summary)
                    .font(.subheadline)
                    .foregroundStyle(.secondary)
                    .lineLimit(2)

                HStack(spacing: 12) {
                    Label(report.formattedDate, systemImage: "calendar")
                    Label(report.fileSize, systemImage: "doc")
                    Label("\(report.worksheetCount) sheets", systemImage: "tablecells")
                }
                .font(.caption)
                .foregroundStyle(.tertiary)
            }

            Spacer()

            // Chevron
            Image(systemName: "chevron.right")
                .font(.caption)
                .foregroundStyle(.quaternary)
        }
        .padding(.vertical, 4)
    }

    private var statusColor: Color {
        switch report.status {
        case .pending: return .gray
        case .inProgress: return .blue
        case .completed: return .green
        case .failed: return .red
        }
    }
}

#Preview {
    List {
        ReportRowView(report: Report(
            title: "Large Deal Report",
            date: Date(),
            summary: "Daily large deal report with updated pipeline data.",
            status: .completed,
            fileSize: "2.4 MB",
            worksheetCount: 3,
            isFavorite: true
        ))
        ReportRowView(report: Report(
            title: "Large Deal Report",
            date: Date().addingTimeInterval(-86400),
            summary: "Yesterday's report with Q4 data.",
            status: .completed,
            fileSize: "2.1 MB",
            worksheetCount: 3
        ))
    }
    .listStyle(.insetGrouped)
}
