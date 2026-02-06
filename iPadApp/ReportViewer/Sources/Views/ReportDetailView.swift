import SwiftUI

/// Detailed view for a single report, showing metadata, worksheets, and action buttons.
/// Optimized for iPad's larger screen with a grid layout and rich information display.
struct ReportDetailView: View {
    @EnvironmentObject var appState: AppState
    let report: Report
    @State private var selectedWorksheet: Worksheet?
    @State private var showingShareSheet = false
    @State private var showingDeleteConfirmation = false

    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 24) {
                // Header Card
                headerCard

                // Action Buttons
                actionButtons

                // Worksheets Section
                worksheetsSection

                // Metadata Section
                metadataSection
            }
            .padding()
        }
        .background(Color(.systemGroupedBackground))
        .navigationTitle(report.title)
        .navigationBarTitleDisplayMode(.large)
        .toolbar {
            ToolbarItemGroup(placement: .primaryAction) {
                Button {
                    appState.toggleFavorite(report)
                } label: {
                    Image(systemName: report.isFavorite ? "star.fill" : "star")
                        .foregroundStyle(report.isFavorite ? .yellow : .gray)
                }

                Button {
                    showingShareSheet = true
                } label: {
                    Image(systemName: "square.and.arrow.up")
                }
            }
        }
        .confirmationDialog(
            "Delete Report",
            isPresented: $showingDeleteConfirmation,
            titleVisibility: .visible
        ) {
            Button("Delete", role: .destructive) {
                appState.deleteReport(report)
            }
            Button("Cancel", role: .cancel) {}
        } message: {
            Text("Are you sure you want to delete \"\(report.filename)\"? This action cannot be undone.")
        }
    }

    // MARK: - Header Card

    private var headerCard: some View {
        VStack(alignment: .leading, spacing: 16) {
            HStack(spacing: 16) {
                // Large status icon
                ZStack {
                    RoundedRectangle(cornerRadius: 16)
                        .fill(statusColor.opacity(0.15))
                        .frame(width: 64, height: 64)
                    Image(systemName: report.status.iconName)
                        .font(.system(size: 28))
                        .foregroundStyle(statusColor)
                }

                VStack(alignment: .leading, spacing: 4) {
                    Text(report.filename)
                        .font(.title3)
                        .fontWeight(.semibold)

                    HStack(spacing: 8) {
                        StatusBadge(status: report.status)
                        if report.isToday {
                            Text("Today")
                                .font(.caption)
                                .fontWeight(.semibold)
                                .padding(.horizontal, 8)
                                .padding(.vertical, 3)
                                .background(Color.blue.opacity(0.15))
                                .foregroundStyle(.blue)
                                .clipShape(Capsule())
                        }
                    }
                }

                Spacer()
            }

            Text(report.summary)
                .font(.body)
                .foregroundStyle(.secondary)

            Divider()

            // Quick stats
            HStack(spacing: 24) {
                DetailStat(title: "Date", value: report.formattedDate, icon: "calendar")
                DetailStat(title: "Size", value: report.fileSize, icon: "doc")
                DetailStat(title: "Sheets", value: "\(report.worksheetCount)", icon: "tablecells")
                DetailStat(title: "Status", value: report.status.rawValue, icon: report.status.iconName)
            }
        }
        .padding(20)
        .background(Color(.systemBackground))
        .clipShape(RoundedRectangle(cornerRadius: 16))
        .shadow(color: .black.opacity(0.05), radius: 8, y: 4)
    }

    // MARK: - Action Buttons

    private var actionButtons: some View {
        HStack(spacing: 12) {
            ActionButton(title: "Open", icon: "doc.text.fill", color: .blue) {
                // Open report action
            }

            ActionButton(title: "Share", icon: "square.and.arrow.up", color: .green) {
                showingShareSheet = true
            }

            ActionButton(title: "Download", icon: "arrow.down.circle.fill", color: .orange) {
                // Download action
            }

            ActionButton(title: "Delete", icon: "trash.fill", color: .red) {
                showingDeleteConfirmation = true
            }
        }
    }

    // MARK: - Worksheets Section

    private var worksheetsSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Worksheets")
                .font(.title3)
                .fontWeight(.semibold)

            LazyVGrid(columns: [
                GridItem(.flexible()),
                GridItem(.flexible()),
                GridItem(.flexible())
            ], spacing: 12) {
                ForEach(report.worksheets) { worksheet in
                    WorksheetCard(worksheet: worksheet, isSelected: selectedWorksheet?.id == worksheet.id)
                        .onTapGesture {
                            withAnimation(.spring(response: 0.3)) {
                                selectedWorksheet = worksheet
                            }
                        }
                }
            }
        }
    }

    // MARK: - Metadata Section

    private var metadataSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("Details")
                .font(.title3)
                .fontWeight(.semibold)

            VStack(spacing: 0) {
                MetadataRow(label: "Report ID", value: report.id.uuidString.prefix(8).uppercased())
                Divider()
                MetadataRow(label: "File Name", value: report.filename)
                Divider()
                MetadataRow(label: "File Size", value: report.fileSize)
                Divider()
                MetadataRow(label: "Generated", value: report.formattedDate)
                Divider()
                MetadataRow(label: "Worksheets", value: "\(report.worksheetCount)")
                Divider()
                MetadataRow(label: "Status", value: report.status.rawValue)
            }
            .background(Color(.systemBackground))
            .clipShape(RoundedRectangle(cornerRadius: 12))
        }
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

// MARK: - Supporting Views

struct StatusBadge: View {
    let status: ReportStatus

    var body: some View {
        HStack(spacing: 4) {
            Image(systemName: status.iconName)
                .font(.caption2)
            Text(status.rawValue)
                .font(.caption)
                .fontWeight(.medium)
        }
        .padding(.horizontal, 8)
        .padding(.vertical, 3)
        .background(color.opacity(0.15))
        .foregroundStyle(color)
        .clipShape(Capsule())
    }

    private var color: Color {
        switch status {
        case .pending: return .gray
        case .inProgress: return .blue
        case .completed: return .green
        case .failed: return .red
        }
    }
}

struct DetailStat: View {
    let title: String
    let value: String
    let icon: String

    var body: some View {
        VStack(spacing: 4) {
            Image(systemName: icon)
                .font(.caption)
                .foregroundStyle(.secondary)
            Text(value)
                .font(.subheadline)
                .fontWeight(.medium)
            Text(title)
                .font(.caption2)
                .foregroundStyle(.tertiary)
        }
        .frame(maxWidth: .infinity)
    }
}

struct ActionButton: View {
    let title: String
    let icon: String
    let color: Color
    let action: () -> Void

    var body: some View {
        Button(action: action) {
            VStack(spacing: 6) {
                Image(systemName: icon)
                    .font(.title3)
                Text(title)
                    .font(.caption)
                    .fontWeight(.medium)
            }
            .frame(maxWidth: .infinity)
            .padding(.vertical, 12)
            .background(color.opacity(0.1))
            .foregroundStyle(color)
            .clipShape(RoundedRectangle(cornerRadius: 12))
        }
        .buttonStyle(.plain)
    }
}

struct WorksheetCard: View {
    let worksheet: Worksheet
    let isSelected: Bool

    var body: some View {
        VStack(alignment: .leading, spacing: 8) {
            HStack {
                Image(systemName: "tablecells")
                    .foregroundStyle(.blue)
                Spacer()
                if isSelected {
                    Image(systemName: "checkmark.circle.fill")
                        .foregroundStyle(.blue)
                }
            }

            Text(worksheet.name)
                .font(.subheadline)
                .fontWeight(.semibold)

            Text(worksheet.description)
                .font(.caption)
                .foregroundStyle(.secondary)
        }
        .padding(12)
        .background(isSelected ? Color.blue.opacity(0.08) : Color(.systemBackground))
        .clipShape(RoundedRectangle(cornerRadius: 12))
        .overlay(
            RoundedRectangle(cornerRadius: 12)
                .stroke(isSelected ? Color.blue : Color(.systemGray4), lineWidth: isSelected ? 2 : 1)
        )
    }
}

struct MetadataRow<V: StringProtocol>: View {
    let label: String
    let value: V

    var body: some View {
        HStack {
            Text(label)
                .foregroundStyle(.secondary)
            Spacer()
            Text(value)
                .fontWeight(.medium)
        }
        .padding(.horizontal, 16)
        .padding(.vertical, 12)
    }
}

#Preview {
    NavigationStack {
        ReportDetailView(report: Report(
            title: "Large Deal Report",
            date: Date(),
            summary: "Daily large deal report with updated pipeline data and forecasts.",
            status: .completed,
            fileSize: "2.4 MB",
            worksheetCount: 3,
            isFavorite: true
        ))
        .environmentObject(AppState())
    }
}
