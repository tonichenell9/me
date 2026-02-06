import SwiftUI

/// The sidebar view displaying filter options and categories.
/// Optimized for iPad's sidebar navigation pattern.
struct SidebarView: View {
    @EnvironmentObject var viewModel: NotesViewModel

    var body: some View {
        List(selection: $viewModel.selectedFilter) {
            // Smart Filters Section
            Section("Smart Filters") {
                sidebarRow(for: .all)
                sidebarRow(for: .favorites)
            }

            // Categories Section
            Section("Categories") {
                ForEach(Category.allCases) { category in
                    sidebarRow(for: .category(category))
                }
            }
        }
        .listStyle(.sidebar)
        .navigationTitle("Notes")
        .toolbar {
            ToolbarItem(placement: .primaryAction) {
                Button {
                    viewModel.createNote()
                } label: {
                    Image(systemName: "square.and.pencil")
                }
                .help("New Note (⌘N)")
            }

            ToolbarItem(placement: .secondaryAction) {
                sortMenu
            }
        }
    }

    // MARK: - Sidebar Row

    @ViewBuilder
    private func sidebarRow(for filter: SidebarFilter) -> some View {
        let count = viewModel.noteCounts[filter] ?? 0

        Label {
            HStack {
                Text(filter.title)
                Spacer()
                Text("\(count)")
                    .font(.caption)
                    .fontWeight(.medium)
                    .foregroundStyle(.secondary)
                    .padding(.horizontal, 8)
                    .padding(.vertical, 2)
                    .background(
                        Capsule()
                            .fill(Color(.systemGray5))
                    )
            }
        } icon: {
            Image(systemName: filter.iconName)
                .foregroundStyle(filter.color)
        }
        .tag(filter)
    }

    // MARK: - Sort Menu

    private var sortMenu: some View {
        Menu {
            ForEach(NotesViewModel.SortOrder.allCases, id: \.self) { order in
                Button {
                    viewModel.sortOrder = order
                } label: {
                    HStack {
                        Text(order.rawValue)
                        if viewModel.sortOrder == order {
                            Image(systemName: "checkmark")
                        }
                    }
                }
            }
        } label: {
            Image(systemName: "arrow.up.arrow.down")
        }
        .help("Sort Notes")
    }
}

// MARK: - Preview

#Preview {
    NavigationSplitView {
        SidebarView()
            .environmentObject(NotesViewModel())
    } content: {
        Text("Content")
    } detail: {
        Text("Detail")
    }
}
