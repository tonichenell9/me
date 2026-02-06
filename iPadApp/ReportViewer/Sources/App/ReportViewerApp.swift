import SwiftUI

/// Main entry point for the Report Viewer iPad app.
/// This app serves as a companion to the Large Deal Report Automation system,
/// allowing users to view, manage, and share reports directly on their iPad.
@main
struct ReportViewerApp: App {
    @StateObject private var appState = AppState()

    var body: some Scene {
        WindowGroup {
            ContentView()
                .environmentObject(appState)
        }
    }
}
