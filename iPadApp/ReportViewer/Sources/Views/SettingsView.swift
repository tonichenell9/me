import SwiftUI

/// Settings view for configuring server connection, notifications, and app preferences.
/// Maps to the same configuration options as the Python automation system's config.json.
struct SettingsView: View {
    @EnvironmentObject var appState: AppState
    @State private var serverURL: String = ""
    @State private var emailUsername: String = ""
    @State private var notificationsEnabled: Bool = true
    @State private var dailyReminderEnabled: Bool = false
    @State private var reminderHour: Int = 9
    @State private var reminderMinute: Int = 0
    @State private var autoRefreshEnabled: Bool = true
    @State private var refreshInterval: Double = 30 // minutes
    @State private var showingResetConfirmation: Bool = false
    @State private var showingSavedAlert: Bool = false

    var body: some View {
        Form {
            // Server Configuration
            Section {
                TextField("Server URL", text: $serverURL)
                    .textContentType(.URL)
                    .keyboardType(.URL)
                    .autocapitalization(.none)
                    .disableAutocorrection(true)

                TextField("Email (optional)", text: $emailUsername)
                    .textContentType(.emailAddress)
                    .keyboardType(.emailAddress)
                    .autocapitalization(.none)
            } header: {
                Text("Server Connection")
            } footer: {
                Text("Enter the URL of your report automation server. Leave blank to use sample data for demonstration.")
            }

            // Notifications
            Section {
                Toggle("Enable Notifications", isOn: $notificationsEnabled)

                if notificationsEnabled {
                    Toggle("Daily Reminder", isOn: $dailyReminderEnabled)

                    if dailyReminderEnabled {
                        HStack {
                            Text("Reminder Time")
                            Spacer()
                            Text("\(String(format: "%02d", reminderHour)):\(String(format: "%02d", reminderMinute))")
                                .foregroundStyle(.secondary)
                        }

                        Stepper("Hour: \(reminderHour)", value: $reminderHour, in: 0...23)
                        Stepper("Minute: \(reminderMinute)", value: $reminderMinute, in: 0...59, step: 15)
                    }
                }
            } header: {
                Text("Notifications")
            } footer: {
                Text("Get notified when new reports are ready or when report generation fails.")
            }

            // Auto Refresh
            Section {
                Toggle("Auto Refresh", isOn: $autoRefreshEnabled)

                if autoRefreshEnabled {
                    VStack(alignment: .leading) {
                        Text("Refresh every \(Int(refreshInterval)) minutes")
                        Slider(value: $refreshInterval, in: 5...120, step: 5)
                    }
                }
            } header: {
                Text("Data Refresh")
            }

            // App Information
            Section {
                HStack {
                    Text("Version")
                    Spacer()
                    Text("1.0.0")
                        .foregroundStyle(.secondary)
                }
                HStack {
                    Text("Build")
                    Spacer()
                    Text("1")
                        .foregroundStyle(.secondary)
                }
                HStack {
                    Text("Platform")
                    Spacer()
                    Text("iPadOS")
                        .foregroundStyle(.secondary)
                }
                if let lastSync = appState.lastSyncDate {
                    HStack {
                        Text("Last Sync")
                        Spacer()
                        Text(lastSync.formatted())
                            .foregroundStyle(.secondary)
                    }
                }
            } header: {
                Text("About")
            }

            // Actions
            Section {
                Button("Save Settings") {
                    saveSettings()
                }
                .fontWeight(.semibold)

                Button("Reset to Defaults", role: .destructive) {
                    showingResetConfirmation = true
                }
            }
        }
        .formStyle(.grouped)
        .navigationTitle("Settings")
        .onAppear {
            serverURL = appState.serverURL
        }
        .alert("Settings Saved", isPresented: $showingSavedAlert) {
            Button("OK", role: .cancel) {}
        } message: {
            Text("Your settings have been saved successfully.")
        }
        .confirmationDialog(
            "Reset Settings",
            isPresented: $showingResetConfirmation,
            titleVisibility: .visible
        ) {
            Button("Reset", role: .destructive) {
                resetSettings()
            }
            Button("Cancel", role: .cancel) {}
        } message: {
            Text("This will reset all settings to their default values.")
        }
    }

    private func saveSettings() {
        appState.serverURL = serverURL
        appState.saveSettings()

        if notificationsEnabled {
            Task {
                _ = await NotificationService.shared.requestAuthorization()
                if dailyReminderEnabled {
                    NotificationService.shared.scheduleDailyReminder(
                        at: reminderHour,
                        minute: reminderMinute
                    )
                }
            }
        } else {
            NotificationService.shared.cancelAllNotifications()
        }

        showingSavedAlert = true
    }

    private func resetSettings() {
        serverURL = ""
        emailUsername = ""
        notificationsEnabled = true
        dailyReminderEnabled = false
        reminderHour = 9
        reminderMinute = 0
        autoRefreshEnabled = true
        refreshInterval = 30

        appState.serverURL = ""
        appState.saveSettings()
    }
}

#Preview {
    NavigationStack {
        SettingsView()
            .environmentObject(AppState())
    }
}
