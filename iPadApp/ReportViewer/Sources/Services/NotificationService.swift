import Foundation
import UserNotifications

/// Manages local notifications for report status updates.
final class NotificationService {

    static let shared = NotificationService()

    private init() {}

    // MARK: - Authorization

    /// Request notification permissions from the user
    func requestAuthorization() async -> Bool {
        do {
            let granted = try await UNUserNotificationCenter.current()
                .requestAuthorization(options: [.alert, .badge, .sound])
            return granted
        } catch {
            print("Notification authorization error: \(error)")
            return false
        }
    }

    // MARK: - Scheduling

    /// Schedule a notification for when a new report is available
    func scheduleReportReadyNotification(for report: Report) {
        let content = UNMutableNotificationContent()
        content.title = "Report Ready"
        content.body = "\(report.title) for \(report.formattedDate) is now available."
        content.sound = .default
        content.categoryIdentifier = "REPORT_READY"

        let trigger = UNTimeIntervalNotificationTrigger(timeInterval: 1, repeats: false)
        let request = UNNotificationRequest(
            identifier: "report-\(report.id)",
            content: content,
            trigger: trigger
        )

        UNUserNotificationCenter.current().add(request) { error in
            if let error = error {
                print("Failed to schedule notification: \(error)")
            }
        }
    }

    /// Schedule a daily reminder to check for new reports
    func scheduleDailyReminder(at hour: Int, minute: Int) {
        let content = UNMutableNotificationContent()
        content.title = "Report Check"
        content.body = "Time to check for today's Large Deal Report."
        content.sound = .default

        var dateComponents = DateComponents()
        dateComponents.hour = hour
        dateComponents.minute = minute

        let trigger = UNCalendarNotificationTrigger(dateMatching: dateComponents, repeats: true)
        let request = UNNotificationRequest(
            identifier: "daily-reminder",
            content: content,
            trigger: trigger
        )

        UNUserNotificationCenter.current().add(request) { error in
            if let error = error {
                print("Failed to schedule daily reminder: \(error)")
            }
        }
    }

    /// Cancel all pending notifications
    func cancelAllNotifications() {
        UNUserNotificationCenter.current().removeAllPendingNotificationRequests()
    }

    // MARK: - Notification Actions

    /// Register notification categories and actions
    func registerNotificationCategories() {
        let viewAction = UNNotificationAction(
            identifier: "VIEW_REPORT",
            title: "View Report",
            options: .foreground
        )

        let shareAction = UNNotificationAction(
            identifier: "SHARE_REPORT",
            title: "Share",
            options: .foreground
        )

        let reportCategory = UNNotificationCategory(
            identifier: "REPORT_READY",
            actions: [viewAction, shareAction],
            intentIdentifiers: [],
            options: .customDismissAction
        )

        UNUserNotificationCenter.current().setNotificationCategories([reportCategory])
    }
}
