import Foundation

extension Date {
    /// Returns a human-readable relative time string (e.g., "2 hours ago", "Yesterday")
    var relativeDescription: String {
        let formatter = RelativeDateTimeFormatter()
        formatter.unitsStyle = .full
        return formatter.localizedString(for: self, relativeTo: Date())
    }

    /// Returns the date formatted as the automation system expects (YYYY-MM-DD)
    var automationDateString: String {
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyy-MM-dd"
        return formatter.string(from: self)
    }

    /// Returns a compact time string (e.g., "9:30 AM")
    var timeString: String {
        let formatter = DateFormatter()
        formatter.dateFormat = "h:mm a"
        return formatter.string(from: self)
    }

    /// Returns the start of the current day
    var startOfDay: Date {
        Calendar.current.startOfDay(for: self)
    }

    /// Returns true if this date is within the last N days
    func isWithinLast(days: Int) -> Bool {
        guard let startDate = Calendar.current.date(byAdding: .day, value: -days, to: Date()) else {
            return false
        }
        return self >= startDate
    }
}
