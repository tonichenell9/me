import Foundation

/// Service responsible for fetching reports from the server,
/// managing local caching, and handling file downloads.
final class ReportService {

    // MARK: - Properties

    private let session: URLSession
    private let decoder: JSONDecoder
    private let fileManager = FileManager.default

    // MARK: - Initialization

    init() {
        let config = URLSessionConfiguration.default
        config.timeoutIntervalForRequest = 30
        config.timeoutIntervalForResource = 120
        config.waitsForConnectivity = true
        self.session = URLSession(configuration: config)

        self.decoder = JSONDecoder()
        self.decoder.dateDecodingStrategy = .iso8601
    }

    // MARK: - Public Methods

    /// Fetch all reports from the server.
    /// Falls back to local sample data if server is unavailable.
    func fetchReports(from serverURL: String) async throws -> [Report] {
        guard !serverURL.isEmpty, let url = URL(string: "\(serverURL)/api/reports") else {
            // Return sample data when no server is configured
            return generateSampleReports()
        }

        var request = URLRequest(url: url)
        request.httpMethod = "GET"
        request.setValue("application/json", forHTTPHeaderField: "Accept")

        do {
            let (data, response) = try await session.data(for: request)

            guard let httpResponse = response as? HTTPURLResponse else {
                throw ReportServiceError.invalidResponse
            }

            guard (200...299).contains(httpResponse.statusCode) else {
                throw ReportServiceError.serverError(statusCode: httpResponse.statusCode)
            }

            let reports = try decoder.decode([Report].self, from: data)
            return reports
        } catch is DecodingError {
            throw ReportServiceError.decodingFailed
        } catch let error as ReportServiceError {
            throw error
        } catch {
            throw ReportServiceError.networkError(error)
        }
    }

    /// Download a specific report file to the local documents directory.
    func downloadReport(_ report: Report, from serverURL: String) async throws -> URL {
        guard !serverURL.isEmpty,
              let url = URL(string: "\(serverURL)/api/reports/\(report.id)/download") else {
            throw ReportServiceError.invalidURL
        }

        let (tempURL, response) = try await session.download(from: url)

        guard let httpResponse = response as? HTTPURLResponse,
              (200...299).contains(httpResponse.statusCode) else {
            throw ReportServiceError.downloadFailed
        }

        // Move to documents directory
        let documentsURL = fileManager.urls(for: .documentDirectory, in: .userDomainMask).first!
        let reportsDir = documentsURL.appendingPathComponent("Reports", isDirectory: true)

        try fileManager.createDirectory(at: reportsDir, withIntermediateDirectories: true)

        let destinationURL = reportsDir.appendingPathComponent(report.filename)

        if fileManager.fileExists(atPath: destinationURL.path) {
            try fileManager.removeItem(at: destinationURL)
        }

        try fileManager.moveItem(at: tempURL, to: destinationURL)
        return destinationURL
    }

    /// Get locally cached reports
    func getCachedReports() -> [URL] {
        let documentsURL = fileManager.urls(for: .documentDirectory, in: .userDomainMask).first!
        let reportsDir = documentsURL.appendingPathComponent("Reports", isDirectory: true)

        guard let files = try? fileManager.contentsOfDirectory(
            at: reportsDir,
            includingPropertiesForKeys: [.creationDateKey, .fileSizeKey],
            options: .skipsHiddenFiles
        ) else {
            return []
        }

        return files.filter { $0.pathExtension == "xlsx" }
            .sorted { $0.lastPathComponent > $1.lastPathComponent }
    }

    /// Delete a locally cached report
    func deleteCachedReport(at url: URL) throws {
        try fileManager.removeItem(at: url)
    }

    // MARK: - Private Methods

    /// Generate sample reports for demonstration and offline use
    private func generateSampleReports() -> [Report] {
        let calendar = Calendar.current
        let today = Date()

        return (0..<14).map { daysAgo in
            let date = calendar.date(byAdding: .day, value: -daysAgo, to: today) ?? today
            let status: ReportStatus = daysAgo == 0 ? .inProgress : .completed
            let size = String(format: "%.1f MB", Double.random(in: 1.8...3.2))

            return Report(
                title: "Large Deal Report",
                date: date,
                summary: "Daily large deal report for \(formatDate(date)).",
                status: status,
                fileSize: size,
                worksheetCount: 3,
                isFavorite: daysAgo == 0 || daysAgo == 1
            )
        }
    }

    private func formatDate(_ date: Date) -> String {
        let formatter = DateFormatter()
        formatter.dateStyle = .medium
        return formatter.string(from: date)
    }
}

// MARK: - Errors

enum ReportServiceError: LocalizedError {
    case invalidURL
    case invalidResponse
    case serverError(statusCode: Int)
    case decodingFailed
    case downloadFailed
    case networkError(Error)

    var errorDescription: String? {
        switch self {
        case .invalidURL:
            return "Invalid server URL. Please check your settings."
        case .invalidResponse:
            return "Received an invalid response from the server."
        case .serverError(let code):
            return "Server returned error code \(code)."
        case .decodingFailed:
            return "Failed to parse server response."
        case .downloadFailed:
            return "Failed to download the report file."
        case .networkError(let error):
            return "Network error: \(error.localizedDescription)"
        }
    }
}
