import Foundation

/// Configuration model matching the Python automation system's config.json format.
/// This allows the iPad app to connect to the same infrastructure.
struct ServerConfig: Codable {
    var email: EmailConfig
    var distributionList: [String]
    var senderName: String
    var reportsDirectory: String
    var tempDirectory: String
    var worksheet1Name: String
    var worksheet2Name: String
    var worksheet3Name: String
    var logLevel: String
    var logFile: String?

    enum CodingKeys: String, CodingKey {
        case email
        case distributionList = "distribution_list"
        case senderName = "sender_name"
        case reportsDirectory = "reports_directory"
        case tempDirectory = "temp_directory"
        case worksheet1Name = "worksheet_1_name"
        case worksheet2Name = "worksheet_2_name"
        case worksheet3Name = "worksheet_3_name"
        case logLevel = "log_level"
        case logFile = "log_file"
    }

    static var `default`: ServerConfig {
        ServerConfig(
            email: .default,
            distributionList: [],
            senderName: "",
            reportsDirectory: "./reports",
            tempDirectory: "./temp",
            worksheet1Name: "Sheet1",
            worksheet2Name: "Sheet2",
            worksheet3Name: "Sheet3",
            logLevel: "INFO",
            logFile: nil
        )
    }
}

/// Email configuration matching the Python automation system
struct EmailConfig: Codable {
    var username: String
    var password: String
    var senderEmail: String
    var imapServer: String
    var imapPort: Int
    var smtpServer: String
    var smtpPort: Int
    var useOutlookForSending: Bool?

    enum CodingKeys: String, CodingKey {
        case username
        case password
        case senderEmail = "sender_email"
        case imapServer = "imap_server"
        case imapPort = "imap_port"
        case smtpServer = "smtp_server"
        case smtpPort = "smtp_port"
        case useOutlookForSending = "use_outlook_for_sending"
    }

    static var `default`: EmailConfig {
        EmailConfig(
            username: "",
            password: "",
            senderEmail: "",
            imapServer: "imap.gmail.com",
            imapPort: 993,
            smtpServer: "smtp.gmail.com",
            smtpPort: 587,
            useOutlookForSending: false
        )
    }
}
