import Foundation
import UniformTypeIdentifiers

/// A format that converts to PDF, from officeimo_conversion_sources_json.
struct ConversionSource: Decodable, Sendable {
    let route: String
    let name: String
    let extensions: [String]
}

struct ConversionSourceList: Decodable, Sendable {
    let sources: [ConversionSource]
}

/// A finished conversion and the converter's notes about content the PDF shows differently.
struct ConversionResult: Decodable, Sendable {
    struct Note: Decodable, Sendable {
        let severity: String
        let message: String
    }

    let route: String
    let source: String
    let outputBytes: Int64
    let notes: [Note]
}

/// Converting Word, Excel, PowerPoint and other documents to PDF with the engine, as OfficeIMO Studio's Convert does
/// on every platform. The result is a new PDF beside the app's other temporary documents.
enum DocumentConversion {
    /// The file types the engine converts, for the system's file pickers.
    static func contentTypes() async throws -> [UTType] {
        types(for: try await extensions())
    }

    /// The file extensions the engine converts, lowercased and without the dot. A picker filtering by type also offers
    /// types that conform to these, such as CSV as plain text, so a picker that can, checks the extension too.
    static func extensions() async throws -> Set<String> {
        Set(try await Engine.conversionSources().flatMap(\.extensions).map { String($0.drop { $0 == "." }).lowercased() })
    }

    static func types(for extensions: Set<String>) -> [UTType] {
        var seen = Set<String>()
        return extensions.sorted().compactMap { UTType(filenameExtension: $0) }.filter { seen.insert($0.identifier).inserted }
    }

    /// Converts a document to a new PDF named after it.
    static func convert(_ input: URL) async throws -> (url: URL, result: ConversionResult) {
        let folder = FileManager.default.temporaryDirectory.appending(path: "Converted/\(UUID().uuidString)", directoryHint: .isDirectory)
        try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
        let output = folder.appending(path: input.deletingPathExtension().lastPathComponent + ".pdf")
        let accessing = input.startAccessingSecurityScopedResource()
        defer { if accessing { input.stopAccessingSecurityScopedResource() } }
        let result = try await Engine.convertToPDF(input, at: output)
        return (output, result)
    }
}
