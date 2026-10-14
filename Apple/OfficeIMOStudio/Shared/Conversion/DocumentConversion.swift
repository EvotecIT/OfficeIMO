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
        let folder = StudioTemporaryFiles.root.appending(path: "Converted/\(UUID().uuidString)", directoryHint: .isDirectory)
        try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
        let output = folder.appending(path: input.deletingPathExtension().lastPathComponent + ".pdf")
        let accessing = input.startAccessingSecurityScopedResource()
        defer { if accessing { input.stopAccessingSecurityScopedResource() } }
        do {
            let result = try await Engine.convertToPDF(input, at: output)
            return (output, result)
        } catch {
            try? FileManager.default.removeItem(at: folder)
            throw error
        }
    }
}

/// Files that become pages, when inserting pages or combining files: a PDF as it is, an image as a page with its
/// text recognized on this device, and a Word, Excel, PowerPoint or other document converted by the engine.
enum PageSources {
    /// The extensions of the documents the engine converts, known without asking it, for Shortcuts' fixed file types.
    static let documentExtensions = ["doc", "docx", "docm", "dotx", "dotm", "xlsx", "xlsm", "xltx", "xltm", "pptx", "pptm", "potx",
                                     "potm", "ppsx", "ppsm", "rtf", "txt", "md", "markdown", "html", "htm", "xps", "oxps", "vsdx",
                                     "vdx", "odg", "fodg", "pub", "chm", "djvu", "djv"]

    /// PDFs, images and the documents above, as file types.
    static var contentTypes: [UTType] {
        [.pdf, .image] + DocumentConversion.types(for: Set(documentExtensions))
    }

    /// A PDF of the file: the file itself, or a new PDF in the app's temporary folder.
    static func pdf(for url: URL) async throws -> URL {
        let name = url.lastPathComponent
        let type = UTType(filenameExtension: url.pathExtension)
        if type?.conforms(to: .pdf) == true { return url }
        if type?.conforms(to: .image) == true {
            let accessing = url.startAccessingSecurityScopedResource()
            defer { if accessing { url.stopAccessingSecurityScopedResource() } }
            guard let image = ScanBuilder.image(from: try Data(contentsOf: url)) else {
                throw PageSourceProblem(message: String(localized: "“\(name)” isn’t an image OfficeIMO Studio can read."))
            }
            return try await ScanBuilder.makePDF(from: [image], named: url.deletingPathExtension().lastPathComponent)
        }
        do {
            return try await DocumentConversion.convert(url).url
        } catch let error as EngineError where error.status == .unsupported {
            throw PageSourceProblem(message: String(localized: "“\(name)” isn’t a format OfficeIMO Studio converts to PDF."))
        }
    }
}

/// A file that can't become pages, said in the app's windows and in Shortcuts.
struct PageSourceProblem: LocalizedError, CustomLocalizedStringResourceConvertible {
    let message: String

    var errorDescription: String? { message }
    var localizedStringResource: LocalizedStringResource { "\(message)" }
}
