import Foundation
import PDFKit
import SwiftUI
import UniformTypeIdentifiers

/// A Word, Excel, PowerPoint or other document opened to view. Studio edits PDFs, so the file is shown as the engine's
/// PDF conversion and stays as it is; Edit as PDF and Save as PDF make the PDF a document of its own. A separate,
/// read-only document type keeps these windows apart from PDF windows, whose titles and saves belong to a PDF.
struct ViewerDocument: FileDocument {
    static let readableContentTypes: [UTType] = ViewerFormats.identifiers.map { UTType(importedAs: $0) }
    static let writableContentTypes: [UTType] = []

    /// The app's copy of the opened file, which the engine converts.
    let source: URL
    let name: String
    let problem: String?

    init(configuration: ReadConfiguration) throws {
        guard let data = configuration.file.regularFileContents else { throw CocoaError(.fileReadCorruptFile) }
        try self.init(data: data, name: configuration.file.filename ?? String(localized: "Document"))
    }

    /// Opens a file the app can read directly, without the system's document browser.
    init(contentsOf url: URL) throws {
        try self.init(data: Data(contentsOf: url, options: .mappedIfSafe), name: url.lastPathComponent)
    }

    private init(data: Data, name: String) throws {
        self.name = name
        let folder = StudioTemporaryFiles.root.appending(path: "Viewing/\(UUID().uuidString)", directoryHint: .isDirectory)
        source = folder.appending(path: try StudioTemporaryFiles.outputName(name))
        problem = ViewerFormats.sizeProblem(byteCount: Int64(data.count), name: name)
        guard problem == nil else { return }
        try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
        try data.write(to: source)
    }

    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper {
        throw CocoaError(.featureUnsupported)
    }
}

/// The converted pages of one viewed document, made once per window.
@MainActor @Observable
final class ViewerModel {
    enum Phase {
        case converting
        case ready(PDFDocument, pdf: URL, differs: Bool)
        case failed(String)
    }

    private(set) var phase = Phase.converting
    let document: ViewerDocument
    private var isClosed = false

    init(document: ViewerDocument) {
        self.document = document
    }

    var pdf: URL? {
        if case .ready(_, let pdf, _) = phase { pdf } else { nil }
    }

    /// The PDF's name: the document's name with a .pdf extension.
    var pdfName: String { (document.name as NSString).deletingPathExtension }

    func load() async {
        if let problem = document.problem {
            phase = .failed(problem)
            return
        }
        guard case .converting = phase else { return }
        do {
            try await EngineRuntime.ready()
            let converted = try await DocumentConversion.convert(document.source)
            // The engine can finish after the window closed; its PDF then has no window to remove it.
            guard !isClosed else {
                StudioTemporaryFiles.removeGeneratedPDF(converted.url)
                return
            }
            guard let pages = PDFDocument(url: converted.url), pages.pageCount > 0 else { throw CocoaError(.fileReadCorruptFile) }
            phase = .ready(pages, pdf: converted.url, differs: !converted.result.notes.isEmpty)
        } catch is CancellationError {
            return
        } catch {
            phase = .failed(String(localized: "OfficeIMO Studio couldn’t show “\(document.name)”. \(error.localizedDescription)"))
        }
    }

    /// Removes the conversion and the copy of the opened file when the window closes.
    func close() {
        isClosed = true
        if let pdf { StudioTemporaryFiles.removeGeneratedPDF(pdf) }
        try? FileManager.default.removeItem(at: document.source.deletingLastPathComponent())
    }
}

/// A converted PDF to save with the system's export panel.
struct ConvertedPDF: FileDocument {
    static let readableContentTypes: [UTType] = [.pdf]
    let url: URL

    init(url: URL) {
        self.url = url
    }

    init(configuration: ReadConfiguration) throws {
        throw CocoaError(.featureUnsupported)
    }

    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper {
        try FileWrapper(url: url, options: .immediate)
    }
}
