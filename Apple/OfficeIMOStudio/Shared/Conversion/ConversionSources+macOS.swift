#if os(macOS)
import AppKit
import SwiftUI

/// File > Convert to PDF on the Mac: an open panel offers the formats the engine converts, and the converted PDF opens
/// as a new untitled document, as Preview opens the PDFs it makes.
@MainActor
enum ConversionSources {
    static func newDocumentFromFile(_ newDocument: NewDocumentAction) {
        Task {
            do {
                let extensions = try await DocumentConversion.extensions()
                let panel = NSOpenPanel()
                panel.allowedContentTypes = DocumentConversion.types(for: extensions)
                let filter = ExtensionFilter(extensions)
                panel.delegate = filter
                panel.message = String(localized: "Choose a document to convert to PDF.")
                panel.prompt = String(localized: "Convert")
                let response = panel.runModal()
                withExtendedLifetime(filter) {}
                guard response == .OK, let input = panel.url else { return }
                convert(input, newDocument)
            } catch {
                NSAlert(error: error).runModal()
            }
        }
    }

    static func convert(_ input: URL, _ newDocument: NewDocumentAction, then finish: @escaping @MainActor () -> Void = {}) {
        open(newDocument, then: finish) { try await DocumentConversion.convert(input).url }
    }

    /// File > Combine Files on the Mac: one new PDF from several files, opened as an untitled document.
    static func newDocumentFromFiles(_ newDocument: NewDocumentAction) {
        let panel = NSOpenPanel()
        panel.allowedContentTypes = PageSources.contentTypes
        panel.allowsMultipleSelection = true
        panel.message = String(localized: "Choose the files to combine into one PDF.")
        panel.prompt = String(localized: "Combine")
        guard panel.runModal() == .OK else { return }
        open(newDocument) { try await DocumentConversion.combine(panel.urls) }
    }

    private static func open(_ newDocument: NewDocumentAction, then finish: @escaping @MainActor () -> Void = {},
                             _ prepare: @escaping @MainActor () async throws -> URL) {
        Task {
            defer { finish() }
            do {
                let output = try await prepare()
                defer { StudioTemporaryFiles.removeGeneratedPDF(output) }
                let document = try StudioDocument(contentsOf: output)
                newDocument { document }
            } catch {
                NSAlert(error: error).runModal()
            }
        }
    }
}

/// Enables only files whose extension the engine converts, and folders to browse.
private final class ExtensionFilter: NSObject, NSOpenSavePanelDelegate {
    private let extensions: Set<String>

    init(_ extensions: Set<String>) {
        self.extensions = extensions
    }

    func panel(_ sender: Any, shouldEnable url: URL) -> Bool {
        url.hasDirectoryPath || extensions.contains(url.pathExtension.lowercased())
    }
}
#endif
