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
                let converted = try await DocumentConversion.convert(input)
                let document = try StudioDocument(contentsOf: converted.url)
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
