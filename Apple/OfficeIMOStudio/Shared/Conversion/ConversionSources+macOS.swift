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
                let panel = NSOpenPanel()
                panel.allowedContentTypes = try await DocumentConversion.contentTypes()
                panel.message = String(localized: "Choose a document to convert to PDF.")
                panel.prompt = String(localized: "Convert")
                guard panel.runModal() == .OK, let input = panel.url else { return }
                let converted = try await DocumentConversion.convert(input)
                let document = try StudioDocument(contentsOf: converted.url)
                newDocument { document }
            } catch {
                NSAlert(error: error).runModal()
            }
        }
    }
}
#endif
