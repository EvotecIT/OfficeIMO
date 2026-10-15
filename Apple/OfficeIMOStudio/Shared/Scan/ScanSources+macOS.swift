#if os(macOS)
import AppKit
import SwiftUI
import UniformTypeIdentifiers

/// Where a new scan's pages come from on the Mac: image files, or Continuity Camera's Import from iPhone or iPad,
/// which scans with the phone's document camera. Each makes a new untitled document, as Preview does.
@MainActor
enum ScanSources {
    /// Asks for image files and makes a searchable PDF from them, one page each in the order chosen.
    static func newDocumentFromImageFiles(_ newDocument: NewDocumentAction) {
        let panel = NSOpenPanel()
        panel.allowedContentTypes = [.image]
        panel.allowsMultipleSelection = true
        panel.message = String(localized: "Choose images for the pages of a new PDF.")
        panel.prompt = String(localized: "Create PDF")
        guard panel.runModal() == .OK else { return }
        let urls = panel.urls
        make(pageCount: urls.count, newDocument) { index in
            let url = urls[index]
            let accessing = url.startAccessingSecurityScopedResource()
            defer { if accessing { url.stopAccessingSecurityScopedResource() } }
            guard let image = ScanBuilder.image(at: url) else {
                throw CocoaError(.fileReadCorruptFile, userInfo: [NSFilePathErrorKey: url.path])
            }
            return image
        }
    }

    /// Makes a document from images Continuity Camera delivers.
    static func newDocument(from providers: [NSItemProvider], _ newDocument: NewDocumentAction) -> Bool {
        let imageProviders = providers.filter { $0.canLoadObject(ofClass: NSImage.self) }
        guard !imageProviders.isEmpty else { return false }
        make(pageCount: imageProviders.count, newDocument) { index in
            try await ScanBuilder.image(from: imageProviders[index])
        }
        return true
    }

    private static func make(pageCount: Int, _ newDocument: NewDocumentAction,
                             loadPage: @escaping @MainActor (Int) async throws -> CGImage) {
        guard pageCount > 0 else { return }
        Task {
            do {
                let url = try await ScanBuilder.makePDF(pageCount: pageCount, named: ScanBuilder.newName, loadPage: loadPage)
                defer { StudioTemporaryFiles.removeGeneratedPDF(url) }
                let document = try StudioDocument(contentsOf: url)
                newDocument { document }
            } catch {
                NSAlert(error: error).runModal()
            }
        }
    }
}
#endif
