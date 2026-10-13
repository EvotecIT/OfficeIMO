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
        let images = panel.urls.compactMap(ScanBuilder.image(at:))
        make(images, newDocument)
    }

    /// Makes a document from images Continuity Camera delivers.
    static func newDocument(from providers: [NSItemProvider], _ newDocument: NewDocumentAction) -> Bool {
        let imageProviders = providers.filter { $0.canLoadObject(ofClass: NSImage.self) }
        guard !imageProviders.isEmpty else { return false }
        Task {
            var images: [CGImage] = []
            for provider in imageProviders {
                if let image = await load(provider) { images.append(image) }
            }
            make(images, newDocument)
        }
        return true
    }

    private static func make(_ images: [CGImage], _ newDocument: NewDocumentAction) {
        guard !images.isEmpty else { return }
        Task {
            do {
                let url = try await ScanBuilder.makePDF(from: images, named: ScanBuilder.newName)
                let document = try StudioDocument(contentsOf: url)
                newDocument { document }
            } catch {
                NSAlert(error: error).runModal()
            }
        }
    }

    private static func load(_ provider: NSItemProvider) async -> CGImage? {
        await withCheckedContinuation { continuation in
            _ = provider.loadObject(ofClass: NSImage.self) { object, _ in
                continuation.resume(returning: (object as? NSImage)?.cgImage(forProposedRect: nil, context: nil, hints: nil))
            }
        }
    }
}
#endif
