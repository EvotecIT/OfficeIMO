import Foundation
import SwiftUI

/// A short PDF with text to search and a form to fill, for trying the app with nothing of one's own to hand.
enum SampleDocument {
    static var url: URL? { Bundle.main.url(forResource: "OfficeIMO Studio Sample", withExtension: "pdf") }

    #if os(macOS)
    @MainActor
    static func open(_ newDocument: NewDocumentAction) {
        guard let url, let document = try? StudioDocument(contentsOf: url) else { return }
        newDocument { document }
    }
    #endif
}

#if os(macOS)
/// The window of an opened Word, Excel, PowerPoint or RTF file: Studio edits PDFs, so the file becomes a new PDF that
/// opens in its place, and the original stays as it is. The iOS document browser keeps stale titles after a window
/// like this closes, so there these files go through Convert to PDF.
struct OfficeDocumentView: View {
    let source: URL

    @Environment(\.dismiss) private var dismiss
    @Environment(\.newDocument) private var newDocument

    var body: some View {
        ProgressView("Converting to PDF…")
            .frame(maxWidth: .infinity, maxHeight: .infinity)
            .task { ConversionSources.convert(source, newDocument) { dismiss() } }
    }
}
#endif
