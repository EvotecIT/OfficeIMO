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
