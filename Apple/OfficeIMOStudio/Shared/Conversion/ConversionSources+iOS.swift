#if os(iOS)
import UIKit
import UniformTypeIdentifiers

/// Convert to PDF on iPad and iPhone: the system document picker offers the formats the engine converts, and the
/// converted PDF opens as a new document.
@MainActor
enum ConversionSources {
    /// A new PDF converted from a document the user picks, for the launch screen's Convert to PDF.
    static func newDocumentFromFile() async throws -> URL? {
        let types = try await DocumentConversion.contentTypes()
        guard let input = await DocumentPickerSession(types: types).run() else { return nil }
        return try await DocumentConversion.convert(input).url
    }
}

@MainActor
private final class DocumentPickerSession: NSObject, UIDocumentPickerDelegate {
    private let types: [UTType]
    private var continuation: CheckedContinuation<URL?, Never>?
    private var keepAlive: DocumentPickerSession?

    init(types: [UTType]) {
        self.types = types
    }

    func run() async -> URL? {
        await withCheckedContinuation { continuation in
            self.continuation = continuation
            // A copy, so the conversion reads a file the app owns.
            let picker = UIDocumentPickerViewController(forOpeningContentTypes: types, asCopy: true)
            picker.delegate = self
            picker.allowsMultipleSelection = false
            keepAlive = self
            Task { if !(await ScanSources.present(picker)) { self.finish(nil) } }
        }
    }

    func documentPicker(_ controller: UIDocumentPickerViewController, didPickDocumentsAt urls: [URL]) {
        finish(urls.first)
    }

    func documentPickerWasCancelled(_ controller: UIDocumentPickerViewController) {
        finish(nil)
    }

    private func finish(_ url: URL?) {
        continuation?.resume(returning: url)
        continuation = nil
        keepAlive = nil
    }
}
#endif
