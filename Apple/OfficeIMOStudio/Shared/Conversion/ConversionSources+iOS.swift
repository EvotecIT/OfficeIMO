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
        guard let input = await DocumentPickerSession(types: types, allowsMultiple: false).run().first else { return nil }
        let output = try await DocumentConversion.convert(input).url
        StudioTemporaryFiles.markReturnedOutput(output)
        return output
    }

    /// One new PDF from the files the user picks, for Create PDF's Combine Files.
    static func newDocumentFromFiles() async throws -> URL? {
        let inputs = await DocumentPickerSession(types: PageSources.contentTypes, allowsMultiple: true).run()
        guard !inputs.isEmpty else { return nil }
        let output = try await DocumentConversion.combine(inputs)
        StudioTemporaryFiles.markReturnedOutput(output)
        return output
    }
}

@MainActor
private final class DocumentPickerSession: NSObject, UIDocumentPickerDelegate {
    private let types: [UTType]
    private let allowsMultiple: Bool
    private var continuation: CheckedContinuation<[URL], Never>?
    private var keepAlive: DocumentPickerSession?

    init(types: [UTType], allowsMultiple: Bool) {
        self.types = types
        self.allowsMultiple = allowsMultiple
    }

    func run() async -> [URL] {
        await withCheckedContinuation { continuation in
            self.continuation = continuation
            // A copy, so the conversion reads a file the app owns.
            let picker = UIDocumentPickerViewController(forOpeningContentTypes: types, asCopy: true)
            picker.delegate = self
            picker.allowsMultipleSelection = allowsMultiple
            keepAlive = self
            Task { if !(await ScanSources.present(picker)) { self.finish([]) } }
        }
    }

    func documentPicker(_ controller: UIDocumentPickerViewController, didPickDocumentsAt urls: [URL]) {
        finish(urls)
    }

    func documentPickerWasCancelled(_ controller: UIDocumentPickerViewController) {
        finish([])
    }

    private func finish(_ urls: [URL]) {
        continuation?.resume(returning: urls)
        continuation = nil
        keepAlive = nil
    }
}
#endif
