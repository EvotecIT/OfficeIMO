#if os(iOS)
import UIKit

/// The launch screen's actions that start from something other than a blank page. A NewDocumentButton makes a default
/// document without running its prepare action on iOS 27, so these buttons make the PDF themselves, keep it with the
/// app's documents and open it through the document browser, as if the user had picked it there.
@MainActor
enum LaunchDocuments {
    static func open(_ prepare: @escaping @MainActor () async throws -> URL?) {
        Task {
            do {
                guard let output = try await prepare() else { return }
                defer { StudioTemporaryFiles.removeGeneratedPDF(output) }
                guard let browser = documentBrowser() else { throw CocoaError(.featureUnsupported) }
                browser.delegate?.documentBrowser?(browser, didPickDocumentsAt: [try keep(output)])
            } catch {
                let alert = UIAlertController(title: nil, message: error.localizedDescription, preferredStyle: .alert)
                alert.addAction(UIAlertAction(title: String(localized: "OK"), style: .default))
                _ = await ScanSources.present(alert)
            }
        }
    }

    /// A copy in the app's Documents folder, under a name not yet taken there.
    private static func keep(_ output: URL) throws -> URL {
        let folder = URL.documentsDirectory
        let name = output.deletingPathExtension().lastPathComponent
        var kept = folder.appending(path: output.lastPathComponent)
        var number = 2
        while FileManager.default.fileExists(atPath: kept.path) {
            kept = folder.appending(path: "\(name) \(number).pdf")
            number += 1
        }
        try FileManager.default.copyItem(at: output, to: kept)
        return kept
    }

    private static func documentBrowser() -> UIDocumentBrowserViewController? {
        let roots = UIApplication.shared.connectedScenes.compactMap { $0 as? UIWindowScene }.flatMap(\.windows).compactMap(\.rootViewController)
        var pending = roots
        while let controller = pending.popLast() {
            if let browser = controller as? UIDocumentBrowserViewController { return browser }
            pending += controller.children
            if let presented = controller.presentedViewController { pending.append(presented) }
        }
        return nil
    }
}
#endif
