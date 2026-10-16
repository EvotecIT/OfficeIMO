import CryptoKit
import Foundation

/// Foundation's document hosts serialize publication. Notifications can coalesce; readback acknowledges the observed
/// prepared bytes and retires earlier preparations, while retaining newer ones and any unobserved working revision.
final class StudioFilePublicationObserver: NSObject, NSFilePresenter, @unchecked Sendable {
    let presentedItemURL: URL?
    let presentedItemOperationQueue: OperationQueue
    private let engine: EngineDocument

    init(url: URL, engine: EngineDocument) {
        presentedItemURL = url
        self.engine = engine
        let queue = OperationQueue()
        queue.maxConcurrentOperationCount = 1
        queue.qualityOfService = .utility
        presentedItemOperationQueue = queue
        super.init()
        NSFileCoordinator.addFilePresenter(self)
        queue.addOperation { [weak self] in self?.presentedItemDidChange() }
    }

    func stop() { NSFileCoordinator.removeFilePresenter(self) }

    func presentedItemDidChange() {
        guard let url = presentedItemURL else { return }
        let accessing = url.startAccessingSecurityScopedResource()
        defer { if accessing { url.stopAccessingSecurityScopedResource() } }
        var coordinationError: NSError?
        NSFileCoordinator(filePresenter: self).coordinate(readingItemAt: url, options: .withoutChanges, error: &coordinationError) { original in
            // Failed or unavailable readback keeps recovery. It never acknowledges an unobserved publication.
            guard let fingerprint = Self.fingerprint(of: original) else { return }
            if let identity = StudioDocument.recoveryIdentity(for: url) {
                try? engine.acknowledgeOrderedSaveAndWait(fingerprint: fingerprint, identity: identity)
            }
        }
    }

    private static func fingerprint(of url: URL) -> String? {
        guard let stream = InputStream(url: url) else { return nil }
        stream.open()
        defer { stream.close() }
        var hash = SHA256()
        var buffer = [UInt8](repeating: 0, count: 64 * 1024)
        var length = 0
        while true {
            let count = stream.read(&buffer, maxLength: buffer.count)
            guard count >= 0 else { return nil }
            if count == 0 { break }
            length += count
            // An externally replaced, oversized file is not an acknowledgement of a Studio snapshot.
            guard length <= 512 * 1024 * 1024 else { return nil }
            hash.update(data: Data(buffer.prefix(count)))
        }
        return hash.finalize().map { String(format: "%02X", $0) }.joined()
    }
}
