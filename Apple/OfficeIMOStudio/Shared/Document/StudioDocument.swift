import CoreGraphics
import CryptoKit
import Foundation
import Synchronization
import SwiftUI
import UniformTypeIdentifiers

/// A PDF open in a window. The system owns the file: it reads it, autosaves, keeps versions and moves it. The engine
/// works on a private working copy, so it never needs access to the user's file, and each save writes the engine's
/// result back through the system.
final class StudioDocument: ReferenceFileDocument, @unchecked Sendable {
    static let readableContentTypes: [UTType] = [.pdf]

    /// The engine's copy of the document, inside the app's temporary folder.
    let workingCopy: URL
    private let opened = Mutex<EngineDocument?>(nil)
    // A replacement page view must not overwrite the immutable draft left by a failed commit in its predecessor.
    private let pendingDrawings = Mutex<[UUID: [AnnotationRequest]]>([:])
    private let snapshotPreparations = Mutex<[UUID: @MainActor @Sendable () throws -> Void]>([:])
    private let publicationObserver = Mutex<StudioFilePublicationObserver?>(nil)
    private let initializationError: Error?

    /// A blank page, for the system's New command: Letter where it is the common paper size, otherwise A4.
    init() {
        workingCopy = Self.makeWorkingCopyURL(named: "Untitled.pdf")
        let letter = ["US", "CA", "MX", "PH", "CL", "CO", "VE"].contains(Locale.current.region?.identifier ?? "")
        var page = letter ? CGRect(x: 0, y: 0, width: 612, height: 792) : CGRect(x: 0, y: 0, width: 595.28, height: 841.89)
        let data = NSMutableData()
        if let consumer = CGDataConsumer(data: data as CFMutableData), let context = CGContext(consumer: consumer, mediaBox: &page, nil) {
            context.beginPDFPage(nil)
            context.endPDFPage()
            context.closePDF()
        }
        do {
            guard data.length > 0 else { throw CocoaError(.fileWriteUnknown) }
            try FileManager.default.createDirectory(at: workingCopy.deletingLastPathComponent(), withIntermediateDirectories: true)
            try (data as Data).write(to: workingCopy)
            initializationError = nil
        } catch { initializationError = error }
    }

    init(configuration: ReadConfiguration) throws {
        guard let data = configuration.file.regularFileContents else { throw CocoaError(.fileReadCorruptFile) }
        workingCopy = Self.makeWorkingCopyURL(named: configuration.file.filename ?? "Document.pdf")
        initializationError = nil
        try FileManager.default.createDirectory(at: workingCopy.deletingLastPathComponent(), withIntermediateDirectories: true)
        try data.write(to: workingCopy)
    }

    /// Opens a file the app can read directly, without the system's document browser.
    init(contentsOf url: URL) throws {
        workingCopy = Self.makeWorkingCopyURL(named: url.lastPathComponent)
        initializationError = nil
        try FileManager.default.createDirectory(at: workingCopy.deletingLastPathComponent(), withIntermediateDirectories: true)
        try FileManager.default.copyItem(at: url, to: workingCopy)
    }

    deinit {
        publicationObserver.withLock { $0?.stop() }
        // The engine keeps its own file handle, so removing the folder never disturbs a call still finishing.
        try? FileManager.default.removeItem(at: workingCopy.deletingLastPathComponent())
    }

    /// Opens the working copy in the engine once; later calls return the same document.
    func openEngine(password: String? = nil, originalURL: URL? = nil) async throws -> EngineDocument {
        if let initializationError { throw initializationError }
        if let document = opened.withLock({ $0 }) { return document }
        try await EngineRuntime.ready()
        let document = try await EngineDocument.open(workingCopy, password: password, recoveryIdentity: Self.recoveryIdentity(for: originalURL))
        let result = opened.withLock { current in
            if let current { return current }
            current = document
            return document
        }
        observeOriginalURL(originalURL, engine: result)
        return result
    }

    /// Captures the drawing and engine state while SwiftUI has disabled edits. Serialization uses only these bytes.
    func snapshot(contentType: UTType) throws -> Data {
        if let initializationError { throw initializationError }
        for prepare in snapshotPreparations.withLock({ Array($0.values) }) {
            if Thread.isMainThread {
                try MainActor.assumeIsolated { try prepare() }
            } else {
                try DispatchQueue.main.sync { try MainActor.assumeIsolated { try prepare() } }
            }
        }
        // A view may have gone away after a failed commit; its immutable drawing still belongs to this document.
        try commitPendingDrawingsAndWait()
        if let engine = opened.withLock({ $0 }) { return try engine.prepareSaveAndWait() }
        return try Data(contentsOf: workingCopy)
    }

    func fileWrapper(snapshot: Data, configuration: WriteConfiguration) throws -> FileWrapper {
        FileWrapper(regularFileWithContents: snapshot)
    }

    func setSnapshotPreparation(for drawingID: UUID, _ prepare: @escaping @MainActor @Sendable () throws -> Void) {
        snapshotPreparations.withLock { $0[drawingID] = prepare }
    }

    func removeSnapshotPreparation(for drawingID: UUID) {
        snapshotPreparations.withLock { $0.removeValue(forKey: drawingID) }
    }

    var hasPendingDrawing: Bool { pendingDrawings.withLock { !$0.isEmpty } }

    func updatePendingDrawing(_ requests: [AnnotationRequest], for drawingID: UUID) {
        pendingDrawings.withLock {
            if requests.isEmpty { $0.removeValue(forKey: drawingID) }
            else { $0[drawingID] = requests }
        }
    }

    @discardableResult
    func commitPendingDrawingAndWait(for drawingID: UUID) throws -> Bool {
        try pendingDrawings.withLock { drafts in
            guard let requests = drafts[drawingID], !requests.isEmpty else { return false }
            guard let engine = opened.withLock({ $0 }) else { throw CocoaError(.fileWriteUnknown) }
            try engine.addAnnotationsAndWait(requests)
            drafts.removeValue(forKey: drawingID)
            return true
        }
    }

    /// Flushes drafts whose page views have gone away. Live views prepare their own drawing and undo entry first.
    private func commitPendingDrawingsAndWait() throws {
        try pendingDrawings.withLock { drafts in
            let requests = drafts.values.flatMap { $0 }
            guard !requests.isEmpty else { return }
            guard let engine = opened.withLock({ $0 }) else { throw CocoaError(.fileWriteUnknown) }
            try engine.addAnnotationsAndWait(requests)
            drafts.removeAll()
        }
    }

    func observeOriginalURL(_ url: URL?, engine: EngineDocument) {
        publicationObserver.withLock { observer in
            guard observer?.presentedItemURL != url else { return }
            observer?.stop()
            observer = url.map { StudioFilePublicationObserver(url: $0, engine: engine) }
        }
    }

    private static func makeWorkingCopyURL(named name: String) -> URL {
        let folder = workingCopies.appending(path: UUID().uuidString, directoryHint: .isDirectory)
        return folder.appending(path: name)
    }

    /// The original URL identifies recovery without retaining a user's full path in recovery metadata.
    static func recoveryIdentity(for url: URL?) -> String? {
        guard let url else { return nil }
        let digest = SHA256.hash(data: Data(url.standardizedFileURL.absoluteString.utf8))
        return "studio-document://sha256/" + digest.map { String(format: "%02x", $0) }.joined()
    }

    static let workingCopies = StudioTemporaryFiles.root.appending(path: "WorkingCopies", directoryHint: .isDirectory)

    /// Copies of open documents written for sharing; a destination may still read one after the share sheet closes.
    static let sharedCopies = StudioTemporaryFiles.root.appending(path: "Shared", directoryHint: .isDirectory)

    /// Removes working and shared copies left by an earlier process.
    static func removeStaleWorkingCopies() {
        StudioTemporaryFiles.removeStale()
    }
}

/// Starts the engine once per process, before the first document opens.
enum EngineRuntime {
    private static let start = Task<Void, Error> {
        let support = FileManager.default.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
        try await Engine.initialize(dataDirectory: support.appending(path: "Engine", directoryHint: .isDirectory))
    }

    static func ready() async throws {
        try await start.value
    }
}
