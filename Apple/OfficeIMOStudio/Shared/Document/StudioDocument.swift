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
        // The engine keeps its own file handle, so removing the folder never disturbs a call still finishing.
        try? FileManager.default.removeItem(at: workingCopy.deletingLastPathComponent())
    }

    /// Opens the working copy in the engine once; later calls return the same document.
    func openEngine(password: String? = nil, originalURL: URL? = nil) async throws -> EngineDocument {
        if let initializationError { throw initializationError }
        if let document = opened.withLock({ $0 }) { return document }
        try await EngineRuntime.ready()
        let document = try await EngineDocument.open(workingCopy, password: password, recoveryIdentity: Self.recoveryIdentity(for: originalURL))
        return opened.withLock { current in
            if let current { return current }
            current = document
            return document
        }
    }

    /// The engine's document at the moment the system asks to save.
    func snapshot(contentType: UTType) throws -> EngineDocument? {
        opened.withLock { $0 }
    }

    /// Runs on the system's saving thread: the engine saves its working copy, which becomes the saved file.
    func fileWrapper(snapshot: EngineDocument?, configuration: WriteConfiguration) throws -> FileWrapper {
        if let initializationError { throw initializationError }
        try snapshot?.saveInPlaceAndWait()
        return try FileWrapper(url: workingCopy, options: .immediate)
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
