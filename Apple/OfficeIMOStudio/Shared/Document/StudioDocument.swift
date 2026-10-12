import CoreGraphics
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
        try? (data as Data).write(to: workingCopy)
    }

    init(configuration: ReadConfiguration) throws {
        guard let data = configuration.file.regularFileContents else { throw CocoaError(.fileReadCorruptFile) }
        workingCopy = Self.makeWorkingCopyURL(named: configuration.file.filename ?? "Document.pdf")
        try data.write(to: workingCopy)
    }

    /// Opens a file the app can read directly, without the system's document browser.
    init(contentsOf url: URL) throws {
        workingCopy = Self.makeWorkingCopyURL(named: url.lastPathComponent)
        try FileManager.default.copyItem(at: url, to: workingCopy)
    }

    deinit {
        // The engine keeps its own file handle, so removing the folder never disturbs a call still finishing.
        try? FileManager.default.removeItem(at: workingCopy.deletingLastPathComponent())
    }

    /// Opens the working copy in the engine once; later calls return the same document.
    func openEngine(password: String? = nil) async throws -> EngineDocument {
        if let document = opened.withLock({ $0 }) { return document }
        try await EngineRuntime.ready()
        let document = try await EngineDocument.open(workingCopy, password: password)
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
        try snapshot?.saveInPlaceAndWait()
        return try FileWrapper(url: workingCopy, options: .immediate)
    }

    private static func makeWorkingCopyURL(named name: String) -> URL {
        let folder = workingCopies.appending(path: UUID().uuidString, directoryHint: .isDirectory)
        try? FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
        return folder.appending(path: name)
    }

    static let workingCopies = FileManager.default.temporaryDirectory.appending(path: "WorkingCopies", directoryHint: .isDirectory)

    /// Removes working copies left by an earlier process that ended without closing its documents.
    static func removeStaleWorkingCopies() {
        try? FileManager.default.removeItem(at: workingCopies)
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
