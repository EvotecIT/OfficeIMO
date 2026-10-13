import CoreTransferable
import Foundation
import UniformTypeIdentifiers

extension DocumentSession {
    /// The document as it is shown, for sharing: the saved file when nothing is unsaved, otherwise a copy of the open
    /// revision under the document's name, so what is shared matches the window.
    func shareableFile(saved fileURL: URL) async throws -> URL {
        guard facts?.isDirty == true, let engine else { return fileURL }
        let folder = StudioDocument.sharedCopies.appending(path: UUID().uuidString, directoryHint: .isDirectory)
        try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
        let copy = folder.appending(path: fileURL.lastPathComponent)
        try await engine.saveCopy(to: copy)
        return copy
    }
}

/// The open document for ShareLink: the file is written when a destination asks for it.
struct SharedDocument: Transferable {
    let session: DocumentSession
    let fileURL: URL

    static var transferRepresentation: some TransferRepresentation {
        FileRepresentation(exportedContentType: .pdf) { shared in
            SentTransferredFile(try await shared.session.shareableFile(saved: shared.fileURL))
        }
        .suggestedFileName { $0.fileURL.lastPathComponent }
    }
}
