import SwiftUI
import PDFKit
import UniformTypeIdentifiers

/// SwiftUI owns coordinated document reads, autosave, and unsaved-window lifecycle.
final class StudioDocument: ReferenceFileDocument {
    static var readableContentTypes: [UTType] { [.pdf] }
    typealias Snapshot = Data
    private let dataLock = NSLock()
    private var storedData: Data
    var data: Data { dataLock.withLock { storedData } }
    private(set) var pdf: PDFDocument

    init(data: Data) throws {
        let inspection = try StudioEngine.inspect(data)
        guard let pdf = DisplayPDFDocument(data: data), inspection.pages > 0,
              pdf.pageCount == inspection.pages, !pdf.isLocked else {
            throw StudioEngine.EngineError(message: "This PDF could not be opened for editing.")
        }
        self.storedData = data
        self.pdf = pdf
    }

    required convenience init(configuration: ReadConfiguration) throws {
        guard let data = configuration.file.regularFileContents else {
            throw StudioEngine.EngineError(message: "Choose a PDF document.")
        }
        try self.init(data: data)
    }

    func snapshot(contentType: UTType) throws -> Data { data }
    func fileWrapper(snapshot: Data, configuration: WriteConfiguration) throws -> FileWrapper {
        FileWrapper(regularFileWithContents: snapshot)
    }

    /// Installs verified engine output as one undoable document edit.
    func replace(with updated: Data, undoManager: UndoManager?) throws {
        _ = try StudioEngine.inspect(updated)
        guard let rendered = DisplayPDFDocument(data: updated) else {
            throw StudioEngine.EngineError(message: "The edited PDF could not be displayed. Your document was kept unchanged.")
        }
        apply(updated, rendered: rendered, undoManager: undoManager)
    }

    private func apply(_ updated: Data, rendered: PDFDocument, undoManager: UndoManager?) {
        let previous = data
        let previousPDF = pdf
        undoManager?.levelsOfUndo = max(1, min(20, (128 * 1024 * 1024) / max(updated.count, previous.count, 1)))
        undoManager?.registerUndo(withTarget: self) { [weak undoManager] target in
            target.apply(previous, rendered: previousPDF, undoManager: undoManager)
        }
        undoManager?.setActionName("Add Note")
        objectWillChange.send()
        pdf = rendered
        dataLock.withLock { storedData = updated }
    }
}

struct PDFExport: FileDocument {
    static var readableContentTypes: [UTType] { [.pdf] }
    var data: Data
    init(data: Data) { self.data = data }
    init(configuration: ReadConfiguration) throws {
        guard let data = configuration.file.regularFileContents else { throw CocoaError(.fileReadCorruptFile) }
        self.data = data
    }
    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper { FileWrapper(regularFileWithContents: data) }
}

/// PDFKit must not offer edits that bypass the OfficeIMO byte snapshot.
private final class DisplayPDFDocument: PDFDocument {
    override var allowsCommenting: Bool { false }
    override var allowsFormFieldEntry: Bool { false }
    override var allowsDocumentChanges: Bool { false }
}
