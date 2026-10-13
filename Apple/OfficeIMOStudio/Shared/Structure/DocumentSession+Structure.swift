import Foundation
import SwiftUI
import UniformTypeIdentifiers

/// What the sidebar lists while reading, as Preview's View menu offers.
enum SidebarMode: String, CaseIterable, Identifiable {
    case thumbnails, outline

    var id: Self { self }
    var title: LocalizedStringKey { self == .thumbnails ? "Thumbnails" : "Table of Contents" }
    var systemImage: String { self == .thumbnails ? "square.grid.2x2" : "list.bullet.indent" }
}

/// The page production and output sheets a window shows.
enum StructureSheet: String, Identifiable {
    case crop, pageSize, headerFooter, watermark, bates, split, export, flatten, signatures, standard, health, redact, link

    var id: String { rawValue }
}

extension DocumentSession {
    /// The structure, if read for the revision shown.
    var currentStructure: DocumentStructure? {
        guard let structure, structure.revision == revision else { return nil }
        return structure
    }

    var canEditOutline: Bool { isReady && currentStructure?.canEditOutline == true }
    var canEditProperties: Bool { isReady && currentStructure?.canEditProperties == true }
    var canEditAttachments: Bool { isReady && currentStructure?.canEditAttachments == true }
    var canCrop: Bool { isReady && currentStructure?.canCrop == true }
    var canResize: Bool { isReady && currentStructure?.canResize == true }
    var canStampPages: Bool { isReady && currentStructure?.canStampPages == true }
    var canFlattenForms: Bool { isReady && currentStructure?.canFlattenForms == true }
    var canSplit: Bool { canExtractPages && pageCount > 1 }
    var canExportText: Bool { isReady && facts?.canExtractText != false }

    /// Reads the structure once per revision.
    func loadStructure() async {
        guard let engine, isReady, currentStructure == nil else { return }
        if let structureLoad {
            await structureLoad.value
            return
        }
        let load = Task {
            if let read = try? await engine.structure() { structure = read }
        }
        structureLoad = load
        await load.value
        structureLoad = nil
    }

    /// Applies an edit as one undoable change, then reads the structure it changed.
    func editStructure(_ name: String, _ request: EditRequest, then: ((DocumentSession) -> Void)? = nil) {
        edit(name, { try await $0.edit(request) }, then: { session in
            Task { await session.loadStructure() }
            then?(session)
        })
    }

    // MARK: Properties

    func setProperties(_ properties: DocumentStructure.Properties) {
        guard canEditProperties, properties != currentStructure?.properties else { return }
        editStructure(String(localized: "Change Properties"), EditRequest(properties: properties))
    }

    // MARK: Outline

    /// Adds a bookmark to the page in view, after the existing ones, as Preview's Add Bookmark does.
    func addBookmark(title: String? = nil) {
        guard canEditOutline else { return }
        let name = title ?? String(localized: "Page \(currentPage)")
        editStructure(String(localized: "Add Bookmark"), EditRequest(outline: [OutlineEdit(op: .add, title: name, page: currentPage)])) { session in
            session.sidebarMode = .outline
        }
    }

    func renameBookmark(_ node: OutlineNode, to title: String) {
        let title = title.trimmingCharacters(in: .whitespacesAndNewlines)
        guard canEditOutline, let id = node.bookmark, !title.isEmpty, title != node.title else { return }
        editStructure(String(localized: "Rename Bookmark"), EditRequest(outline: [OutlineEdit(op: .rename, id: id, title: title)]))
    }

    func removeBookmark(_ node: OutlineNode) {
        guard canEditOutline, let id = node.bookmark else { return }
        editStructure(String(localized: "Delete Bookmark"), EditRequest(outline: [OutlineEdit(op: .remove, id: id)]))
    }

    /// Points a bookmark at the page in view.
    func retargetBookmark(_ node: OutlineNode) {
        guard canEditOutline, let id = node.bookmark, node.page != currentPage else { return }
        editStructure(String(localized: "Change Bookmark Page"), EditRequest(outline: [OutlineEdit(op: .retarget, id: id, page: currentPage)]))
    }

    /// Makes the outline from the document's headings.
    func rebuildOutline() {
        guard canEditOutline else { return }
        editStructure(String(localized: "Make Bookmarks from Headings"), EditRequest(outline: [OutlineEdit(op: .rebuild)]))
    }

    // MARK: Attachments

    /// Attaches a file the user chose. The engine reads a copy, so the file's security scope is needed only here.
    func attach(_ url: URL) {
        guard canEditAttachments else { return }
        let folder = FileManager.default.temporaryDirectory.appending(path: "Attaching/\(UUID().uuidString)", directoryHint: .isDirectory)
        let copy = folder.appending(path: url.lastPathComponent)
        do {
            let accessing = url.startAccessingSecurityScopedResource()
            defer { if accessing { url.stopAccessingSecurityScopedResource() } }
            try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
            try FileManager.default.copyItem(at: url, to: copy)
        } catch {
            errorMessage = error.localizedDescription
            return
        }
        edit(String(localized: "Add Attachment"), { try await $0.edit(EditRequest(addAttachment: AttachmentAdd(path: copy.path))) }, then: { session in
            try? FileManager.default.removeItem(at: folder)
            Task { await session.loadStructure() }
        }, failed: {
            try? FileManager.default.removeItem(at: folder)
        })
    }

    func removeAttachment(_ attachment: DocumentStructure.Attachment) {
        guard canEditAttachments else { return }
        editStructure(String(localized: "Delete Attachment"), EditRequest(removeAttachment: attachment.fileName))
    }

    /// Writes an attachment's contents, then asks where to save them.
    func saveAttachment(_ attachment: DocumentStructure.Attachment) {
        Task {
            do {
                exportedFile = try await writeOutput(named: attachment.fileName, contentType: .data) { engine, url in
                    _ = try await engine.write(WriteRequest(attachment: .init(fileName: attachment.fileName, path: url.path)))
                }
            } catch {
                errorMessage = error.localizedDescription
            }
        }
    }

    // MARK: Pages

    func crop(_ pages: [Int], left: Double, top: Double, right: Double, bottom: Double) {
        guard canCrop, !pages.isEmpty else { return }
        editStructure(pages.count == 1 ? String(localized: "Crop Page") : String(localized: "Crop Pages"),
                      EditRequest(crop: CropEdit(pages: pages, left: left, top: top, right: right, bottom: bottom)))
    }

    func resize(_ pages: [Int], width: Double, height: Double, mode: PageSizeEdit.Mode) {
        guard canResize, !pages.isEmpty else { return }
        editStructure(String(localized: "Change Page Size"), EditRequest(pageSize: PageSizeEdit(pages: pages, width: width, height: height, mode: mode)))
    }

    func addHeaderFooter(_ edit: HeaderFooterEdit) {
        guard canStampPages else { return }
        editStructure(String(localized: "Add Header and Footer"), EditRequest(headerFooter: edit))
    }

    func addWatermark(_ edit: WatermarkEdit) {
        guard canStampPages else { return }
        editStructure(String(localized: "Add Watermark"), EditRequest(watermark: edit))
    }

    func addBatesNumbers(_ edit: BatesEdit) {
        guard canStampPages else { return }
        editStructure(String(localized: "Add Bates Numbers"), EditRequest(bates: edit))
    }

    /// Draws every form field into its page, so the form can no longer be filled.
    func flattenForm() {
        guard canFlattenForms else { return }
        editStructure(String(localized: "Flatten Form"), EditRequest(flatten: FlattenEdit()))
    }

    // MARK: Outputs

    /// Splits the document into a folder of PDFs the window then offers to save.
    func split(pagesPerPart: Int? = nil, parts: [SplitPart]? = nil) async throws {
        pendingExportedFile = try await writeOutput(named: String(localized: "\(documentBaseName) – Parts"), contentType: .folder) { engine, url in
            _ = try await engine.write(WriteRequest(split: .init(directory: url.path, pagesPerPart: pagesPerPart, parts: parts)))
        }
    }

    /// Writes the document's text as plain text, Markdown or JSON, which the window then offers to save.
    func exportText(_ format: TextExportFormat) async throws {
        pendingExportedFile = try await writeOutput(named: "\(documentBaseName).\(format.fileExtension)", contentType: format.contentType) { engine, url in
            _ = try await engine.write(WriteRequest(export: .init(format: format.rawValue, path: url.path)))
        }
    }

    /// Offers the output a closed sheet wrote for saving.
    func presentPendingExportedFile() {
        guard let pendingExportedFile else { return }
        self.pendingExportedFile = nil
        exportedFile = pendingExportedFile
    }

    /// Writes an output into the app's temporary folder and reads it for the system's export panel.
    private func writeOutput(named name: String, contentType: UTType,
                             write: @escaping @Sendable (EngineDocument, URL) async throws -> Void) async throws -> ExportedFile? {
        guard let engine else { return nil }
        let folder = FileManager.default.temporaryDirectory.appending(path: "Exported/\(UUID().uuidString)", directoryHint: .isDirectory)
        let url = folder.appending(path: name, directoryHint: contentType == .folder ? .isDirectory : .notDirectory)
        isWorking = true
        defer { isWorking = false }
        do {
            try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
            try await write(engine, url)
            return ExportedFile(document: try ExportedDocument(url: url), contentType: contentType, name: name, folder: folder)
        } catch {
            try? FileManager.default.removeItem(at: folder)
            throw error
        }
    }
}

enum TextExportFormat: String, CaseIterable, Identifiable {
    case text, markdown, json

    var id: Self { self }
    var title: LocalizedStringKey {
        switch self {
        case .text: "Plain Text"
        case .markdown: "Markdown"
        case .json: "JSON"
        }
    }
    var fileExtension: String {
        switch self {
        case .text: "txt"
        case .markdown: "md"
        case .json: "json"
        }
    }
    var contentType: UTType {
        switch self {
        case .text: .plainText
        case .markdown: .markdown
        case .json: .json
        }
    }
}

extension UTType {
    static let markdown = UTType(importedAs: "net.daringfireball.markdown", conformingTo: .plainText)
}

/// An output in the app's temporary folder, saved where the user chooses; the folder goes once it is saved or not.
struct ExportedFile: Identifiable {
    let document: ExportedDocument
    let contentType: UTType
    let name: String
    let folder: URL
    var id: URL { folder }
}

/// A written file or folder of files as a document for the system's export panel.
struct ExportedDocument: FileDocument {
    static let readableContentTypes: [UTType] = [.pdf, .plainText, .markdown, .json, .folder, .data]
    let wrapper: FileWrapper

    init(url: URL) throws {
        wrapper = try FileWrapper(url: url, options: .immediate)
    }

    init(configuration: ReadConfiguration) throws {
        wrapper = configuration.file
    }

    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper { wrapper }
}
