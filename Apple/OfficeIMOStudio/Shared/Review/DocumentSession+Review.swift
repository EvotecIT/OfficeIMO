import Foundation
import SwiftUI

extension DocumentSession {
    var canRedact: Bool { isReady && !isWorking && currentStructure?.canRedact == true }
    var canCheck: Bool { isReady }

    /// Checks the document's signatures, a standard or its health; the document is unchanged.
    func check(_ request: CheckRequest) async throws -> CheckResult {
        guard let engine else { throw CancellationError() }
        return try await engine.checkDocument(request)
    }

    /// Writes an optimized, repaired or sanitized copy, which the window offers to save once the sheet closes, and
    /// returns what changed.
    func writeHealthCopy(_ kind: HealthCopyKind, profile: String? = nil) async throws -> String? {
        let name = String(localized: "\(documentBaseName) – \(kind.suffix).pdf")
        var summary: String?
        pendingExportedFile = try await writeOutputFile(named: name) { engine, url in
            let copy = WriteRequest.HealthCopy(path: url.path, profile: profile)
            let request = switch kind {
            case .optimize: WriteRequest(optimize: copy)
            case .repair: WriteRequest(repair: copy)
            case .sanitize: WriteRequest(sanitize: copy)
            }
            return try await engine.writeResult(request).summary
        } summary: { summary = $0 }
        return summary
    }

    /// Finds the areas that would redact text, without changing the document.
    func findRedactions(_ find: RedactionFind) async throws -> [RedactionMark] {
        guard let engine else { return [] }
        return try await engine.redact(RedactionRequest(find: find)).marks ?? []
    }

    /// Redacts as one undoable change and reports how the removal was verified.
    func redact(_ apply: RedactionRequest.Apply, name: String = String(localized: "Redact")) {
        guard canRedact else { return }
        let box = VerificationBox()
        edit(name, { engine in
            box.value = try await engine.redact(RedactionRequest(apply: apply)).verification
        }, then: { session in
            session.textSelection = nil
            guard let verification = box.value else { return }
            session.noticeTitle = verification.verified ? String(localized: "Redacted") : String(localized: "Redaction Not Verified")
            session.notice = verification.verified
                ? String(localized: "The content was removed from \(verification.pages.count) page(s) and checked. Save to keep the change; until then Undo restores it.")
                : String(localized: "Some of the marked content is still in the document or couldn’t be checked. Undo the redaction and mark a larger area.")
        })
    }

    /// Redacts the selected text.
    func redactSelection() {
        guard let selection = textSelection else { return }
        let marks = selectionRects.map { RedactionMark(page: selection.page, bounds: .init(x: $0.minX, y: $0.minY, width: $0.width, height: $0.height)) }
        guard !marks.isEmpty else { return }
        redact(.init(areas: marks))
    }

    /// Makes the selected text a link to a web or mail address or to a page, one link per line of the selection.
    func linkSelection(uri: String?, page: Int?) {
        guard let selection = textSelection, canEditAnnotations else { return }
        let requests = selectionRects.map {
            AnnotationRequest(tool: .link, page: selection.page, bounds: .init(x: $0.minX, y: $0.minY, width: $0.width, height: $0.height),
                              linkUri: uri, linkPage: page)
        }
        guard !requests.isEmpty else { return }
        edit(String(localized: "Add Link"), steps: requests.count, { engine in
            for request in requests { try await engine.addAnnotation(request) }
        }, then: { session in
            session.textSelection = nil
        })
    }

    /// Puts a standard stamp in the middle of the page in view.
    func addStamp(_ stamp: StandardStamp) {
        guard canEditAnnotations, let facts, facts.pages.indices.contains(currentPage - 1) else { return }
        let size = facts.pages[currentPage - 1].displaySize
        let width = min(180, size.width * 0.6), height = width / 3
        let center = viewCenterOnPage
        let x = min(max(center.x - width / 2, 0), size.width - width), y = min(max(center.y - height / 2, 0), size.height - height)
        let request = AnnotationRequest(tool: .stamp, page: currentPage, bounds: .init(x: x, y: y, width: width, height: height),
                                        author: Self.author, color: stamp.color, stamp: stamp.rawValue)
        edit(String(localized: "Add Stamp")) { try await $0.addAnnotation(request) }
    }

    // MARK: Bookmark order

    /// Moves a bookmark among its siblings, or into or out of the one before it.
    func moveBookmark(_ node: OutlineNode, _ move: BookmarkMove) {
        guard canEditOutline, let id = node.bookmark, let place = outlinePlace(of: node.id) else { return }
        let edit: OutlineEdit
        switch move {
        case .up:
            guard place.index > 0 else { return }
            edit = OutlineEdit(op: .move, id: id, parentId: place.parent?.bookmark, index: place.index - 1)
        case .down:
            guard place.index + 1 < place.siblings.count else { return }
            edit = OutlineEdit(op: .move, id: id, parentId: place.parent?.bookmark, index: place.index + 1)
        case .indent:
            guard place.index > 0, let parent = place.siblings[place.index - 1].bookmark else { return }
            edit = OutlineEdit(op: .move, id: id, parentId: parent, index: -1)
        case .outdent:
            guard let parent = place.parent, let grand = outlinePlace(of: parent.id) else { return }
            edit = OutlineEdit(op: .move, id: id, parentId: grand.parent?.bookmark, index: grand.index + 1)
        }
        editStructure(String(localized: "Move Bookmark"), EditRequest(outline: [edit]))
    }

    func canMoveBookmark(_ node: OutlineNode, _ move: BookmarkMove) -> Bool {
        guard canEditOutline, node.bookmark != nil, let place = outlinePlace(of: node.id) else { return false }
        return switch move {
        case .up: place.index > 0
        case .down: place.index + 1 < place.siblings.count
        case .indent: place.index > 0 && place.siblings[place.index - 1].bookmark != nil
        case .outdent: place.parent?.bookmark != nil
        }
    }

    private func outlinePlace(of key: String) -> (parent: OutlineNode?, siblings: [OutlineNode], index: Int)? {
        // The engine's edit session numbers only bookmarks with a page, so positions count those alone.
        func search(_ nodes: [OutlineNode], parent: OutlineNode?) -> (OutlineNode?, [OutlineNode], Int)? {
            let resolved = nodes.filter { $0.bookmark != nil }
            if let index = resolved.firstIndex(where: { $0.id == key }) { return (parent, resolved, index) }
            for node in nodes { if let found = search(node.children ?? [], parent: node) { return found } }
            return nil
        }
        return search(currentStructure?.outlineNodes ?? [], parent: nil)
    }

    /// Writes a file into the app's temporary folder and reads it for the system's export panel.
    private func writeOutputFile(named name: String, write: @escaping @Sendable (EngineDocument, URL) async throws -> String?,
                                 summary: (String?) -> Void) async throws -> ExportedFile? {
        guard let engine else { return nil }
        let folder = StudioTemporaryFiles.root.appending(path: "Exported/\(UUID().uuidString)", directoryHint: .isDirectory)
        let url = folder.appending(path: name)
        isRunningOperation = true
        defer { isRunningOperation = false }
        do {
            try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
            summary(try await write(engine, url))
            return ExportedFile(document: try ExportedDocument(url: url), contentType: .pdf, name: name, folder: folder)
        } catch {
            try? FileManager.default.removeItem(at: folder)
            throw error
        }
    }
}

enum HealthCopyKind: String, Identifiable {
    case optimize, repair, sanitize
    var id: Self { self }
    var suffix: String {
        switch self {
        case .optimize: String(localized: "Optimized")
        case .repair: String(localized: "Repaired")
        case .sanitize: String(localized: "Sanitized")
        }
    }
}

enum BookmarkMove { case up, down, indent, outdent }

/// Carries an engine call's result out of its edit to the main actor that reports it.
private final class VerificationBox: @unchecked Sendable {
    var value: RedactionResult.Verification?
}
