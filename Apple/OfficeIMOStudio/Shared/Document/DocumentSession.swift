import CoreGraphics
import Foundation
import Observation
import SwiftUI
import UniformTypeIdentifiers
#if os(macOS)
import AppKit
#else
import UIKit
#endif

/// The tasks a document window offers. Only tasks the engine supports are listed.
enum StudioTask: String, CaseIterable, Identifiable {
    case read
    case markup
    case pages

    var id: Self { self }

    var title: LocalizedStringKey {
        switch self {
        case .read: "Read"
        case .markup: "Markup"
        case .pages: "Pages"
        }
    }

    var systemImage: String {
        switch self {
        case .read: "doc.text"
        case .markup: "pencil.tip.crop.circle"
        case .pages: "square.grid.2x2"
        }
    }
}

/// A request for the page view, identified so a repeated request is still seen as new.
struct CanvasRequest: Equatable {
    enum Kind: Equatable {
        case showPage(Int)
        case zoomIn
        case zoomOut
        case actualSize
        case fitWidth
        /// Brings a rectangle of a page, in the displayed page's points, into view without changing the zoom.
        case reveal(page: Int, rect: CGRect)
    }

    let kind: Kind
    let id = UUID()
}

/// One window's document state. It lives in the scene, not the views, so changing between compact and regular width
/// (an iPhone Duo folding, a split view resizing) keeps the page, zoom, task and selection.
@MainActor @Observable
final class DocumentSession {
    enum LoadState {
        case loading
        case needsPassword(retry: Bool)
        case ready
        case failed(String)
    }

    let document: StudioDocument
    let thumbnails = ThumbnailStore()

    private(set) var engine: EngineDocument?
    private(set) var facts: DocumentFacts?
    private(set) var state = LoadState.loading
    private(set) var isWorking = false
    var errorMessage: String?

    var task = StudioTask.read
    /// The page at the middle of the page view, numbered from 1.
    var currentPage = 1
    /// Pages selected in the Pages task.
    var selection = Set<Int>()
    /// The page list: the sidebar in regular width, a sheet in compact width.
    var isPageListShown: Bool?
    var isInspectorShown = false
    /// The page view's zoom as a multiple of the fit-to-width zoom, kept when the width changes.
    var zoomToFit = 1.0
    /// Points on screen per PDF point, reported by the page view.
    var zoom = 1.0
    private(set) var canvasRequest: CanvasRequest?
    /// The window's undo manager, which the system's Edit menu, undo gestures and edited state follow.
    @ObservationIgnored weak var undoManager: UndoManager?

    /// Identifies this window's pages in drag and drop, so pages dropped from another document are refused.
    let id = UUID()

    /// Pages written to a new PDF, waiting for the user to choose where to save them.
    var extractedPages: ExtractedPages?
    /// Whether the panel for choosing PDFs to insert pages from is shown.
    var isChoosingPagesToImport = false
    /// An import waiting for the password of one of its PDFs.
    var pendingImport: PendingImport?
    /// The page being read while Make Searchable runs.
    private(set) var recognitionProgress: RecognitionProgress?
    /// The outcome of a change that turned out to change nothing, to tell the user.
    var notice: String?

    // Markup
    /// Whether drawing not yet added to the document is on the page.
    var hasPendingMarkup = false
    /// Asks the page view to add its pending drawing to the document, as Done does.
    private(set) var markupCommitRequest: UUID?
    /// The drawing style on the Mac, where the page view draws without PencilKit.
    var markupStyle = MarkupStyle.pen
    var markupColor = EngineColor(red: 0.85, green: 0.12, blue: 0.1)
    /// The pending drawing's own history, which Undo and Redo act on before the window's.
    @ObservationIgnored weak var markupUndoManager: UndoManager?
    /// The middle of the page view, in the current page's displayed points, where a new note goes.
    @ObservationIgnored var viewCenterOnPage = CGPoint(x: 72, y: 72)
    /// The annotations on the current page, for the inspector and selection.
    private(set) var pageAnnotations: AnnotationList?
    var selectedAnnotation: Int?

    // Text selection
    /// Text selected in the Read task, on one page.
    var textSelection: TextSelection?
    @ObservationIgnored private var pageTexts: [Int: PageText] = [:]
    @ObservationIgnored private var pageTextLoads: [Int: Task<PageText?, Never>] = [:]

    // Search
    var searchText = ""
    var isSearchFocused = false
    private(set) var searchResult: SearchResult?
    private(set) var isSearching = false
    /// The hit shown, as an index into the current result's hits.
    private(set) var currentHitIndex: Int?
    @ObservationIgnored private var searchTask: Task<Void, Never>?

    /// Engine edits run one after another, in the order they were asked for.
    private var edits: Task<Void, Never>?

    init(document: StudioDocument) {
        self.document = document
    }

    var revision: Int64 { facts?.revision ?? -1 }
    var pageCount: Int { facts?.pages.count ?? 0 }
    var isReady: Bool { if case .ready = state { true } else { false } }

    // MARK: Opening

    func load(password: String? = nil) async {
        guard engine == nil else { return }
        state = .loading
        do {
            let engine = try await document.openEngine(password: password)
            self.engine = engine
            facts = try await engine.facts()
            state = .ready
        } catch let error as EngineError where error.status == .passwordRequired || error.status == .invalidPassword {
            state = .needsPassword(retry: error.status == .invalidPassword)
        } catch {
            state = .failed(error.localizedDescription)
        }
    }

    // MARK: Navigation

    func show(page: Int) {
        guard (1...max(pageCount, 1)).contains(page) else { return }
        currentPage = page
        canvasRequest = CanvasRequest(kind: .showPage(page))
    }

    func showNextPage() { show(page: currentPage + 1) }
    func showPreviousPage() { show(page: currentPage - 1) }

    func zoom(_ kind: CanvasRequest.Kind) {
        canvasRequest = CanvasRequest(kind: kind)
    }

    // MARK: Editing

    var canEditPages: Bool { isReady && facts?.canMutatePages != false }
    var canEditAnnotations: Bool { isReady && facts?.canEditAnnotations != false }
    var canExtractPages: Bool { isReady && facts?.canExtractPages != false }
    var canImportPages: Bool { isReady && facts?.canImportPages != false }
    var canMakeSearchable: Bool { isReady && !isWorking && facts?.canMakeSearchable != false }

    /// The pages a page command applies to: the selection in the Pages task, otherwise the current page.
    var targetPages: [Int] {
        task == .pages && !selection.isEmpty ? selection.sorted() : [currentPage]
    }

    func rotate(_ pages: [Int], degrees: Int) {
        guard !pages.isEmpty else { return }
        edit(String(localized: degrees < 0 ? "Rotate Left" : "Rotate Right")) { try await $0.rotatePages(pages, degrees: degrees) }
    }

    /// Deletes pages; a document keeps at least one.
    func delete(_ pages: [Int]) {
        guard !pages.isEmpty, pages.count < pageCount else { return }
        let remaining = pageCount - pages.count
        edit(String(localized: pages.count == 1 ? "Delete Page" : "Delete Pages")) { try await $0.deletePages(pages) } then: { session in
            session.selection = []
            session.currentPage = min(session.currentPage, remaining)
        }
    }

    /// Duplicates pages; each copy follows its page, and the copies become the selection.
    func duplicate(_ pages: [Int]) {
        let pages = pages.sorted()
        guard !pages.isEmpty else { return }
        // Each copy lands after its page, shifted by the copies inserted before it.
        let copies = Set(pages.enumerated().map { index, page in page + index + 1 })
        edit(String(localized: pages.count == 1 ? "Duplicate Page" : "Duplicate Pages")) { try await $0.duplicatePages(pages) } then: { session in
            session.selection = copies
        }
    }

    /// Moves pages, in their current order, before a page, or to the end with the page count + 1.
    func move(_ pages: [Int], before target: Int) {
        let moving = pages.sorted()
        let count = pageCount
        guard !moving.isEmpty, (1...count + 1).contains(target) else { return }
        let set = Set(moving)
        var order = (1...count).filter { !set.contains($0) }
        let position = order.firstIndex { $0 >= target } ?? order.count
        order.insert(contentsOf: moving, at: position)
        guard order != Array(1...count) else { return }
        let moved = Set((position + 1)...(position + moving.count))
        let newOrder = order
        edit(String(localized: moving.count == 1 ? "Move Page" : "Move Pages")) { try await $0.reorderPages(newOrder) } then: { session in
            session.selection = moved
        }
    }

    /// Inserts a blank page the size of the page it follows, after the last target page.
    func insertBlankPage() {
        guard let facts, let after = targetPages.last, facts.pages.indices.contains(after - 1) else { return }
        let size = facts.pages[after - 1].displaySize
        edit(String(localized: "Insert Blank Page")) {
            try await $0.insertBlankPage(before: after + 1, width: size.width, height: size.height)
        } then: { session in
            session.selection = [after + 1]
            session.show(page: after + 1)
        }
    }

    /// Inserts the pages of PDFs the user chose after the target pages, as one undoable change.
    func importPages(from urls: [URL]) {
        guard !urls.isEmpty, let after = targetPages.last else { return }
        importPages(urls.map { ImportSource(url: $0) }, before: after + 1)
    }

    /// Supplies the password a pending import asked for and tries again.
    func continueImport(password: String) {
        guard var pending = pendingImport else { return }
        pendingImport = nil
        pending.sources[pending.needed.source].password = password
        importPages(pending.sources, before: pending.before)
    }

    private func importPages(_ sources: [ImportSource], before page: Int) {
        let pagesBefore = pageCount
        edit(String(localized: "Insert Pages")) {
            _ = try await $0.importPages(sources, before: page)
        } then: { session in
            let inserted = session.pageCount - pagesBefore
            guard inserted > 0 else { return }
            session.selection = Set(page..<(page + inserted))
            session.show(page: page)
        } passwordNeeded: { [weak self] needed in
            self?.pendingImport = PendingImport(sources: sources, before: page, needed: needed)
        }
    }

    /// Reads the pages' text with Vision and adds it to the document, so scanned pages can be searched, selected and
    /// copied; one undoable change. Pages that already have text keep it.
    func makeSearchable() {
        guard canMakeSearchable else { return }
        recognitionProgress = RecognitionProgress(page: 0, pageCount: pageCount)
        let progress: @Sendable (Int, Int) -> Void = { [weak self] page, count in
            Task { @MainActor in self?.recognitionProgress = RecognitionProgress(page: page, pageCount: count) }
        }
        edit(String(localized: "Make Searchable")) { engine in
            let summary = try await engine.makeSearchable(progress: progress)
            if summary.pages == 0 {
                throw NoChange(message: String(localized: "No text was found that the document doesn’t already have."))
            }
        } then: { session in
            session.recognitionProgress = nil
        } failed: { [weak self] in
            self?.recognitionProgress = nil
        }
    }

    /// Stops the change in progress, such as Make Searchable; the document stays as it was.
    func cancelEdit() {
        edits?.cancel()
    }

    /// Writes pages to a new PDF, then asks where to save it.
    func extract(_ pages: [Int]) {
        guard let engine, !pages.isEmpty else { return }
        let base = (facts?.fileName as NSString?)?.deletingPathExtension ?? "Document"
        let name = pages.count == 1 ? "\(base) – Page \(pages[0]).pdf" : "\(base) – \(pages.count) Pages.pdf"
        let folder = FileManager.default.temporaryDirectory.appending(path: "Extracted/\(UUID().uuidString)", directoryHint: .isDirectory)
        let url = folder.appending(path: name)
        isWorking = true
        Task {
            defer { isWorking = false }
            do {
                try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
                try await engine.extractPages(pages, to: url)
                extractedPages = ExtractedPages(url: url, document: try ExtractedPDF(url: url))
            } catch {
                errorMessage = error.localizedDescription
            }
        }
    }

    // MARK: Text selection

    var canCopyText: Bool { isReady && facts?.canExtractText != false }

    /// A page's text, if loaded for the revision shown; loadPageText fetches it.
    func pageText(_ page: Int) -> PageText? {
        guard let text = pageTexts[page], text.revision == revision else { return nil }
        return text
    }

    /// Loads a page's text once per revision, keeping a few pages.
    @discardableResult
    func loadPageText(_ page: Int) async -> PageText? {
        if let text = pageText(page) { return text }
        guard let engine, canCopyText, (1...max(pageCount, 1)).contains(page) else { return nil }
        if let load = pageTextLoads[page] { return await load.value }
        let load = Task<PageText?, Never> { try? await engine.pageText(page: page) }
        pageTextLoads[page] = load
        let text = await load.value
        pageTextLoads[page] = nil
        if let text, text.revision == revision {
            pageTexts[page] = text
            if pageTexts.count > 8, let far = pageTexts.keys.max(by: { abs($0 - page) < abs($1 - page) }) { pageTexts[far] = nil }
        }
        return pageText(page)
    }

    var selectedText: String? {
        guard let selection = textSelection, let text = pageText(selection.page) else { return nil }
        let string = text.string(selection.range)
        return string.isEmpty ? nil : string
    }

    /// The selection's lines, in its page's displayed points.
    var selectionRects: [CGRect] {
        guard let selection = textSelection, let text = pageText(selection.page) else { return [] }
        return text.lineRects(selection.range)
    }

    func copySelection() {
        guard let text = selectedText else { return }
        #if os(macOS)
        NSPasteboard.general.clearContents()
        NSPasteboard.general.setString(text, forType: .string)
        #else
        UIPasteboard.general.string = text
        #endif
    }

    /// Highlights, underlines or strikes through the selected text, as one undoable change.
    func markUpSelection(_ tool: AnnotationRequest.Tool) {
        guard let selection = textSelection, canEditAnnotations else { return }
        let rects = selectionRects
        guard let first = rects.first else { return }
        let bounds = rects.dropFirst().reduce(first) { $0.union($1) }
        let color = tool == .highlight ? EngineColor(red: 1, green: 0.85, blue: 0.1) : EngineColor(red: 0.85, green: 0.12, blue: 0.1)
        let request = AnnotationRequest(tool: tool, page: selection.page,
                                        bounds: .init(x: bounds.minX, y: bounds.minY, width: bounds.width, height: bounds.height),
                                        author: Self.author, color: color, opacity: tool == .highlight ? 0.5 : nil,
                                        quads: rects.map { .init(x: $0.minX, y: $0.minY, width: $0.width, height: $0.height) })
        let name: String = switch tool {
        case .highlight: String(localized: "Highlight")
        case .underline: String(localized: "Underline")
        default: String(localized: "Strikethrough")
        }
        textSelection = nil
        addMarkup([request], name: name)
    }

    // MARK: Markup

    /// Adds drawn markup as one undoable change, however many annotations it takes.
    func addMarkup(_ requests: [AnnotationRequest], name: String, then: ((DocumentSession) -> Void)? = nil,
                   failed: (() -> Void)? = nil) {
        guard !requests.isEmpty else { return }
        edit(name, steps: requests.count, { engine in
            for (index, request) in requests.enumerated() {
                do {
                    try await engine.addAnnotation(request)
                } catch {
                    // Additions that went through are taken back, so the drawing stays one change or none.
                    for _ in 0..<index { try? await engine.undo() }
                    throw error
                }
            }
        }, then: then, failed: failed)
    }

    /// Asks the page view to add its pending drawing, as Done and leaving the task do.
    func commitMarkup() {
        markupCommitRequest = UUID()
    }

    /// Adds a note at a point of a page, in the displayed page's points.
    func addNote(_ text: String, page: Int, at point: CGPoint) {
        let request = AnnotationRequest(tool: .note, page: page,
                                        bounds: .init(x: point.x - 10, y: point.y - 10, width: 20, height: 20),
                                        text: text, author: Self.author, color: EngineColor(red: 1, green: 0.8, blue: 0.1))
        addMarkup([request], name: String(localized: "Add Note"))
    }

    var canUndoEdit: Bool { (hasPendingMarkup && markupUndoManager?.canUndo == true) || facts?.canUndo == true }
    var canRedoEdit: Bool { (hasPendingMarkup && markupUndoManager?.canRedo == true) || facts?.canRedo == true }

    /// Undo acts on the pending drawing first, then on the document, as ⌘Z does from the canvas.
    func undoEdit() {
        if hasPendingMarkup, let markup = markupUndoManager, markup.canUndo { markup.undo() } else { undoManager?.undo() }
    }

    func redoEdit() {
        if hasPendingMarkup, let markup = markupUndoManager, markup.canRedo { markup.redo() } else { undoManager?.redo() }
    }

    func loadAnnotations() async {
        guard let engine, isReady, facts?.canEditAnnotations != false else {
            pageAnnotations = nil
            return
        }
        let page = currentPage
        guard pageAnnotations?.page != page || pageAnnotations?.revision != revision else { return }
        pageAnnotations = try? await engine.annotations(page: page)
    }

    func removeAnnotation(_ annotation: AnnotationList.Annotation) {
        if selectedAnnotation == annotation.objectNumber { selectedAnnotation = nil }
        edit(String(localized: "Delete Annotation")) { try await $0.removeAnnotation(annotation.objectNumber) }
    }

    func updateNote(_ annotation: AnnotationList.Annotation, text: String) {
        edit(String(localized: "Edit Note")) {
            try await $0.updateAnnotation(AnnotationUpdate(objectNumber: annotation.objectNumber, contents: text))
        }
    }

    /// The name markup carries as its author: the user's full name where the platform gives it.
    static var author: String {
        #if os(macOS)
        NSFullUserName()
        #else
        ""
        #endif
    }

    /// Runs an engine edit after any earlier one and registers its undo once it has succeeded. An edit that takes
    /// several engine changes undoes them together.
    private func edit(_ name: String, steps: Int = 1, _ change: @escaping @Sendable (EngineDocument) async throws -> Void,
                      then: ((DocumentSession) -> Void)? = nil, failed: (() -> Void)? = nil,
                      passwordNeeded: ((ImportPasswordNeeded) -> Void)? = nil) {
        enqueue(change, failed: failed, passwordNeeded: passwordNeeded) { session in
            session.registerUndo(name, steps: steps)
            then?(session)
        }
    }

    // While the system undoes or redoes, the inverse must be registered at once, so it goes on the opposite stack; the
    // engine call follows. A failed engine undo or redo would leave the two histories apart, so it clears the window's.
    private func registerUndo(_ name: String, steps: Int) {
        guard let undoManager else { return }
        undoManager.registerUndo(withTarget: self) { session in
            session.registerRedo(name, steps: steps)
            session.enqueue({ engine in for _ in 0..<steps { try await engine.undo() } }, failureClearsHistory: true)
        }
        undoManager.setActionName(name)
    }

    private func registerRedo(_ name: String, steps: Int) {
        guard let undoManager else { return }
        undoManager.registerUndo(withTarget: self) { session in
            session.registerUndo(name, steps: steps)
            session.enqueue({ engine in for _ in 0..<steps { try await engine.redo() } }, failureClearsHistory: true)
        }
        undoManager.setActionName(name)
    }

    private func enqueue(_ edit: @escaping @Sendable (EngineDocument) async throws -> Void, failureClearsHistory: Bool = false,
                         failed: (() -> Void)? = nil, passwordNeeded: ((ImportPasswordNeeded) -> Void)? = nil,
                         succeeded: ((DocumentSession) -> Void)? = nil) {
        guard let engine else { return }
        let previous = edits
        isWorking = true
        edits = Task {
            await previous?.value
            do {
                try await edit(engine)
                facts = try await engine.facts()
                if textSelection?.revision != revision { textSelection = nil }
                succeeded?(self)
                if searchResult != nil { runSearch() }
            } catch let needed as ImportPasswordNeeded where passwordNeeded != nil {
                // Nothing changed: the app asks for the password and tries again.
                passwordNeeded?(needed)
            } catch let unchanged as NoChange {
                failed?()
                notice = unchanged.message
            } catch where Task.isCancelled || (error as? EngineError)?.status == .cancelled {
                // Cancelled by the user: the document is as it was.
                failed?()
            } catch {
                if failureClearsHistory { undoManager?.removeAllActions(withTarget: self) }
                facts = try? await engine.facts()
                failed?()
                errorMessage = error.localizedDescription
            }
            isWorking = false
        }
    }

    // MARK: Search

    /// The current result's hits, if it belongs to the document as shown.
    var searchHits: [SearchResult.Hit] {
        guard let searchResult, searchResult.revision == revision else { return [] }
        return searchResult.hits
    }

    var currentHit: SearchResult.Hit? {
        guard let currentHitIndex, searchHits.indices.contains(currentHitIndex) else { return nil }
        return searchHits[currentHitIndex]
    }

    /// Finds the search text and shows the first hit at or after the current page.
    func runSearch() {
        searchTask?.cancel()
        let query = searchText.trimmingCharacters(in: .whitespacesAndNewlines)
        guard let engine, !query.isEmpty else {
            clearSearch()
            return
        }
        isSearching = true
        let startPage = currentPage
        searchTask = Task {
            let result: SearchResult
            do {
                result = try await engine.search(query)
            } catch {
                if !Task.isCancelled {
                    isSearching = false
                    errorMessage = error.localizedDescription
                }
                return
            }
            guard !Task.isCancelled else { return }
            isSearching = false
            let keep = searchResult?.query == result.query ? currentHitIndex : nil
            searchResult = result
            currentHitIndex = keep.flatMap { result.hits.indices.contains($0) ? $0 : nil }
                ?? result.hits.firstIndex { $0.page >= startPage } ?? (result.hits.isEmpty ? nil : 0)
            if keep == nil { revealCurrentHit() }
        }
    }

    /// Shows the next hit, or the previous one with a negative step, wrapping around the document.
    func showHit(step: Int) {
        let count = searchHits.count
        guard count > 0 else { return }
        currentHitIndex = ((currentHitIndex ?? -step) + step + count) % count
        revealCurrentHit()
    }

    func clearSearch() {
        searchTask?.cancel()
        searchResult = nil
        currentHitIndex = nil
        isSearching = false
    }

    private func revealCurrentHit() {
        guard let hit = currentHit else { return }
        currentPage = hit.page
        canvasRequest = CanvasRequest(kind: .reveal(page: hit.page, rect: hit.bounds.cgRect))
    }
}

/// The Mac's drawing styles, as Preview's Sketch and Draw tools offer them.
enum MarkupStyle: String, CaseIterable, Identifiable {
    case pen
    case marker

    var id: Self { self }
    var title: LocalizedStringKey { self == .pen ? "Pen" : "Highlighter" }
    var systemImage: String { self == .pen ? "pencil.tip" : "highlighter" }
    /// Width in points and opacity of the stroke.
    var strokeWidth: Double { self == .pen ? 2 : 12 }
    var opacity: Double { self == .pen ? 1 : 0.4 }
}

/// A PDF of extracted pages in the app's temporary folder, saved where the user chooses.
struct ExtractedPages: Identifiable {
    let url: URL
    let document: ExtractedPDF
    var id: URL { url }
}

/// The extracted pages as a document for the system's export panel.
struct ExtractedPDF: FileDocument {
    static let readableContentTypes: [UTType] = [.pdf]
    let data: Data

    init(url: URL) throws {
        data = try Data(contentsOf: url)
    }

    init(configuration: ReadConfiguration) throws {
        data = configuration.file.regularFileContents ?? Data()
    }

    func fileWrapper(configuration: WriteConfiguration) throws -> FileWrapper {
        FileWrapper(regularFileWithContents: data)
    }
}

/// Page thumbnails as PNG from the engine, which keeps them apart from the page view's tiles.
@MainActor
final class ThumbnailStore {
    private struct Key: Hashable {
        let page: Int
        let revision: Int64
        let pixelWidth: Int
    }

    private var images: [Key: CGImage] = [:]
    private var order: [Key] = []
    private let limit = 160

    func image(page: Int, pixelWidth: Int, session: DocumentSession) async -> CGImage? {
        guard let engine = session.engine, let facts = session.facts, facts.pages.indices.contains(page - 1) else { return nil }
        let key = Key(page: page, revision: facts.revision, pixelWidth: pixelWidth)
        if let image = images[key] { return image }
        let scale = Double(pixelWidth) / facts.pages[page - 1].displaySize.width
        guard let rendered = try? await engine.renderPage(page, scale: scale),
              let image = try? await Task.detached(operation: { try PageImage.decode(rendered.png) }).value else { return nil }
        images[key] = image
        order.append(key)
        if order.count > limit { images[order.removeFirst()] = nil }
        return image
    }
}

extension FocusedValues {
    /// The document session of the focused window, for menu commands.
    @Entry var documentSession: DocumentSession?
}

/// An import waiting for a password: what to import, where, and which PDF needs it.
struct PendingImport: Identifiable {
    let id = UUID()
    var sources: [ImportSource]
    let before: Int
    let needed: ImportPasswordNeeded
}

/// Which page Make Searchable is reading, from 1; 0 before the first.
struct RecognitionProgress: Equatable {
    let page: Int
    let pageCount: Int
}

/// A change that completed without changing the document, with the reason to tell the user.
struct NoChange: Error {
    let message: String
}
