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
    var needsRecoveryDecision = false
    var isRunningOperation = false
    private var pendingEdits = 0
    var isWorking: Bool { isRunningOperation || pendingEdits > 0 }
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
    /// How far the pages are turned on screen, in clockwise degrees; the document itself doesn't change.
    var viewRotation = 0
    /// How the pages are laid out for reading; the last choice is kept for new windows.
    var readingLayout = ReadingLayout(rawValue: UserDefaults.standard.string(forKey: "ReadingLayout") ?? "") ?? .continuous {
        didSet {
            UserDefaults.standard.set(readingLayout.rawValue, forKey: "ReadingLayout")
            zoomToFit = 1
        }
    }
    private(set) var canvasRequest: CanvasRequest?
    /// The window's undo manager, which the system's Edit menu, undo gestures and edited state follow.
    @ObservationIgnored weak var undoManager: UndoManager?

    /// Identifies this window's pages in drag and drop, so pages dropped from another document are refused.
    let id = UUID()

    /// Pages written to a new PDF, waiting for the user to choose where to save them.
    var extractedPages: ExtractedPages?
    /// A protected or signed copy, offered for saving once the sheet that wrote it closes.
    var pendingExport: ExtractedPages?
    /// The protection or signing sheet shown.
    var securitySheet: SecuritySheet?
    /// A signature waiting for a tap or click on a page to place it.
    var pendingSignature: Signature?
    /// Whether the sheet for drawing a new signature is shown.
    var isCreatingSignature = false
    /// Whether the panel for choosing PDFs to insert pages from is shown.
    var isChoosingPagesToImport = false
    /// An import waiting for the password of one of its PDFs.
    var pendingImport: PendingImport?
    /// The properties, outline and attachments, for the revision they were read from.
    var structure: DocumentStructure?
    let structureLoad = RevisionRead<DocumentStructure>()
    /// What the sidebar lists in the Read task: page thumbnails or the outline.
    var sidebarMode = SidebarMode.thumbnails
    /// The page production or output sheet shown.
    var structureSheet: StructureSheet?
    /// Whether the panel for choosing a file to attach is shown.
    var isChoosingAttachment = false
    /// Files an output wrote, waiting for the user to choose where to save them.
    var exportedFile: ExportedFile?
    /// An output a sheet wrote, offered for saving once the sheet closes.
    var pendingExportedFile: ExportedFile?
    /// The page being read while Make Searchable runs.
    private(set) var recognitionProgress: RecognitionProgress?
    /// The outcome of a change that turned out to change nothing, to tell the user.
    var notice: String?
    /// The title of the notice's alert, when it reports something other than text recognition.
    var noticeTitle: String?
    /// Form fields of pages near the one in view, for the revision they were read from.
    @ObservationIgnored var formPages: [Int: FormFields] = [:]
    @ObservationIgnored var formLoads: [Int: RevisionRead<FormFields>] = [:]
    /// Bumped when form fields load, so the page view shows where they are.
    var formFieldsVersion = 0

    // Markup
    /// Whether drawing not yet added to the document is on the page.
    var hasPendingMarkup: Bool {
        _ = pendingDrawingVersion
        return document.hasPendingDrawing
    }
    private var pendingDrawingVersion = 0

    /// Invalidates controls observing a drawing change; the document remains the authority for retained drafts.
    func pendingDrawingDidChange() { pendingDrawingVersion &+= 1 }
    /// Asks the page view to add its pending drawing to the document, as Done does.
    private(set) var markupCommitRequest: UUID?
    /// The drawing style on the Mac, where the page view draws without PencilKit.
    var markupStyle = MarkupStyle.pen
    var markupColor = EngineColor(red: 0.85, green: 0.12, blue: 0.1)
    /// The shape a drag on the page draws in Markup, in place of the pen.
    var shapeTool: ShapeTool?
    /// A text box drawn on a page, waiting for its text.
    var pendingTextBox: PendingTextBox?
    /// Areas marked on the pages to redact together; nil when not marking.
    var redactionAreas: [RedactionMark]?
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
    @ObservationIgnored private var pageTextLoads: [Int: RevisionRead<PageText>] = [:]

    // Search
    var searchText = ""
    var isSearchFocused = false
    private(set) var searchResult: SearchResult?
    private(set) var isSearching = false
    /// The share of pages searched while a search runs.
    private(set) var searchProgress: Double?
    /// The hit shown, as an index into the current result's hits.
    private(set) var currentHitIndex: Int?
    @ObservationIgnored private var searchTask: Task<Void, Never>?

    /// Engine edits run one after another, in the order they were asked for.
    var edits: Task<Void, Never>?

    init(document: StudioDocument) {
        self.document = document
    }

    var revision: Int64 { facts?.revision ?? -1 }
    var pageCount: Int { facts?.pages.count ?? 0 }
    var isReady: Bool { if case .ready = state { !needsRecoveryDecision } else { false } }

    // MARK: Opening

    func load(password: String? = nil, originalURL: URL? = nil) async {
        guard engine == nil else { return }
        state = .loading
        do {
            let engine = try await document.openEngine(password: password, originalURL: originalURL)
            self.engine = engine
            facts = try await engine.facts()
            needsRecoveryDecision = facts?.hasRecovery == true
            state = .ready
        } catch let error as EngineError where error.status == .passwordRequired || error.status == .invalidPassword {
            state = .needsPassword(retry: error.status == .invalidPassword)
        } catch {
            state = .failed(error.localizedDescription)
        }
    }

    // MARK: Navigation

    func cancelOpening() {
        needsRecoveryDecision = false
        state = .failed(CocoaError(.userCancelled).localizedDescription)
    }

    func resolveRecovery(restore: Bool) {
        guard let engine else { return }
        isRunningOperation = true
        needsRecoveryDecision = false
        Task {
            defer { isRunningOperation = false }
            do {
                try await engine.resolveRecovery(restore: restore)
                facts = try await engine.facts()
                if restore {
                    registerUndo(String(localized: "Restore Changes"), steps: 1)
                    document.objectWillChange.send()
                }
            } catch { state = .failed(error.localizedDescription) }
        }
    }

    func updateOriginalURL(_ url: URL?) async {
        guard let engine else { return }
        do {
            try await engine.setRecoveryIdentity(StudioDocument.recoveryIdentity(for: url))
            document.observeOriginalURL(url, engine: engine)
        }
        catch { errorMessage = error.localizedDescription }
    }

    func show(page: Int) {
        guard (1...max(pageCount, 1)).contains(page) else { return }
        currentPage = page
        canvasRequest = CanvasRequest(kind: .showPage(page))
    }

    /// Turns the pages on screen a quarter, leaving the document as it is.
    func rotateView(clockwise: Bool) {
        viewRotation = (viewRotation + (clockwise ? 90 : 270)) % 360
        zoomToFit = 1
    }

    func showNextPage() { show(page: currentPage + 1) }
    func showPreviousPage() { show(page: currentPage - 1) }

    func zoom(_ kind: CanvasRequest.Kind) {
        canvasRequest = CanvasRequest(kind: kind)
    }

    // MARK: Editing

    var canEditPages: Bool { isReady && !isWorking && facts?.canMutatePages == true }
    var canRotatePages: Bool { isReady && !isWorking && facts?.canRotatePages == true }
    var canInsertBlankPages: Bool { isReady && !isWorking && facts?.canInsertBlankPages == true }
    var canEditAnnotations: Bool { isReady && !isWorking && facts?.canEditAnnotations == true }
    var canExtractPages: Bool { isReady && !isWorking && facts?.canExtractPages == true }
    var canImportPages: Bool { isReady && !isWorking && facts?.canImportPages == true }
    var canMakeSearchable: Bool { isReady && !isWorking && facts?.canMakeSearchable != false }

    /// The pages a page command applies to: the selection in the Pages task, otherwise the current page.
    var targetPages: [Int] {
        task == .pages && !selection.isEmpty ? selection.sorted() : [currentPage]
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
        let name = pages.count == 1 ? String(localized: "\(base) – Page \(pages[0]).pdf") : String(localized: "\(base) – \(pages.count) Pages.pdf")
        let folder = StudioTemporaryFiles.root.appending(path: "Extracted/\(UUID().uuidString)", directoryHint: .isDirectory)
        let url = folder.appending(path: name)
        isRunningOperation = true
        Task {
            defer { isRunningOperation = false }
            do {
                try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
                try await engine.extractPages(pages, to: url)
                extractedPages = ExtractedPages(url: url, document: try ExtractedPDF(url: url))
            } catch {
                try? FileManager.default.removeItem(at: folder)
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
        let load = pageTextLoads[page] ?? RevisionRead<PageText>()
        pageTextLoads[page] = load
        defer {
            if pageTextLoads[page] === load, !load.isLoading { pageTextLoads[page] = nil }
        }
        do {
            let text = try await load.read(source: engine, revision: { self.revision }) { try await engine.pageText(page: page) }
            guard !Task.isCancelled, isReady, self.engine === engine else { return nil }
            if text.revision == revision {
                pageTexts[page] = text
                if pageTexts.count > 8, let far = pageTexts.keys.max(by: { abs($0 - page) < abs($1 - page) }) { pageTexts[far] = nil }
            }
        } catch {
            if !Task.isCancelled, self.engine === engine { errorMessage = error.localizedDescription }
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
        edit(name, { try await $0.addAnnotations(requests) }, then: then, failed: failed)
    }

    /// Asks the page view to add its pending drawing, as Done and leaving the task do.
    func commitMarkup() {
        markupCommitRequest = UUID()
    }

    /// A system snapshot must finish the drawing before returning; no task on the main actor is needed to publish it.
    func commitPendingDrawingSynchronously(for drawingID: UUID) throws {
        guard let engine else { throw CocoaError(.fileWriteUnknown) }
        if try document.commitPendingDrawingAndWait(for: drawingID) {
            registerUndo(String(localized: "Add Drawing"), steps: 1)
            do { facts = try engine.factsAndWait() }
            catch { errorMessage = error.localizedDescription }
        }
    }

    /// Adds a note at a point of a page, in the displayed page's points.
    func addNote(_ text: String, page: Int, at point: CGPoint) {
        let request = AnnotationRequest(tool: .note, page: page,
                                        bounds: .init(x: point.x - 10, y: point.y - 10, width: 20, height: 20),
                                        text: text, author: Self.author, color: EngineColor(red: 1, green: 0.8, blue: 0.1))
        addMarkup([request], name: String(localized: "Add Note"))
    }

    var canUndoEdit: Bool { !isWorking && (markupUndoManager?.canUndo == true || facts?.canUndo == true) }
    var canRedoEdit: Bool { !isWorking && (markupUndoManager?.canRedo == true || facts?.canRedo == true) }

    /// Undo acts on the pending drawing first, then on the document, as ⌘Z does from the canvas.
    func undoEdit() {
        guard !isWorking else { return }
        if let markup = markupUndoManager, markup.canUndo { markup.undo() } else { undoManager?.undo() }
    }

    func redoEdit() {
        guard !isWorking else { return }
        if let markup = markupUndoManager, markup.canRedo { markup.redo() } else { undoManager?.redo() }
    }

    func loadAnnotations() async {
        guard let engine, isReady, facts?.canEditAnnotations != false else {
            pageAnnotations = nil
            return
        }
        let page = currentPage
        guard pageAnnotations?.page != page || pageAnnotations?.revision != revision else { return }
        do {
            let annotations = try await engine.annotations(page: page)
            if annotations.revision == revision && page == currentPage { pageAnnotations = annotations }
        } catch {
            if !Task.isCancelled { errorMessage = error.localizedDescription }
        }
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
    func edit(_ name: String, steps: Int = 1, _ change: @escaping @Sendable (EngineDocument) async throws -> Void,
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
        pendingEdits += 1
        edits = Task {
            defer { pendingEdits -= 1 }
            await previous?.value
            do {
                try Task.checkCancellation()
                try await edit(engine)
                facts = try await engine.facts()
                currentPage = min(max(currentPage, 1), max(pageCount, 1))
                selection = selection.filter { (1...max(pageCount, 1)).contains($0) }
                if textSelection?.revision != revision { textSelection = nil }
                document.objectWillChange.send()
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
    /// Searches the document, showing hits as their pages are searched: the first hit at or after the page in view
    /// becomes current as soon as it is found. Searching again for the same text keeps the current hit.
    func runSearch() {
        searchTask?.cancel()
        let query = searchText.trimmingCharacters(in: .whitespacesAndNewlines)
        guard let engine, !query.isEmpty else {
            clearSearch()
            return
        }
        isSearching = true
        searchProgress = 0
        let startPage = currentPage
        let isRepeat = searchResult?.query == query
        let keep = isRepeat ? currentHitIndex : nil
        if !isRepeat {
            searchResult = nil
            currentHitIndex = nil
        }
        let (pages, sink) = AsyncStream.makeStream(of: SearchPage.self)
        searchTask = Task {
            let streaming = Task {
                var hits: [SearchResult.Hit] = []
                for await page in pages where !Task.isCancelled {
                    searchProgress = Double(page.pagesSearched) / Double(max(page.pageCount, 1))
                    // A repeated search keeps its hits until it completes, so the current one stays put.
                    guard !isRepeat, !page.hits.isEmpty else { continue }
                    hits += page.hits
                    searchResult = SearchResult(query: page.query, revision: page.revision, hits: hits)
                    if currentHitIndex == nil, let first = hits.firstIndex(where: { $0.page >= startPage }) {
                        currentHitIndex = first
                        revealCurrentHit()
                    }
                }
            }
            let result: SearchResult
            do {
                result = try await engine.search(query) { sink.yield($0) }
                sink.finish()
                await streaming.value
            } catch {
                sink.finish()
                streaming.cancel()
                if !Task.isCancelled {
                    isSearching = false
                    searchProgress = nil
                    errorMessage = error.localizedDescription
                }
                return
            }
            guard !Task.isCancelled else { return }
            isSearching = false
            searchProgress = nil
            searchResult = result
            if let keep, result.hits.indices.contains(keep) {
                currentHitIndex = keep
            } else if !(currentHitIndex.map(result.hits.indices.contains) ?? false) {
                // No hit after the page in view: the search wraps to the first.
                currentHitIndex = result.hits.isEmpty ? nil : 0
                revealCurrentHit()
            }
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
        searchProgress = nil
    }

    /// Scrolls a rectangle of a page, in its displayed points, into view without changing the zoom.
    func reveal(page: Int, rect: CGRect) {
        canvasRequest = CanvasRequest(kind: .reveal(page: page, rect: rect))
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

/// Which page Make Searchable is reading, from 1; 0 before the first.
struct RecognitionProgress: Equatable {
    let page: Int
    let pageCount: Int
}

/// A change that completed without changing the document, with the reason to tell the user.
struct NoChange: Error {
    let message: String
}
