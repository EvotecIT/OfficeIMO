import CoreGraphics
import Foundation
import Observation
import SwiftUI

/// The tasks a document window offers. Only tasks the engine supports are listed.
enum StudioTask: String, CaseIterable, Identifiable {
    case read
    case pages

    var id: Self { self }

    var title: LocalizedStringKey {
        switch self {
        case .read: "Read"
        case .pages: "Pages"
        }
    }

    var systemImage: String {
        switch self {
        case .read: "doc.text"
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

    /// The pages a page command applies to: the selection in the Pages task, otherwise the current page.
    var targetPages: [Int] {
        task == .pages && !selection.isEmpty ? selection.sorted() : [currentPage]
    }

    func rotate(_ pages: [Int], degrees: Int) {
        guard !pages.isEmpty else { return }
        let name = String(localized: degrees < 0 ? "Rotate Left" : "Rotate Right")
        // The undo is registered when the edit is asked for, as the system expects, and runs after it.
        registerUndo(name, undoManager)
        enqueue { try await $0.rotatePages(pages, degrees: degrees) }
    }

    private func registerUndo(_ name: String, _ undoManager: UndoManager?) {
        guard let undoManager else { return }
        undoManager.registerUndo(withTarget: self) { session in
            session.registerRedo(name, undoManager)
            session.enqueue { try await $0.undo() }
        }
        undoManager.setActionName(name)
    }

    private func registerRedo(_ name: String, _ undoManager: UndoManager) {
        undoManager.registerUndo(withTarget: self) { session in
            session.registerUndo(name, undoManager)
            session.enqueue { try await $0.redo() }
        }
        undoManager.setActionName(name)
    }

    private func enqueue(_ edit: @escaping @Sendable (EngineDocument) async throws -> Void) {
        guard let engine else { return }
        let previous = edits
        isWorking = true
        edits = Task {
            await previous?.value
            do {
                try await edit(engine)
                facts = try await engine.facts()
            } catch {
                errorMessage = error.localizedDescription
            }
            isWorking = false
        }
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
