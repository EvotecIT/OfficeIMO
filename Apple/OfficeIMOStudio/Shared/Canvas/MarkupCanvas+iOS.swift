#if os(iOS)
import PencilKit
import UIKit

/// Markup with PencilKit: the system tool picker over a canvas on the current page. Strokes stay editable, with the
/// tool picker's eraser, lasso and undo, until Done, a page change or a zoom adds them to the document as ink
/// annotations; the drawing goes once the page's new images show them.
@MainActor
final class MarkupController: NSObject, PKCanvasViewDelegate, PKToolPickerObserver {
    private let session: DocumentSession
    private unowned let scrollView: PageScrollView
    private let canvas = MarkupCanvasView()
    private let toolPicker = PKToolPicker()
    private let drawingID = UUID()
    private var page = 0
    /// The zoom the strokes were drawn at: the canvas is in screen points, the document in page points.
    private var drawnZoom: CGFloat = 1
    /// Committed strokes stay visible until the current engine revision is shown.
    private var isAwaitingCommitRender = false
    /// How many of the drawing's first strokes were added; strokes drawn meanwhile stay.
    private var addedStrokes = 0
    private var handledCommit: UUID?
    private(set) var isActive = false

    init(session: DocumentSession, scrollView: PageScrollView) {
        self.session = session
        self.scrollView = scrollView
        super.init()
        canvas.delegate = self
        canvas.backgroundColor = .clear
        canvas.isOpaque = false
        canvas.isScrollEnabled = false
        canvas.drawingPolicy = .default
        // Pages are white in both appearances, so ink keeps the colour it was chosen in.
        canvas.overrideUserInterfaceStyle = .light
        toolPicker.overrideUserInterfaceStyle = .unspecified
        toolPicker.addObserver(canvas)
        toolPicker.addObserver(self)
        session.document.setSnapshotPreparation(for: drawingID) { [weak self] in try self?.commitSynchronously() }
    }

    func setActive(_ active: Bool) {
        guard active != isActive else { return }
        isActive = active
        if active {
            scrollView.addSubview(canvas)
            canvas.isUserInteractionEnabled = true
            session.markupUndoManager = canvas.undoManager
            move(to: session.currentPage)
            toolPicker.setVisible(true, forFirstResponder: canvas)
            canvas.becomeFirstResponder()
        } else {
            commit()
            if session.markupUndoManager === canvas.undoManager { session.markupUndoManager = nil }
            toolPicker.setVisible(false, forFirstResponder: canvas)
            canvas.resignFirstResponder()
            // An addition still on its way stays visible until the page shows it.
            canvas.isUserInteractionEnabled = false
            if !isAwaitingCommitRender { canvas.removeFromSuperview() }
        }
    }

    /// Retain a failed draft in the document when SwiftUI replaces this page view.
    func detach() {
        setActive(false)
        session.document.removeSnapshotPreparation(for: drawingID)
    }

    /// Follows the session: a commit request, or the page in view changing.
    func update() {
        guard isActive else { return }
        if let request = session.markupCommitRequest, request != handledCommit {
            handledCommit = request
            commit()
        }
        if session.currentPage != page { move(to: session.currentPage) }
    }

    /// Repositions the canvas after scrolling or a layout change. While a pinch runs the canvas hides; when it ends,
    /// the strokes scale to the new zoom, so they stay where they were drawn on the page.
    func layout(isZooming: Bool) {
        // An inactive canvas still showing an addition on its way follows the page too.
        guard isActive || canvas.superview != nil else { return }
        canvas.isHidden = isZooming
        guard !isZooming else { return }
        let zoom = scrollView.zoomScale
        if zoom != drawnZoom {
            let ratio = zoom / drawnZoom
            canvas.drawing = canvas.drawing.transformed(using: CGAffineTransform(scaleX: ratio, y: ratio))
            drawnZoom = zoom
            rememberDrawing()
        }
        place()
    }

    /// Clears the added drawing once the page shows it.
    func revisionShown(_ revision: Int64) {
        guard isAwaitingCommitRender, revision == session.revision, !session.isWorking else { return }
        isAwaitingCommitRender = false
        canvas.drawing = PKDrawing(strokes: canvas.drawing.strokes.dropFirst(addedStrokes))
        addedStrokes = 0
        canvas.undoManager?.removeAllActions()
        if isActive { session.markupUndoManager = canvas.undoManager }
        rememberDrawing()
        canvas.isUserInteractionEnabled = isActive
        if !isActive {
            if !canvas.drawing.strokes.isEmpty { commit() }
            if !isAwaitingCommitRender { canvas.removeFromSuperview() }
        }
    }

    private func move(to newPage: Int) {
        if page != newPage {
            do { try commitSynchronously() }
            catch {
                session.errorMessage = error.localizedDescription
                // Keep both the canvas and its requests on their original page so the user can retry or undo.
                session.show(page: page)
                return
            }
            // Strokes not yet added belong to the page they were drawn on; they go once added.
            if !isAwaitingCommitRender { canvas.drawing = PKDrawing() }
        }
        page = newPage
        drawnZoom = scrollView.zoomScale
        place()
    }

    private func place() {
        let frames = scrollView.tiles.layout.frames
        guard frames.indices.contains(page - 1) else { return }
        canvas.frame = scrollView.contentView.convert(frames[page - 1], to: scrollView)
    }

    /// Adds the drawing as ink annotations, one per colour, width and opacity, as one undoable change.
    func commit() {
        do { try commitSynchronously() }
        catch { session.errorMessage = error.localizedDescription }
    }

    private func commitSynchronously() throws {
        let strokes = canvas.drawing.strokes
        guard !strokes.isEmpty, !isAwaitingCommitRender, page > 0 else { return }
        guard !session.isWorking else { throw CocoaError(.fileWriteUnknown) }
        let requests = annotationRequests()
        guard !requests.isEmpty else { return }
        session.document.updatePendingDrawing(requests, for: drawingID)
        canvas.isUserInteractionEnabled = false
        if session.markupUndoManager === canvas.undoManager { session.markupUndoManager = nil }
        isAwaitingCommitRender = true
        addedStrokes = strokes.count
        do { try session.commitPendingDrawingSynchronously(for: drawingID) }
        catch {
            isAwaitingCommitRender = false
            addedStrokes = 0
            canvas.isUserInteractionEnabled = isActive
            if isActive { session.markupUndoManager = canvas.undoManager }
            throw error
        }
    }

    private func annotationRequests() -> [AnnotationRequest] {
        let strokes = canvas.drawing.strokes
        guard !strokes.isEmpty, !isAwaitingCommitRender, page > 0 else { return [] }
        var groups: [InkKey: [[EnginePoint]]] = [:]
        var order: [InkKey] = []
        let zoom = drawnZoom
        // The canvas covers the page as shown, which may be turned; the strokes go on the upright page.
        let layout = scrollView.tiles.layout
        guard layout.frames.indices.contains(page - 1) else { return [] }
        let frame = layout.frames[page - 1], drawnPage = page
        for stroke in strokes {
            let points = stroke.path.interpolatedPoints(by: .distance(2)).map { point -> EnginePoint in
                let location = point.location.applying(stroke.transform)
                let upright = layout.pagePoint(CGPoint(x: frame.minX + location.x / zoom, y: frame.minY + location.y / zoom), page: drawnPage)
                return EnginePoint(x: upright.x, y: upright.y)
            }
            guard points.count > 1 else { continue }
            let widths = stroke.path.map(\.size.width)
            let width = (widths.reduce(0, +) / CGFloat(max(widths.count, 1))) / zoom
            let key = InkKey(stroke.ink, width: width)
            if groups[key] == nil { order.append(key) }
            groups[key, default: []].append(points)
        }
        return order.map { key in
            AnnotationRequest(tool: .ink, page: page, strokes: groups[key], author: DocumentSession.author,
                              color: key.color, strokeWidth: key.width, opacity: key.opacity)
        }
    }

    // MARK: PKCanvasViewDelegate

    func canvasViewDrawingDidChange(_ canvasView: PKCanvasView) {
        rememberDrawing()
        // PencilKit has a private undo manager; explicitly mark the system document as changed too.
        session.document.objectWillChange.send()
    }

    private func rememberDrawing() {
        session.document.updatePendingDrawing(annotationRequests(), for: drawingID)
        session.pendingDrawingDidChange()
    }
}

/// Strokes that share an ink become one annotation.
private struct InkKey: Hashable {
    let color: EngineColor
    let width: Double
    let opacity: Double

    init(_ ink: PKInk, width: CGFloat) {
        var red: CGFloat = 0, green: CGFloat = 0, blue: CGFloat = 0, alpha: CGFloat = 1
        let resolved = ink.color.resolvedColor(with: UITraitCollection(userInterfaceStyle: .light))
        resolved.getRed(&red, green: &green, blue: &blue, alpha: &alpha)
        func clamp(_ value: CGFloat) -> Double { (min(max(Double(value), 0), 1) * 1000).rounded() / 1000 }
        color = EngineColor(red: clamp(red), green: clamp(green), blue: clamp(blue))
        // Width to a quarter point, so strokes of one tool group together.
        self.width = max(0.25, (Double(width) * 4).rounded() / 4)
        // The marker is translucent, as PencilKit draws it.
        let markerOpacity = ink.inkType == .marker ? 0.5 : 1
        opacity = max(0.05, clamp(alpha) * markerOpacity)
    }
}

/// The canvas keeps its own undo history for strokes not yet added, so ⌘Z and the tool picker's undo act on them
/// first; once added, the drawing is one change in the window's history.
final class MarkupCanvasView: PKCanvasView {
    private let strokeUndo = UndoManager()
    override var undoManager: UndoManager? { strokeUndo }
}
#endif
