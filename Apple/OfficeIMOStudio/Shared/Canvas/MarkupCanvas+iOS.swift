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
    private var page = 0
    /// The zoom the strokes were drawn at: the canvas is in screen points, the document in page points.
    private var drawnZoom: CGFloat = 1
    /// The revision the last addition started from; its strokes clear once a later one is shown.
    private var clearAfterRevision: Int64?
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
    }

    func setActive(_ active: Bool) {
        guard active != isActive else { return }
        isActive = active
        if active {
            scrollView.addSubview(canvas)
            canvas.isUserInteractionEnabled = true
            session.markupUndoManager = canvas.undoManager
            // One finger draws; two scroll and pinch, as in Notes and Files.
            scrollView.panGestureRecognizer.minimumNumberOfTouches = 2
            move(to: session.currentPage)
            toolPicker.setVisible(true, forFirstResponder: canvas)
            canvas.becomeFirstResponder()
        } else {
            commit()
            toolPicker.setVisible(false, forFirstResponder: canvas)
            canvas.resignFirstResponder()
            // An addition still on its way stays visible until the page shows it.
            canvas.isUserInteractionEnabled = false
            if clearAfterRevision == nil { canvas.removeFromSuperview() }
            scrollView.panGestureRecognizer.minimumNumberOfTouches = 1
        }
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
        }
        place()
    }

    /// Clears the added drawing once the page shows it.
    func revisionShown(_ revision: Int64) {
        guard let clearAfterRevision, revision > clearAfterRevision else { return }
        self.clearAfterRevision = nil
        canvas.drawing = PKDrawing(strokes: canvas.drawing.strokes.dropFirst(addedStrokes))
        addedStrokes = 0
        canvas.undoManager?.removeAllActions()
        session.hasPendingMarkup = !canvas.drawing.strokes.isEmpty
        if !isActive { canvas.removeFromSuperview() }
    }

    private func move(to newPage: Int) {
        if page != newPage {
            commit()
            // Strokes not yet added belong to the page they were drawn on; they go once added.
            if clearAfterRevision == nil { canvas.drawing = PKDrawing() }
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
        let strokes = canvas.drawing.strokes
        guard !strokes.isEmpty, clearAfterRevision == nil, page > 0 else { return }
        var groups: [InkKey: [[EnginePoint]]] = [:]
        var order: [InkKey] = []
        let zoom = drawnZoom
        for stroke in strokes {
            let points = stroke.path.interpolatedPoints(by: .distance(2)).map { point -> EnginePoint in
                let location = point.location.applying(stroke.transform)
                return EnginePoint(x: location.x / zoom, y: location.y / zoom)
            }
            guard points.count > 1 else { continue }
            let widths = stroke.path.map(\.size.width)
            let width = (widths.reduce(0, +) / CGFloat(max(widths.count, 1))) / zoom
            let key = InkKey(stroke.ink, width: width)
            if groups[key] == nil { order.append(key) }
            groups[key, default: []].append(points)
        }
        let requests = order.map { key in
            AnnotationRequest(tool: .ink, page: page, strokes: groups[key], author: DocumentSession.author,
                              color: key.color, strokeWidth: key.width, opacity: key.opacity)
        }
        guard !requests.isEmpty else { return }
        clearAfterRevision = session.revision
        addedStrokes = strokes.count
        // A failed addition keeps the strokes, so the drawing can be added again or erased.
        session.addMarkup(requests, name: String(localized: "Add Drawing"), failed: { [weak self] in
            guard let self else { return }
            clearAfterRevision = nil
            addedStrokes = 0
            if !isActive { canvas.removeFromSuperview() }
        })
    }

    // MARK: PKCanvasViewDelegate

    func canvasViewDrawingDidChange(_ canvasView: PKCanvasView) {
        session.hasPendingMarkup = !canvasView.drawing.strokes.isEmpty
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
