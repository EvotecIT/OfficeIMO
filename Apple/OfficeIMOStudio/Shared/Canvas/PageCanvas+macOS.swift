#if os(macOS)
import AppKit
import SwiftUI

/// The page view: the system's magnifying scroll view over the engine's tiles, with trackpad pinch, smart zoom,
/// overlay scrollers and the under-page background, running under the toolbar's glass.
struct PageCanvas: NSViewRepresentable {
    let session: DocumentSession

    func makeNSView(context: Context) -> PageScrollView {
        PageScrollView(session: session)
    }

    func updateNSView(_ view: PageScrollView, context: Context) {
        if let engine = session.engine, let facts = session.facts { view.show(engine, facts: facts) }
        view.tiles.highlight(session.searchHits, current: session.currentHit?.occurrence)
        if let request = session.canvasRequest { view.handle(request) }
        let dragTool = session.pageDragTool
        view.setMarkupActive(session.task == .markup && session.canEditAnnotations)
        view.setSelectionActive(session.task == .read && dragTool == nil
                                && (session.canCopyText || session.canFillForms || session.pendingSignature != nil))
        view.forms.setEnabled(session.task == .read && dragTool == nil && (session.canFillForms || session.pendingSignature != nil))
        view.setDragTool(dragTool)
        _ = session.formFieldsVersion
        view.tiles.markFields(session.fillableFieldBoxes)
        view.tiles.selectionColor = NSColor.selectedTextBackgroundColor.withAlphaComponent(0.6).cgColor
        if let selection = session.textSelection {
            view.tiles.select(page: selection.page, rects: session.selectionRects)
        } else {
            view.tiles.select(page: 0, rects: [])
        }
    }

    static func dismantleNSView(_ view: PageScrollView, coordinator: ()) {
        view.tiles.cancelAll()
    }
}

final class PageScrollView: NSScrollView {
    let tiles = PageTiles()
    private let session: DocumentSession
    private let pages = FlippedView()
    private var handledRequest: UUID?
    private var laidOutWidth: CGFloat = 0
    private var fitZoom: CGFloat = 1
    private var isMagnifying = false
    private lazy var markup = MarkupInkView(session: session) { [weak self] in self?.tiles.layout ?? PageLayout(pages: []) }
    private lazy var selection = TextSelectionView(session: session) { [weak self] in self?.tiles.layout ?? PageLayout(pages: []) }
    private(set) lazy var forms = FormEditor(session: session, pages: pages) { [weak self] in self?.tiles.layout ?? PageLayout(pages: []) }
    private lazy var drag = PageDragView(session: session) { [weak self] in self?.tiles.layout ?? PageLayout(pages: []) }

    init(session: DocumentSession) {
        self.session = session
        super.init(frame: .zero)
        contentView = CenteringClipView()
        contentView.postsBoundsChangedNotifications = true
        documentView = pages
        pages.wantsLayer = true
        pages.layer?.addSublayer(tiles.layer)
        tiles.revisionShown = { [weak self] revision in
            self?.markup.revisionShown(revision)
            self?.drag.revisionShown(revision)
        }
        hasVerticalScroller = true
        hasHorizontalScroller = true
        autohidesScrollers = true
        allowsMagnification = true
        backgroundColor = .underPageBackgroundColor
        let center = NotificationCenter.default
        center.addObserver(self, selector: #selector(viewChanged), name: NSView.boundsDidChangeNotification, object: contentView)
        center.addObserver(self, selector: #selector(magnifyStarted), name: NSScrollView.willStartLiveMagnifyNotification, object: self)
        center.addObserver(self, selector: #selector(magnifyEnded), name: NSScrollView.didEndLiveMagnifyNotification, object: self)
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    func show(_ engine: EngineDocument, facts: DocumentFacts) {
        let previous = tiles.layout
        tiles.show(engine, facts: facts)
        // A new revision of the same pages renders in place; the earlier images stay until the new ones cover them.
        guard tiles.layout != previous else {
            updateTiles()
            return
        }
        pages.frame = CGRect(origin: .zero, size: tiles.layout.size)
        markup.frame = pages.bounds
        drag.frame = pages.bounds
        window?.invalidateCursorRects(for: markup)
        window?.invalidateCursorRects(for: drag)
        laidOutWidth = 0
        needsLayout = true
    }

    func handle(_ request: CanvasRequest) {
        guard request.id != handledRequest else { return }
        handledRequest = request.id
        guard laidOutWidth > 0 else { return }
        switch request.kind {
        case .showPage(let page): scroll(toPage: page)
        case .zoomIn: setZoom(magnification * 1.25)
        case .zoomOut: setZoom(magnification / 1.25)
        case .actualSize: setZoom(1)
        case .fitWidth: setZoom(fitZoom)
        case .reveal(let page, let rect): reveal(rect, onPage: page)
        }
    }

    /// Scrolls a rectangle of a page into the middle of the visible area, unless it is already in view.
    private func reveal(_ rect: CGRect, onPage page: Int) {
        guard tiles.layout.frames.indices.contains(page - 1) else { return }
        let frame = tiles.layout.frames[page - 1]
        let target = rect.offsetBy(dx: frame.minX, dy: frame.minY)
        var visible = pages.visibleRect
        let hidden = contentInsets.top / magnification
        visible.origin.y += hidden
        visible.size.height -= hidden
        guard !visible.insetBy(dx: 8, dy: 8).contains(target) else { return }
        let origin = CGPoint(x: target.midX - visible.width / 2, y: target.midY - visible.height / 2 - hidden)
        contentView.scroll(to: contentView.constrainBoundsRect(CGRect(origin: origin, size: contentView.bounds.size)).origin)
        reflectScrolledClipView(contentView)
    }

    override func layout() {
        super.layout()
        let available = contentSize.width
        guard tiles.layout.size.width > 0, available > 0 else { return }
        if available != laidOutWidth {
            // A new width keeps the page and the zoom relative to the width.
            let page = session.currentPage
            laidOutWidth = available
            fitZoom = available / tiles.layout.size.width
            minMagnification = min(fitZoom / 2, 0.25)
            maxMagnification = max(fitZoom * 4, 8)
            magnification = min(max(fitZoom * session.zoomToFit, minMagnification), maxMagnification)
            scroll(toPage: page)
            viewChanged()
        }
        updateTiles()
    }

    override func viewDidChangeBackingProperties() {
        super.viewDidChangeBackingProperties()
        updateTiles()
    }

    /// Adds the text selection view over the pages while reading.
    func setSelectionActive(_ active: Bool) {
        if active, selection.superview == nil {
            selection.forms = forms
            selection.frame = pages.bounds
            selection.autoresizingMask = [.width, .height]
            pages.addSubview(selection)
        } else if !active, selection.superview != nil {
            selection.removeFromSuperview()
            if session.textSelection != nil { session.textSelection = nil }
        }
    }

    /// Adds the drawing view over the pages while Markup is open; it stays until its strokes show, then leaves.
    func setMarkupActive(_ active: Bool) {
        if active, markup.superview == nil {
            markup.frame = pages.bounds
            markup.autoresizingMask = [.width, .height]
            pages.addSubview(markup)
            window?.invalidateCursorRects(for: markup)
        } else if !active, markup.superview != nil {
            markup.removeFromSuperview()
        }
    }

    /// Keeps the view for shapes and areas to redact over the others, so a drag with a tool reaches it.
    func setDragTool(_ tool: PageDragTool?) {
        if drag.superview == nil {
            drag.frame = pages.bounds
            drag.autoresizingMask = [.width, .height]
            pages.addSubview(drag)
        } else if tool != nil, pages.subviews.last !== drag {
            pages.addSubview(drag, positioned: .above, relativeTo: nil)
        }
        drag.update(tool: tool)
    }

    // MARK: Zoom and position

    @objc private func magnifyStarted() {
        isMagnifying = true
    }

    @objc private func magnifyEnded() {
        isMagnifying = false
        settleZoom()
    }

    @objc private func viewChanged() {
        updateTiles()
        var visible = pages.visibleRect
        let hidden = contentInsets.top / magnification
        visible.origin.y += hidden
        visible.size.height -= hidden
        let page = tiles.layout.mostVisiblePage(in: visible)
        if page != session.currentPage { session.currentPage = page }
        // Notes and stamps go in the middle of the part of the page in view.
        guard tiles.layout.frames.indices.contains(page - 1) else { return }
        let frame = tiles.layout.frames[page - 1]
        let middle = CGPoint(x: min(max(visible.midX, frame.minX + 24), frame.maxX - 24), y: min(max(visible.midY, frame.minY + 24), frame.maxY - 24))
        session.viewCenterOnPage = CGPoint(x: middle.x - frame.minX, y: middle.y - frame.minY)
    }

    private func setZoom(_ scale: CGFloat) {
        let visible = pages.visibleRect
        setMagnification(min(max(scale, minMagnification), maxMagnification), centeredAt: CGPoint(x: visible.midX, y: visible.midY))
        settleZoom()
    }

    private func settleZoom() {
        session.zoomToFit = magnification / fitZoom
        session.zoom = magnification
        updateTiles()
    }

    private func scroll(toPage page: Int) {
        guard tiles.layout.frames.indices.contains(page - 1) else { return }
        let top = tiles.layout.frames[page - 1].minY - PageLayout.spacing / 2 - contentInsets.top / magnification
        contentView.scroll(to: CGPoint(x: contentView.bounds.minX, y: top))
        reflectScrolledClipView(contentView)
    }

    private func updateTiles() {
        guard let window else { return }
        tiles.update(visible: pages.visibleRect, pixelsPerPoint: magnification * window.backingScaleFactor, settled: !isMagnifying)
    }
}

/// The pages' coordinates start at the top left, as the page layout does.
final class FlippedView: NSView {
    override var isFlipped: Bool { true }
}

/// Centers pages narrower than the window instead of pinning them to the left edge.
final class CenteringClipView: NSClipView {
    override func constrainBoundsRect(_ proposedBounds: NSRect) -> NSRect {
        var rect = super.constrainBoundsRect(proposedBounds)
        if let documentView, rect.width > documentView.frame.width {
            rect.origin.x = (documentView.frame.width - rect.width) / 2
        }
        return rect
    }
}
#endif
