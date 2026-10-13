#if os(iOS)
import SwiftUI
import UIKit

/// The page view: the system's zooming scroll view over the engine's tiles, with pinch, double tap and the scroll
/// indicators people expect. It runs under the bars, which float over the pages as glass.
struct PageCanvas: UIViewRepresentable {
    let session: DocumentSession

    func makeUIView(context: Context) -> PageScrollView {
        PageScrollView(session: session)
    }

    func updateUIView(_ view: PageScrollView, context: Context) {
        if let engine = session.engine, let facts = session.facts { view.show(engine, facts: facts) }
        view.tiles.highlight(session.searchHits, current: session.currentHit?.occurrence)
        view.setReadingLayout(session.readingLayout, page: session.currentPage)
        if let request = session.canvasRequest { view.handle(request) }
        let dragTool = session.pageDragTool
        let drawing = session.task == .markup && session.canEditAnnotations && dragTool == nil
        view.markup.setActive(drawing)
        view.drag.update(tool: dragTool)
        // One finger draws, drags a shape or marks an area; two scroll and pinch, as in Notes and Files.
        view.panGestureRecognizer.minimumNumberOfTouches = drawing || dragTool != nil ? 2 : 1
        view.textSelection.setEnabled(session.task == .read && dragTool == nil && session.canCopyText)
        view.forms.setEnabled(session.task == .read && dragTool == nil && (session.canFillForms || session.pendingSignature != nil))
        _ = session.formFieldsVersion
        view.tiles.markFields(session.fillableFieldBoxes)
        view.tiles.selectionColor = view.tintColor.withAlphaComponent(0.3).cgColor
        if let selection = session.textSelection {
            view.tiles.select(page: selection.page, rects: session.selectionRects)
        } else {
            view.tiles.select(page: 0, rects: [])
        }
        view.textSelection.layoutHandles()
        // A double tap on the page draws a dot while marking up, rather than zooming.
        view.doubleTap.isEnabled = !drawing && dragTool == nil
        view.markup.update()
    }

    static func dismantleUIView(_ view: PageScrollView, coordinator: ()) {
        view.markup.setActive(false)
        view.tiles.cancelAll()
    }
}

final class PageScrollView: UIScrollView, UIScrollViewDelegate {
    let tiles = PageTiles()
    private let session: DocumentSession
    private let content = UIView()
    /// The view the pages are laid out in, in PDF points.
    var contentView: UIView { content }
    private(set) lazy var markup = MarkupController(session: session, scrollView: self)
    let doubleTap = UITapGestureRecognizer()
    /// In Single Page, a swipe left shows the next page and a swipe right the previous one.
    private let swipes = [UISwipeGestureRecognizer(), UISwipeGestureRecognizer()]
    private(set) lazy var textSelection = TextSelectionController(session: session, scrollView: self)
    private(set) lazy var forms = FormController(session: session, scrollView: self)
    private(set) lazy var drag = PageDragController(session: session, scrollView: self)
    private var handledRequest: UUID?
    private var laidOutWidth: CGFloat = 0
    private var fitZoom: CGFloat = 1

    init(session: DocumentSession) {
        self.session = session
        super.init(frame: .zero)
        delegate = self
        backgroundColor = .secondarySystemBackground
        contentInsetAdjustmentBehavior = .always
        alwaysBounceVertical = true
        bouncesZoom = true
        content.layer.addSublayer(tiles.layer)
        addSubview(content)
        tiles.revisionShown = { [weak self] revision in
            self?.markup.revisionShown(revision)
            self?.drag.revisionShown(revision)
        }
        swipes[0].direction = .left
        swipes[1].direction = .right
        for swipe in swipes {
            swipe.addTarget(self, action: #selector(swiped(_:)))
            swipe.isEnabled = false
            addGestureRecognizer(swipe)
        }
        doubleTap.addTarget(self, action: #selector(doubleTapped(_:)))
        doubleTap.numberOfTapsRequired = 2
        addGestureRecognizer(doubleTap)
        _ = textSelection
        _ = forms
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    func show(_ engine: EngineDocument, facts: DocumentFacts) {
        let previous = tiles.layout
        tiles.show(engine, facts: facts)
        // A new revision of the same pages renders in place; the earlier images stay until the new ones cover them.
        guard tiles.layout != previous else {
            updateTiles(settled: true)
            return
        }
        layoutChanged()
    }

    /// Lays the pages out for reading: Single Page shows the given page, which a swipe across it changes.
    func setReadingLayout(_ mode: ReadingLayout, page: Int) {
        for swipe in swipes { swipe.isEnabled = mode == .singlePage }
        if tiles.setLayout(mode, shownPage: page) { layoutChanged() }
    }

    /// Sizes the content to a new layout; the next layout pass fits the width and shows the current page.
    private func layoutChanged() {
        let zoom = zoomScale
        zoomScale = 1
        content.frame = CGRect(origin: .zero, size: tiles.layout.size)
        contentSize = tiles.layout.size
        zoomScale = zoom
        // A new layout, such as a rotated page, fits the width again on the next layout pass, keeping the page.
        laidOutWidth = 0
        setNeedsLayout()
    }

    func handle(_ request: CanvasRequest) {
        guard request.id != handledRequest else { return }
        handledRequest = request.id
        guard laidOutWidth > 0 else { return }
        switch request.kind {
        case .showPage(let page): scroll(toPage: page, animated: true)
        case .zoomIn: setZoom(zoomScale * 1.25)
        case .zoomOut: setZoom(zoomScale / 1.25)
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
        let visible = convert(bounds.inset(by: adjustedContentInset), to: content)
        guard !visible.insetBy(dx: 8, dy: 8).contains(target) else { return }
        let size = visible.size
        var offset = CGPoint(x: (target.midX - size.width / 2) * zoomScale - adjustedContentInset.left,
                             y: (target.midY - size.height / 2) * zoomScale - adjustedContentInset.top)
        offset.x = min(max(offset.x, -adjustedContentInset.left), max(-adjustedContentInset.left, contentSize.width + adjustedContentInset.right - bounds.width))
        offset.y = min(max(offset.y, -adjustedContentInset.top), max(-adjustedContentInset.top, contentSize.height + adjustedContentInset.bottom - bounds.height))
        setContentOffset(offset, animated: true)
    }

    override func layoutSubviews() {
        super.layoutSubviews()
        let available = bounds.width - safeAreaInsets.left - safeAreaInsets.right
        guard tiles.layout.size.width > 0, available > 0 else { return }
        if available != laidOutWidth {
            // A new width, from rotation, resizing or folding, keeps the page and the zoom relative to the width.
            let page = session.currentPage
            laidOutWidth = available
            fitZoom = available / tiles.layout.size.width
            minimumZoomScale = min(fitZoom / 2, 0.25)
            maximumZoomScale = max(fitZoom * 4, 8)
            zoomScale = min(max(fitZoom * session.zoomToFit, minimumZoomScale), maximumZoomScale)
            centerContent()
            scroll(toPage: page, animated: false)
        }
        centerContent()
        updateTiles(settled: !isZooming && !isZoomBouncing)
        markup.layout(isZooming: isZooming || isZoomBouncing)
        forms.layout()
        textSelection.layoutHandles()
    }

    /// The sidebar and inspector float over the pages, so the width the pages fit changes with the safe area.
    override func safeAreaInsetsDidChange() {
        super.safeAreaInsetsDidChange()
        setNeedsLayout()
    }

    // MARK: UIScrollViewDelegate

    func viewForZooming(in scrollView: UIScrollView) -> UIView? { content }

    func scrollViewDidZoom(_ scrollView: UIScrollView) {
        centerContent()
        updateTiles(settled: false)
        markup.layout(isZooming: true)
        forms.layout()
        textSelection.layoutHandles()
    }

    func scrollViewDidEndZooming(_ scrollView: UIScrollView, with view: UIView?, atScale scale: CGFloat) {
        session.zoomToFit = scale / fitZoom
        session.zoom = scale
        updateTiles(settled: true)
        markup.layout(isZooming: false)
    }

    func scrollViewDidScroll(_ scrollView: UIScrollView) {
        updateTiles(settled: !isZooming && !isZoomBouncing)
        reportPage()
        markup.layout(isZooming: isZooming || isZoomBouncing)
        forms.layout()
        textSelection.layoutHandles()
    }

    // MARK: Zoom and position

    @objc private func swiped(_ gesture: UISwipeGestureRecognizer) {
        // Only while the page fits the width; zoomed in, a swipe scrolls across the page.
        guard contentSize.width <= bounds.width - safeAreaInsets.left - safeAreaInsets.right + 1 else { return }
        if gesture.direction == .left { session.showNextPage() } else { session.showPreviousPage() }
    }

    @objc private func doubleTapped(_ gesture: UITapGestureRecognizer) {
        if zoomScale > fitZoom * 1.05 {
            setZoom(fitZoom)
        } else {
            let point = gesture.location(in: content)
            let scale = min(fitZoom * 2.5, maximumZoomScale)
            let size = CGSize(width: bounds.width / scale, height: bounds.height / scale)
            zoom(to: CGRect(x: point.x - size.width / 2, y: point.y - size.height / 2, width: size.width, height: size.height), animated: true)
        }
    }

    private func setZoom(_ scale: CGFloat) {
        let scale = min(max(scale, minimumZoomScale), maximumZoomScale)
        let visible = convert(bounds.inset(by: adjustedContentInset), to: content)
        let size = CGSize(width: bounds.width / scale, height: bounds.height / scale)
        zoom(to: CGRect(x: visible.midX - size.width / 2, y: visible.midY - size.height / 2, width: size.width, height: size.height), animated: true)
    }

    private func scroll(toPage page: Int, animated: Bool) {
        guard tiles.layout.frames.indices.contains(page - 1) else { return }
        let top = (tiles.layout.frames[page - 1].minY - PageLayout.spacing / 2) * zoomScale - adjustedContentInset.top
        let bottom = max(-adjustedContentInset.top, contentSize.height + adjustedContentInset.bottom - bounds.height)
        setContentOffset(CGPoint(x: contentOffset.x, y: min(top, bottom)), animated: animated)
    }

    /// Centers pages narrower than the view; the system's safe area insets keep them clear of the bars.
    private func centerContent() {
        let available = bounds.width - safeAreaInsets.left - safeAreaInsets.right
        let horizontal = max(0, (available - contentSize.width) / 2)
        let inset = UIEdgeInsets(top: 0, left: horizontal, bottom: 0, right: horizontal)
        if contentInset != inset { contentInset = inset }
    }

    private func updateTiles(settled: Bool) {
        guard window != nil else { return }
        let visible = convert(bounds, to: content)
        tiles.update(visible: visible, pixelsPerPoint: zoomScale * traitCollection.displayScale, settled: settled)
    }

    private func reportPage() {
        let visible = convert(bounds.inset(by: adjustedContentInset), to: content)
        let page = tiles.layout.mostVisiblePage(in: visible)
        if page != session.currentPage { session.currentPage = page }
        guard tiles.layout.frames.indices.contains(page - 1) else { return }
        let frame = tiles.layout.frames[page - 1]
        let middle = CGPoint(x: min(max(visible.midX, frame.minX + 24), frame.maxX - 24), y: min(max(visible.midY, frame.minY + 24), frame.maxY - 24))
        session.viewCenterOnPage = CGPoint(x: middle.x - frame.minX, y: middle.y - frame.minY)
    }
}
#endif
