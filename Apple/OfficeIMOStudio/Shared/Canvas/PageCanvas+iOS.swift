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
        if let request = session.canvasRequest { view.handle(request) }
    }

    static func dismantleUIView(_ view: PageScrollView, coordinator: ()) {
        view.tiles.cancelAll()
    }
}

final class PageScrollView: UIScrollView, UIScrollViewDelegate {
    let tiles = PageTiles()
    private let session: DocumentSession
    private let content = UIView()
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
        let doubleTap = UITapGestureRecognizer(target: self, action: #selector(doubleTapped(_:)))
        doubleTap.numberOfTapsRequired = 2
        addGestureRecognizer(doubleTap)
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    func show(_ engine: EngineDocument, facts: DocumentFacts) {
        let previous = tiles.layout
        tiles.show(engine, facts: facts)
        guard tiles.layout != previous else { return }
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
        }
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
    }

    func scrollViewDidEndZooming(_ scrollView: UIScrollView, with view: UIView?, atScale scale: CGFloat) {
        session.zoomToFit = scale / fitZoom
        session.zoom = scale
        updateTiles(settled: true)
    }

    func scrollViewDidScroll(_ scrollView: UIScrollView) {
        updateTiles(settled: !isZooming && !isZoomBouncing)
        reportPage()
    }

    // MARK: Zoom and position

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
    }
}
#endif
