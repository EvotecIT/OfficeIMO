#if os(iOS)
import UIKit

/// Drags across the pages on iPad and iPhone that draw a shape in Markup or mark an area to redact. One finger or
/// Apple Pencil drags while two scroll and pinch, as drawing does; a tap on a marked area unmarks it. The outlines
/// sit in the zooming content, so they stay on the page while it scrolls and zooms.
@MainActor
final class PageDragController: NSObject, UIGestureRecognizerDelegate {
    private let session: DocumentSession
    private unowned let scrollView: PageScrollView
    private let pan = UIPanGestureRecognizer()
    private let tap = UITapGestureRecognizer()
    private let marks = CAShapeLayer()
    private var draft: CAShapeLayer?
    private var start: CGPoint?
    /// Where the finger or Pencil last touched down, where a drag that follows starts.
    private var touchDown: CGPoint?
    private var dragPage = 0
    private(set) var tool: PageDragTool?
    /// Shapes added to the document and waiting for the page's images to show them, by the revision they followed.
    private var waiting: [(revision: Int64, layer: CAShapeLayer)] = []

    init(session: DocumentSession, scrollView: PageScrollView) {
        self.session = session
        self.scrollView = scrollView
        super.init()
        pan.addTarget(self, action: #selector(panned(_:)))
        pan.maximumNumberOfTouches = 1
        pan.delegate = self
        pan.isEnabled = false
        scrollView.addGestureRecognizer(pan)
        tap.addTarget(self, action: #selector(tapped(_:)))
        tap.delegate = self
        tap.isEnabled = false
        scrollView.addGestureRecognizer(tap)
        marks.fillColor = UIColor(white: 0, alpha: 0.35).cgColor
        marks.strokeColor = UIColor(white: 0, alpha: 0.9).cgColor
        marks.lineWidth = 1
        marks.zPosition = 1
        scrollView.contentView.layer.addSublayer(marks)
    }

    /// Follows the session: the tool a drag uses and the areas marked to redact.
    func update(tool: PageDragTool?) {
        self.tool = tool
        pan.isEnabled = tool != nil
        tap.isEnabled = tool == .redaction
        if tool == nil { cancelDraft() }
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        marks.path = DragOutline.path(for: session.redactionAreas ?? [], layout: scrollView.tiles.layout)
        CATransaction.commit()
    }

    /// Removes shapes the page's images now show.
    func revisionShown(_ revision: Int64) {
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        waiting.removeAll { entry in
            guard revision > entry.revision else { return false }
            entry.layer.removeFromSuperlayer()
            return true
        }
        CATransaction.commit()
    }

    // MARK: Gestures

    func gestureRecognizer(_ gestureRecognizer: UIGestureRecognizer, shouldReceive touch: UITouch) -> Bool {
        if gestureRecognizer === pan { touchDown = touch.location(in: scrollView.contentView) }
        return true
    }

    func gestureRecognizerShouldBegin(_ gestureRecognizer: UIGestureRecognizer) -> Bool {
        let location = gestureRecognizer.location(in: scrollView.contentView)
        guard let index = scrollView.tiles.layout.frames.firstIndex(where: { $0.contains(location) }) else { return false }
        if gestureRecognizer === tap {
            let frame = scrollView.tiles.layout.frames[index]
            let point = CGPoint(x: location.x - frame.minX, y: location.y - frame.minY)
            return session.redactionAreas?.contains { $0.page == index + 1 && $0.bounds.cgRect.contains(point) } == true
        }
        return true
    }

    @objc private func panned(_ gesture: UIPanGestureRecognizer) {
        guard let tool else { return }
        let location = gesture.location(in: scrollView.contentView)
        switch gesture.state {
        case .began:
            // The drag starts where the finger first touched, before it moved far enough to be a drag.
            begin(tool, at: touchDown ?? location)
            move(to: location)
        case .changed:
            move(to: location)
        case .ended:
            finish(tool, at: location)
        default:
            cancelDraft()
        }
    }

    @objc private func tapped(_ gesture: UITapGestureRecognizer) {
        let location = gesture.location(in: scrollView.contentView)
        let layout = scrollView.tiles.layout
        guard let index = layout.frames.firstIndex(where: { $0.contains(location) }) else { return }
        let frame = layout.frames[index]
        let point = CGPoint(x: location.x - frame.minX, y: location.y - frame.minY)
        UISelectionFeedbackGenerator().selectionChanged()
        session.markRedaction(page: index + 1, from: point, to: point)
    }

    private func begin(_ tool: PageDragTool, at point: CGPoint) {
        cancelDraft()
        let frames = scrollView.tiles.layout.frames
        guard let index = frames.firstIndex(where: { $0.contains(point) }) else { return }
        dragPage = index + 1
        start = point
        let shape = CAShapeLayer()
        if tool == .redaction {
            shape.fillColor = UIColor(white: 0, alpha: 0.35).cgColor
            shape.strokeColor = UIColor(white: 0, alpha: 0.9).cgColor
            shape.lineWidth = 1
        } else {
            shape.fillColor = nil
            shape.strokeColor = CGColor(srgbRed: session.markupColor.red, green: session.markupColor.green, blue: session.markupColor.blue, alpha: 1)
            shape.lineWidth = DocumentSession.shapeStrokeWidth
            shape.lineCap = .round
            shape.lineJoin = .round
            if tool == .shape(.textBox) { shape.lineDashPattern = [4, 3] }
        }
        shape.contentsScale = scrollView.traitCollection.displayScale * 4
        shape.zPosition = 2
        scrollView.contentView.layer.addSublayer(shape)
        draft = shape
    }

    private func move(to point: CGPoint) {
        guard let tool, let start, let draft else { return }
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        draft.path = DragOutline.path(for: tool, from: start, to: clamped(point))
        CATransaction.commit()
    }

    private func finish(_ tool: PageDragTool, at point: CGPoint) {
        guard let start, let draft else { return }
        self.start = nil
        self.draft = nil
        let frames = scrollView.tiles.layout.frames
        guard frames.indices.contains(dragPage - 1) else {
            draft.removeFromSuperlayer()
            return
        }
        let origin = frames[dragPage - 1].origin
        let end = clamped(point)
        let from = CGPoint(x: start.x - origin.x, y: start.y - origin.y), to = CGPoint(x: end.x - origin.x, y: end.y - origin.y)
        switch tool {
        case .redaction:
            draft.removeFromSuperlayer()
            UISelectionFeedbackGenerator().selectionChanged()
            session.markRedaction(page: dragPage, from: from, to: to)
        case .shape(let shape):
            waiting.append((session.revision, draft))
            session.addShape(shape, page: dragPage, from: from, to: to, failed: { [weak self, weak draft] in
                draft?.removeFromSuperlayer()
                self?.waiting.removeAll { $0.layer === draft }
            })
        }
    }

    private func cancelDraft() {
        draft?.removeFromSuperlayer()
        draft = nil
        start = nil
    }

    /// A point kept on the page the drag started on.
    private func clamped(_ point: CGPoint) -> CGPoint {
        let frames = scrollView.tiles.layout.frames
        guard frames.indices.contains(dragPage - 1) else { return point }
        let page = frames[dragPage - 1]
        return CGPoint(x: min(max(point.x, page.minX), page.maxX), y: min(max(point.y, page.minY), page.maxY))
    }
}
#endif
