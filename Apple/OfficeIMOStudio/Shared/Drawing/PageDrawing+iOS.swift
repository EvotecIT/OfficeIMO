#if os(iOS)
import UIKit

/// Drags across the pages on iPad and iPhone with a page tool: drawing a shape, marking an area to redact, or
/// selecting markup and moving it or resizing it by a corner. One finger or Apple Pencil drags while two scroll and
/// pinch, as drawing does; a tap selects markup or unmarks a marked area. The outlines sit in the zooming content, so
/// they stay on the page while it scrolls and zooms.
@MainActor
final class PageDragController: NSObject, UIGestureRecognizerDelegate {
    private let session: DocumentSession
    private unowned let scrollView: PageScrollView
    private let pan = UIPanGestureRecognizer()
    private let tap = UITapGestureRecognizer()
    private let marks = CAShapeLayer()
    private let selectionOutline = CAShapeLayer()
    private let selectionHandles = CAShapeLayer()
    private var draft: CAShapeLayer?
    private var gesture: PageDragGesture?
    /// Where the finger or Pencil last touched down, where a drag that follows starts.
    private var touchDown: CGPoint?
    private(set) var tool: PageDragTool?
    /// Changes made in the document and waiting for the page's images to show them, by the revision they followed.
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
        Self.style(marks, as: .area, tint: scrollView.tintColor)
        selectionOutline.fillColor = nil
        selectionOutline.lineWidth = 1
        selectionOutline.lineDashPattern = [4, 3]
        for (index, layer) in [marks, selectionOutline, selectionHandles].enumerated() {
            layer.zPosition = 1 + CGFloat(index) / 10
            scrollView.contentView.layer.addSublayer(layer)
        }
    }

    /// Follows the session: the tool a drag uses, the areas marked to redact and the selected markup.
    func update(tool: PageDragTool?) {
        self.tool = tool
        pan.isEnabled = tool != nil
        tap.isEnabled = tool == .redaction || tool == .select
        if tool == nil { cancelDraft() }
        let layout = scrollView.tiles.layout
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        marks.path = DragOutline.path(for: session.redactionAreas ?? [], layout: layout)
        selectionOutline.strokeColor = scrollView.tintColor.cgColor
        selectionHandles.fillColor = scrollView.tintColor.cgColor
        if tool == .select, let selected = session.selectedPageAnnotation, layout.frames.indices.contains(selected.page - 1) {
            let paths = SelectionOutline.paths(for: selected.bounds.cgRect, page: selected.page, layout: layout)
            selectionOutline.path = paths.outline
            selectionHandles.path = paths.handles
        } else {
            selectionOutline.path = nil
            selectionHandles.path = nil
        }
        CATransaction.commit()
    }

    /// Removes outlines the page's images now show.
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
        if gestureRecognizer === tap, tool == .redaction {
            // A tap while marking only unmarks; elsewhere it goes to the page as before.
            let point = scrollView.tiles.layout.pagePoint(location, page: index + 1)
            return session.redactionAreas?.contains { $0.page == index + 1 && $0.bounds.cgRect.contains(point) } == true
        }
        return true
    }

    @objc private func panned(_ recognizer: UIPanGestureRecognizer) {
        guard let tool else { return }
        let location = recognizer.location(in: scrollView.contentView)
        switch recognizer.state {
        case .began:
            // The drag starts where the finger first touched, before it moved far enough to be a drag.
            begin(tool, at: touchDown ?? location)
            move(to: location)
        case .changed:
            move(to: location)
        case .ended:
            move(to: location)
            finish()
        default:
            cancelDraft()
        }
    }

    @objc private func tapped(_ recognizer: UITapGestureRecognizer) {
        guard let tool else { return }
        let location = recognizer.location(in: scrollView.contentView)
        UISelectionFeedbackGenerator().selectionChanged()
        if tool == .select {
            // Selects the markup under the tap, or clears the selection.
            _ = PageDragGesture.begin(.select, at: location, layout: scrollView.tiles.layout, session: session)
            return
        }
        let layout = scrollView.tiles.layout
        guard let index = layout.frames.firstIndex(where: { $0.contains(location) }) else { return }
        let point = layout.pagePoint(location, page: index + 1)
        session.markRedaction(page: index + 1, from: point, to: point)
    }

    private func begin(_ tool: PageDragTool, at point: CGPoint) {
        cancelDraft()
        guard let gesture = PageDragGesture.begin(tool, at: point, layout: scrollView.tiles.layout, session: session) else { return }
        self.gesture = gesture
        let shape = CAShapeLayer()
        Self.style(shape, as: gesture.look, color: session.markupColor, tint: scrollView.tintColor)
        shape.contentsScale = scrollView.traitCollection.displayScale * 4
        shape.zPosition = 2
        scrollView.contentView.layer.addSublayer(shape)
        draft = shape
    }

    private func move(to point: CGPoint) {
        guard gesture != nil, let draft else { return }
        gesture?.move(to: point)
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        draft.path = gesture?.outline
        CATransaction.commit()
    }

    private func finish() {
        guard let gesture, let draft else { return }
        self.gesture = nil
        self.draft = nil
        if case .draw(.redaction) = gesture.action { UISelectionFeedbackGenerator().selectionChanged() }
        let revision = session.revision
        let stays = gesture.finish(session: session, failed: { [weak self, weak draft] in
            draft?.removeFromSuperlayer()
            self?.waiting.removeAll { $0.layer === draft }
        })
        if stays, draft.superlayer != nil { waiting.append((revision, draft)) } else { draft.removeFromSuperlayer() }
    }

    private func cancelDraft() {
        draft?.removeFromSuperlayer()
        draft = nil
        gesture = nil
    }

    private static func style(_ layer: CAShapeLayer, as look: PageDragGesture.Look, color: EngineColor? = nil, tint: UIColor) {
        switch look {
        case .area:
            layer.fillColor = UIColor(white: 0, alpha: 0.35).cgColor
            layer.strokeColor = UIColor(white: 0, alpha: 0.9).cgColor
            layer.lineWidth = 1
        case .shape(let dashed):
            let color = color ?? EngineColor(red: 0, green: 0, blue: 0)
            layer.fillColor = nil
            layer.strokeColor = CGColor(srgbRed: color.red, green: color.green, blue: color.blue, alpha: 1)
            layer.lineWidth = DocumentSession.shapeStrokeWidth
            layer.lineCap = .round
            layer.lineJoin = .round
            if dashed { layer.lineDashPattern = [4, 3] }
        case .selection:
            layer.fillColor = nil
            layer.strokeColor = tint.cgColor
            layer.lineWidth = 1
            layer.lineDashPattern = [4, 3]
        }
    }
}
#endif
