#if os(macOS)
import AppKit
import QuartzCore

/// Drags across the pages on the Mac that draw a shape in Markup or mark an area to redact, as Preview's Shapes and
/// Redact tools do. The outline follows the pointer; a shape stays as a vector outline until the page's new images
/// show it, and marked areas stay shaded until they are redacted or the marking is cancelled.
@MainActor
final class PageDragView: NSView {
    private let session: DocumentSession
    private let layout: () -> PageLayout
    private let marks = CAShapeLayer()
    private var draft: CAShapeLayer?
    private var start: CGPoint?
    private var dragPage = 0
    private var tool: PageDragTool?
    /// Shapes added to the document and waiting for the page's images to show them, by the revision they followed.
    private var waiting: [(revision: Int64, layer: CAShapeLayer)] = []

    init(session: DocumentSession, layout: @escaping () -> PageLayout) {
        self.session = session
        self.layout = layout
        super.init(frame: .zero)
        wantsLayer = true
        marks.fillColor = CGColor(gray: 0, alpha: 0.35)
        marks.strokeColor = CGColor(gray: 0, alpha: 0.9)
        marks.lineWidth = 1
        layer?.addSublayer(marks)
        setAccessibilityElement(true)
        setAccessibilityRole(.layoutArea)
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    override var isFlipped: Bool { true }
    override var acceptsFirstResponder: Bool { tool != nil }

    /// Without a tool, clicks go to the views below: text selection, forms and the pen.
    override func hitTest(_ point: NSPoint) -> NSView? {
        tool == nil ? nil : super.hitTest(point)
    }

    /// Follows the session: the tool a drag uses and the areas marked to redact.
    func update(tool: PageDragTool?) {
        if tool != self.tool {
            self.tool = tool
            window?.invalidateCursorRects(for: self)
            setAccessibilityLabel(tool == .redaction ? String(localized: "Redaction area") : String(localized: "Drawing area"))
        }
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        marks.path = DragOutline.path(for: session.redactionAreas ?? [], layout: layout())
        CATransaction.commit()
    }

    override func resetCursorRects() {
        guard tool != nil else { return }
        for frame in layout().frames { addCursorRect(frame, cursor: .crosshair) }
    }

    override func mouseDown(with event: NSEvent) {
        guard let tool else { return }
        let point = convert(event.locationInWindow, from: nil)
        guard let index = layout().frames.firstIndex(where: { $0.contains(point) }) else { return }
        window?.makeFirstResponder(self)
        dragPage = index + 1
        start = point
        let shape = CAShapeLayer()
        if tool == .redaction {
            shape.fillColor = CGColor(gray: 0, alpha: 0.35)
            shape.strokeColor = CGColor(gray: 0, alpha: 0.9)
            shape.lineWidth = 1
        } else {
            shape.fillColor = nil
            shape.strokeColor = CGColor(srgbRed: session.markupColor.red, green: session.markupColor.green, blue: session.markupColor.blue, alpha: 1)
            shape.lineWidth = DocumentSession.shapeStrokeWidth
            shape.lineCap = .round
            shape.lineJoin = .round
            if tool == .shape(.textBox) { shape.lineDashPattern = [4, 3] }
        }
        shape.contentsScale = (window?.backingScaleFactor ?? 2) * 4
        layer?.addSublayer(shape)
        draft = shape
    }

    override func mouseDragged(with event: NSEvent) {
        guard let tool, let start, let draft else { return }
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        draft.path = DragOutline.path(for: tool, from: start, to: clamped(convert(event.locationInWindow, from: nil)))
        CATransaction.commit()
    }

    override func mouseUp(with event: NSEvent) {
        guard let tool, let start, let draft else { return }
        self.start = nil
        self.draft = nil
        let frames = layout().frames
        guard frames.indices.contains(dragPage - 1) else {
            draft.removeFromSuperlayer()
            return
        }
        let origin = frames[dragPage - 1].origin
        let end = clamped(convert(event.locationInWindow, from: nil))
        let from = CGPoint(x: start.x - origin.x, y: start.y - origin.y), to = CGPoint(x: end.x - origin.x, y: end.y - origin.y)
        switch tool {
        case .redaction:
            draft.removeFromSuperlayer()
            session.markRedaction(page: dragPage, from: from, to: to)
        case .shape(let shape):
            waiting.append((session.revision, draft))
            session.addShape(shape, page: dragPage, from: from, to: to, failed: { [weak self, weak draft] in
                draft?.removeFromSuperlayer()
                self?.waiting.removeAll { $0.layer === draft }
            })
        }
    }

    /// Esc stops marking areas to redact, as it cancels other modes.
    override func cancelOperation(_ sender: Any?) {
        if session.redactionAreas != nil { session.redactionAreas = nil }
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

    /// A point kept on the page the drag started on.
    private func clamped(_ point: CGPoint) -> CGPoint {
        let frames = layout().frames
        guard frames.indices.contains(dragPage - 1) else { return point }
        let page = frames[dragPage - 1]
        return CGPoint(x: min(max(point.x, page.minX), page.maxX), y: min(max(point.y, page.minY), page.maxY))
    }
}
#endif
