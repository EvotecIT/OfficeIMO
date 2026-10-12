#if os(macOS)
import AppKit
import QuartzCore

/// Markup on the Mac, as Preview's Sketch and Draw tools work: drag across a page to draw with the pen or the
/// highlighter in the chosen colour. Each stroke becomes an ink annotation when the mouse is released, one undoable
/// change in the window's history; it shows as a vector layer until the page's new images include it.
@MainActor
final class MarkupInkView: NSView {
    private let session: DocumentSession
    private let layout: () -> PageLayout
    private var stroke: [CGPoint] = []
    private var strokePage = 0
    private var strokeLayer: CAShapeLayer?
    /// Strokes added to the document and waiting for the page's images to show them, by the revision they followed.
    private var waiting: [(revision: Int64, layer: CAShapeLayer)] = []

    init(session: DocumentSession, layout: @escaping () -> PageLayout) {
        self.session = session
        self.layout = layout
        super.init(frame: .zero)
        wantsLayer = true
        setAccessibilityElement(true)
        setAccessibilityRole(.layoutArea)
        setAccessibilityLabel(String(localized: "Drawing area"))
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    override var isFlipped: Bool { true }

    override func resetCursorRects() {
        for frame in layout().frames { addCursorRect(frame, cursor: .crosshair) }
    }

    override func mouseDown(with event: NSEvent) {
        let point = convert(event.locationInWindow, from: nil)
        let frames = layout().frames
        guard let index = frames.firstIndex(where: { $0.contains(point) }) else { return }
        strokePage = index + 1
        stroke = [point]
        let style = session.markupStyle
        let shape = CAShapeLayer()
        shape.fillColor = nil
        shape.strokeColor = CGColor(srgbRed: session.markupColor.red, green: session.markupColor.green, blue: session.markupColor.blue,
                                    alpha: style.opacity)
        shape.lineWidth = style.strokeWidth
        shape.lineCap = .round
        shape.lineJoin = .round
        shape.contentsScale = (window?.backingScaleFactor ?? 2) * 4
        layer?.addSublayer(shape)
        strokeLayer = shape
    }

    override func mouseDragged(with event: NSEvent) {
        guard strokeLayer != nil else { return }
        let frames = layout().frames
        guard frames.indices.contains(strokePage - 1) else { return }
        // The stroke stays on the page it started on.
        let page = frames[strokePage - 1]
        var point = convert(event.locationInWindow, from: nil)
        point.x = min(max(point.x, page.minX), page.maxX)
        point.y = min(max(point.y, page.minY), page.maxY)
        guard let last = stroke.last, hypot(point.x - last.x, point.y - last.y) >= 0.75 else { return }
        stroke.append(point)
        let path = CGMutablePath()
        path.addLines(between: stroke)
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        strokeLayer?.path = path
        CATransaction.commit()
    }

    override func mouseUp(with event: NSEvent) {
        guard let shape = strokeLayer else { return }
        strokeLayer = nil
        let frames = layout().frames
        guard stroke.count > 1, frames.indices.contains(strokePage - 1) else {
            shape.removeFromSuperlayer()
            return
        }
        let origin = frames[strokePage - 1].origin
        let style = session.markupStyle
        let request = AnnotationRequest(tool: .ink, page: strokePage,
                                        strokes: [stroke.map { EnginePoint(x: $0.x - origin.x, y: $0.y - origin.y) }],
                                        author: DocumentSession.author, color: session.markupColor,
                                        strokeWidth: style.strokeWidth, opacity: style.opacity)
        waiting.append((session.revision, shape))
        session.addMarkup([request], name: String(localized: style == .pen ? "Draw" : "Highlight"), failed: { [weak self, weak shape] in
            shape?.removeFromSuperlayer()
            self?.waiting.removeAll { $0.layer === shape }
        })
        stroke = []
    }

    /// Removes strokes the page's images now show.
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
}
#endif
