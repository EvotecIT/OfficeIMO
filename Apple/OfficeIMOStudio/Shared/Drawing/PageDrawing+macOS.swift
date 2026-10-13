#if os(macOS)
import AppKit
import QuartzCore

/// Drags across the pages on the Mac with a page tool, as Preview's Select, Shapes and Redact tools work: drawing a
/// shape, marking an area to redact, or selecting markup and moving it or resizing it by a corner. The outline follows
/// the pointer; a change stays as a vector outline until the page's new images show it, and marked areas stay shaded
/// until they are redacted or the marking is cancelled. Delete removes the selected markup.
@MainActor
final class PageDragView: NSView, NSMenuItemValidation {
    private let session: DocumentSession
    private let layout: () -> PageLayout
    private let marks = CAShapeLayer()
    private let selectionOutline = CAShapeLayer()
    private let selectionHandles = CAShapeLayer()
    private var draft: CAShapeLayer?
    private var gesture: PageDragGesture?
    private var tool: PageDragTool?
    /// Changes made in the document and waiting for the page's images to show them, by the revision they followed.
    private var waiting: [(revision: Int64, layer: CAShapeLayer)] = []

    init(session: DocumentSession, layout: @escaping () -> PageLayout) {
        self.session = session
        self.layout = layout
        super.init(frame: .zero)
        wantsLayer = true
        Self.style(marks, as: .area)
        selectionOutline.fillColor = nil
        selectionOutline.lineWidth = 1
        selectionOutline.lineDashPattern = [4, 3]
        for layer in [marks, selectionOutline, selectionHandles] { self.layer?.addSublayer(layer) }
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

    /// Follows the session: the tool a drag uses, the areas marked to redact and the selected markup.
    func update(tool: PageDragTool?) {
        if tool != self.tool {
            self.tool = tool
            window?.invalidateCursorRects(for: self)
            setAccessibilityLabel(tool == .redaction ? String(localized: "Redaction area") : String(localized: "Drawing area"))
        }
        let layout = layout()
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        marks.path = DragOutline.path(for: session.redactionAreas ?? [], layout: layout)
        let accent = NSColor.controlAccentColor.cgColor
        selectionOutline.strokeColor = accent
        selectionHandles.fillColor = accent
        if tool == .select, let selected = session.selectedPageAnnotation, layout.frames.indices.contains(selected.page - 1) {
            let paths = SelectionOutline.paths(for: selected.bounds.cgRect, page: layout.frames[selected.page - 1])
            selectionOutline.path = paths.outline
            selectionHandles.path = paths.handles
        } else {
            selectionOutline.path = nil
            selectionHandles.path = nil
        }
        CATransaction.commit()
    }

    override func resetCursorRects() {
        guard let tool else { return }
        for frame in layout().frames { addCursorRect(frame, cursor: tool == .select ? .arrow : .crosshair) }
    }

    override func mouseDown(with event: NSEvent) {
        guard let tool else { return }
        window?.makeFirstResponder(self)
        let point = convert(event.locationInWindow, from: nil)
        guard let gesture = PageDragGesture.begin(tool, at: point, layout: layout(), session: session) else { return }
        self.gesture = gesture
        let shape = CAShapeLayer()
        Self.style(shape, as: gesture.look, color: session.markupColor)
        shape.contentsScale = (window?.backingScaleFactor ?? 2) * 4
        layer?.addSublayer(shape)
        draft = shape
    }

    override func mouseDragged(with event: NSEvent) {
        guard gesture != nil, let draft else { return }
        gesture?.move(to: convert(event.locationInWindow, from: nil))
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        draft.path = gesture?.outline
        CATransaction.commit()
    }

    override func mouseUp(with event: NSEvent) {
        guard var gesture, let draft else { return }
        self.gesture = nil
        self.draft = nil
        gesture.move(to: convert(event.locationInWindow, from: nil))
        let revision = session.revision
        let stays = gesture.finish(session: session, failed: { [weak self, weak draft] in
            draft?.removeFromSuperlayer()
            self?.waiting.removeAll { $0.layer === draft }
        })
        if stays, draft.superlayer != nil { waiting.append((revision, draft)) } else { draft.removeFromSuperlayer() }
    }

    /// Delete and Forward Delete remove the selected markup.
    override func keyDown(with event: NSEvent) {
        if tool == .select, session.selectedPageAnnotation != nil, [51, 117].contains(event.keyCode) {
            delete(nil)
        } else {
            super.keyDown(with: event)
        }
    }

    /// Edit ▸ Delete removes the selected markup.
    @objc func delete(_ sender: Any?) {
        session.removeSelectedAnnotation()
    }

    func validateMenuItem(_ menuItem: NSMenuItem) -> Bool {
        menuItem.action == #selector(delete(_:)) ? tool == .select && session.selectedPageAnnotation != nil : true
    }

    /// Esc stops marking areas to redact and clears the selection, as it cancels other modes.
    override func cancelOperation(_ sender: Any?) {
        if session.redactionAreas != nil { session.redactionAreas = nil }
        session.selectedAnnotation = nil
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

    private static func style(_ layer: CAShapeLayer, as look: PageDragGesture.Look, color: EngineColor? = nil) {
        switch look {
        case .area:
            layer.fillColor = CGColor(gray: 0, alpha: 0.35)
            layer.strokeColor = CGColor(gray: 0, alpha: 0.9)
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
            layer.strokeColor = NSColor.controlAccentColor.cgColor
            layer.lineWidth = 1
            layer.lineDashPattern = [4, 3]
        }
    }
}
#endif
