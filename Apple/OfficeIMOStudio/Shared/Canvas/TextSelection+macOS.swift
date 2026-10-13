#if os(macOS)
import AppKit

/// Text selection in the Read task, as Preview's text tool works: drag across text to select it, double-click for a
/// word, Edit ▸ Copy or ⌘C to copy, and the context menu for Copy and text markup. The pointer is an I-beam over text.
@MainActor
final class TextSelectionView: NSView, NSMenuItemValidation {
    private let session: DocumentSession
    private let layout: () -> PageLayout
    /// Form fields under the pointer take clicks before text selection does.
    var forms: FormEditor?

    init(session: DocumentSession, layout: @escaping () -> PageLayout) {
        self.session = session
        self.layout = layout
        super.init(frame: .zero)
        setAccessibilityElement(false)
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    override var isFlipped: Bool { true }
    override var acceptsFirstResponder: Bool { true }

    override func updateTrackingAreas() {
        super.updateTrackingAreas()
        trackingAreas.forEach(removeTrackingArea)
        addTrackingArea(NSTrackingArea(rect: .zero, options: [.mouseMoved, .activeInKeyWindow, .inVisibleRect], owner: self))
    }

    override func mouseMoved(with event: NSEvent) {
        if session.pendingSignature != nil {
            NSCursor.crosshair.set()
            return
        }
        if let kind = forms?.isField(at: convert(event.locationInWindow, from: nil)) {
            (kind == .text ? NSCursor.iBeam : NSCursor.pointingHand).set()
            return
        }
        let over = pageText(at: convert(event.locationInWindow, from: nil)).flatMap { text, point in text.index(nearest: point, reach: 2) } != nil
        (over ? NSCursor.iBeam : NSCursor.arrow).set()
    }

    override func mouseDown(with event: NSEvent) {
        let location = convert(event.locationInWindow, from: nil)
        if forms?.click(at: location) == true {
            session.textSelection = nil
            return
        }
        window?.makeFirstResponder(self)
        guard let (text, point) = pageText(at: location), let index = text.index(nearest: point, reach: 12) else {
            session.textSelection = nil
            return
        }
        let anchor = event.clickCount >= 2 ? text.word(at: index) : index...index
        session.textSelection = TextSelection(page: text.page, revision: text.revision, anchor: anchor, range: anchor)
    }

    override func mouseDragged(with event: NSEvent) {
        guard let selection = session.textSelection, let text = session.pageText(selection.page) else { return }
        let frames = layout().frames
        guard frames.indices.contains(selection.page - 1) else { return }
        let location = convert(event.locationInWindow, from: nil)
        autoscroll(with: event)
        let point = layout().pagePoint(location, page: selection.page)
        if let index = text.index(nearest: point, reach: 80) {
            let extended = selection.extended(to: index)
            if extended != selection { session.textSelection = extended }
        }
    }

    override func mouseUp(with event: NSEvent) {
        // A click without a drag on a single element selects nothing, as a click in text places no selection.
        if let selection = session.textSelection, event.clickCount < 2, selection.range.count == 1, selection.anchor.count == 1 {
            session.textSelection = nil
        }
    }

    override func menu(for event: NSEvent) -> NSMenu? {
        guard session.textSelection != nil else { return nil }
        let menu = NSMenu()
        menu.addItem(withTitle: String(localized: "Copy"), action: #selector(copy(_:)), keyEquivalent: "")
        if session.canEditAnnotations {
            menu.addItem(.separator())
            menu.addItem(withTitle: String(localized: "Highlight"), action: #selector(highlightSelection(_:)), keyEquivalent: "")
            menu.addItem(withTitle: String(localized: "Underline"), action: #selector(underlineSelection(_:)), keyEquivalent: "")
            menu.addItem(withTitle: String(localized: "Strikethrough"), action: #selector(strikeSelection(_:)), keyEquivalent: "")
            menu.addItem(.separator())
            menu.addItem(withTitle: String(localized: "Add Link…"), action: #selector(linkSelection(_:)), keyEquivalent: "")
        }
        if session.canRedact {
            menu.addItem(withTitle: String(localized: "Redact"), action: #selector(redactSelection(_:)), keyEquivalent: "")
        }
        return menu
    }

    @objc func linkSelection(_ sender: Any?) { session.structureSheet = .link }
    @objc func redactSelection(_ sender: Any?) { session.redactSelection() }

    @objc func copy(_ sender: Any?) { session.copySelection() }
    @objc func highlightSelection(_ sender: Any?) { session.markUpSelection(.highlight) }
    @objc func underlineSelection(_ sender: Any?) { session.markUpSelection(.underline) }
    @objc func strikeSelection(_ sender: Any?) { session.markUpSelection(.strikeOut) }

    override func selectAll(_ sender: Any?) {
        guard let text = session.pageText(session.currentPage), text.count > 0 else { return }
        session.textSelection = TextSelection(page: text.page, revision: text.revision, anchor: 0...0, range: 0...(text.count - 1))
    }

    func validateMenuItem(_ menuItem: NSMenuItem) -> Bool {
        switch menuItem.action {
        case #selector(copy(_:)): session.selectedText != nil
        case #selector(selectAll(_:)): session.pageText(session.currentPage) != nil
        case #selector(redactSelection(_:)): session.textSelection != nil && session.canRedact
        default: session.textSelection != nil && session.canEditAnnotations
        }
    }

    /// The loaded page text under a point and the point in that page's displayed points.
    private func pageText(at location: CGPoint) -> (PageText, CGPoint)? {
        let pages = layout()
        guard !pages.frames.isEmpty else { return nil }
        let page = pages.page(at: location)
        let frame = pages.frames[page - 1]
        guard frame.insetBy(dx: -8, dy: -8).contains(location), let text = session.pageText(page) else { return nil }
        return (text, pages.pagePoint(location, page: page))
    }
}
#endif
