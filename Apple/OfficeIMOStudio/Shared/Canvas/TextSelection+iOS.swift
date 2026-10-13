#if os(iOS)
import UIKit

/// Text selection in the Read task, as in other iOS document views: press and hold on text to select a word, keep
/// dragging to extend, then drag either grab handle to adjust it; the system edit menu offers Copy and text markup.
/// A tap elsewhere clears it.
@MainActor
final class TextSelectionController: NSObject, @preconcurrency UIEditMenuInteractionDelegate, UIGestureRecognizerDelegate {
    private let session: DocumentSession
    private unowned let scrollView: PageScrollView
    private let press = UILongPressGestureRecognizer()
    private let tap = UITapGestureRecognizer()
    private lazy var editMenu = UIEditMenuInteraction(delegate: self)
    private let startHandle = SelectionHandle(isStart: true)
    private let endHandle = SelectionHandle(isStart: false)
    /// From the finger to the end of the selection being dragged, so the end follows the finger wherever it took hold.
    private var grabOffset = CGPoint.zero

    init(session: DocumentSession, scrollView: PageScrollView) {
        self.session = session
        self.scrollView = scrollView
        super.init()
        press.minimumPressDuration = 0.35
        press.addTarget(self, action: #selector(pressed(_:)))
        press.delegate = self
        tap.addTarget(self, action: #selector(tapped(_:)))
        tap.require(toFail: scrollView.doubleTap)
        tap.delegate = self
        scrollView.addGestureRecognizer(press)
        scrollView.addGestureRecognizer(tap)
        scrollView.addInteraction(editMenu)
        for handle in [startHandle, endHandle] {
            handle.isHidden = true
            handle.addGestureRecognizer(UIPanGestureRecognizer(target: self, action: #selector(dragged(_:))))
            scrollView.addSubview(handle)
        }
    }

    /// Places the grab handles at the selection's ends; called as the selection changes and the pages scroll or zoom.
    func layoutHandles() {
        let ends = selectionEnds()
        startHandle.isHidden = ends == nil
        endHandle.isHidden = ends == nil
        guard let ends else { return }
        startHandle.tintColor = scrollView.tintColor
        endHandle.tintColor = scrollView.tintColor
        startHandle.place(along: ends.start)
        endHandle.place(along: ends.end)
        scrollView.bringSubviewToFront(startHandle)
        scrollView.bringSubviewToFront(endHandle)
    }

    /// The first line's leading edge and the last line's trailing edge, in the scroll view's coordinates.
    private func selectionEnds() -> (start: CGRect, end: CGRect)? {
        guard let selection = session.textSelection, press.state != .changed else { return nil }
        let rects = session.selectionRects
        let frames = scrollView.tiles.layout.frames
        guard let first = rects.first, let last = rects.last, frames.indices.contains(selection.page - 1) else { return nil }
        let origin = frames[selection.page - 1].origin
        func convert(_ rect: CGRect) -> CGRect { scrollView.contentView.convert(rect.offsetBy(dx: origin.x, dy: origin.y), to: scrollView) }
        let start = convert(CGRect(x: first.minX, y: first.minY, width: 0, height: first.height))
        let end = convert(CGRect(x: last.maxX, y: last.minY, width: 0, height: last.height))
        return (start, end)
    }

    @objc private func dragged(_ gesture: UIPanGestureRecognizer) {
        guard let handle = gesture.view as? SelectionHandle, var selection = session.textSelection,
              let text = session.pageText(selection.page) else { return }
        switch gesture.state {
        case .began:
            editMenu.dismissMenu()
            // The other end stays put while this one moves.
            let fixed = handle.isStart ? selection.range.upperBound : selection.range.lowerBound
            selection = TextSelection(page: selection.page, revision: selection.revision, anchor: fixed...fixed, range: selection.range)
            session.textSelection = selection
            let frames = scrollView.tiles.layout.frames
            guard frames.indices.contains(selection.page - 1) else { return }
            let box = text.box(handle.isStart ? selection.range.lowerBound : selection.range.upperBound)
            let end = CGPoint(x: (handle.isStart ? box.minX : box.maxX) + frames[selection.page - 1].minX,
                              y: box.midY + frames[selection.page - 1].minY)
            let finger = gesture.location(in: scrollView.contentView)
            grabOffset = CGPoint(x: end.x - finger.x, y: end.y - finger.y)
        case .changed:
            let frames = scrollView.tiles.layout.frames
            guard frames.indices.contains(selection.page - 1) else { return }
            let finger = gesture.location(in: scrollView.contentView)
            let location = CGPoint(x: finger.x + grabOffset.x, y: finger.y + grabOffset.y)
            let point = CGPoint(x: location.x - frames[selection.page - 1].minX, y: location.y - frames[selection.page - 1].minY)
            if let index = text.index(nearest: point, reach: 60) {
                let extended = selection.extended(to: index)
                if extended != selection {
                    session.textSelection = extended
                    UISelectionFeedbackGenerator().selectionChanged()
                }
            }
        case .ended, .cancelled:
            showMenu()
        default:
            break
        }
        layoutHandles()
    }

    func setEnabled(_ enabled: Bool) {
        press.isEnabled = enabled
        tap.isEnabled = enabled
        if !enabled {
            editMenu.dismissMenu()
            if session.textSelection != nil { session.textSelection = nil }
        }
    }

    // MARK: Gestures

    @objc private func pressed(_ gesture: UILongPressGestureRecognizer) {
        switch gesture.state {
        case .began:
            editMenu.dismissMenu()
            guard let (page, point) = pagePoint(gesture), let text = session.pageText(page) else {
                if let (page, _) = pagePoint(gesture) { Task { await session.loadPageText(page) } }
                return
            }
            guard let index = text.index(nearest: point, reach: 12) else {
                session.textSelection = nil
                return
            }
            let word = text.word(at: index)
            session.textSelection = TextSelection(page: page, revision: text.revision, anchor: word, range: word)
            UISelectionFeedbackGenerator().selectionChanged()
        case .changed:
            guard let selection = session.textSelection, let text = session.pageText(selection.page) else { return }
            let frames = scrollView.tiles.layout.frames
            guard frames.indices.contains(selection.page - 1) else { return }
            let location = gesture.location(in: scrollView.contentView)
            let point = CGPoint(x: location.x - frames[selection.page - 1].minX, y: location.y - frames[selection.page - 1].minY)
            if let index = text.index(nearest: point, reach: 60) {
                let extended = selection.extended(to: index)
                if extended != selection { session.textSelection = extended }
            }
        case .ended:
            showMenu()
        default:
            break
        }
        layoutHandles()
    }

    @objc private func tapped(_ gesture: UITapGestureRecognizer) {
        editMenu.dismissMenu()
        if session.textSelection != nil { session.textSelection = nil }
        layoutHandles()
    }

    func gestureRecognizerShouldBegin(_ gestureRecognizer: UIGestureRecognizer) -> Bool {
        // A tap clears only an existing selection, so it never takes a tap from anything else.
        gestureRecognizer === tap ? session.textSelection != nil : session.canCopyText
    }

    private func pagePoint(_ gesture: UIGestureRecognizer) -> (Int, CGPoint)? {
        let location = gesture.location(in: scrollView.contentView)
        let layout = scrollView.tiles.layout
        guard !layout.frames.isEmpty else { return nil }
        let page = layout.page(atY: location.y)
        let frame = layout.frames[page - 1]
        guard frame.insetBy(dx: -8, dy: -8).contains(location) else { return nil }
        return (page, CGPoint(x: location.x - frame.minX, y: location.y - frame.minY))
    }

    // MARK: Edit menu

    private func showMenu() {
        guard let rect = selectionRect() else { return }
        let configuration = UIEditMenuConfiguration(identifier: nil, sourcePoint: CGPoint(x: rect.midX, y: rect.minY))
        editMenu.presentEditMenu(with: configuration)
    }

    /// The selection's bounds in the scroll view's coordinates.
    private func selectionRect() -> CGRect? {
        guard let selection = session.textSelection else { return nil }
        let rects = session.selectionRects
        let frames = scrollView.tiles.layout.frames
        guard let first = rects.first, frames.indices.contains(selection.page - 1) else { return nil }
        let bounds = rects.dropFirst().reduce(first) { $0.union($1) }.offsetBy(dx: frames[selection.page - 1].minX, dy: frames[selection.page - 1].minY)
        return scrollView.contentView.convert(bounds, to: scrollView)
    }

    func editMenuInteraction(_ interaction: UIEditMenuInteraction, menuFor configuration: UIEditMenuConfiguration,
                             suggestedActions: [UIMenuElement]) -> UIMenu? {
        guard session.textSelection != nil else { return nil }
        var actions: [UIMenuElement] = [
            UIAction(title: String(localized: "Copy"), image: UIImage(systemName: "doc.on.doc")) { [weak self] _ in
                self?.session.copySelection()
            }
        ]
        if session.canEditAnnotations {
            actions.append(contentsOf: [
                UIAction(title: String(localized: "Highlight"), image: UIImage(systemName: "highlighter")) { [weak self] _ in
                    self?.session.markUpSelection(.highlight)
                },
                UIAction(title: String(localized: "Underline"), image: UIImage(systemName: "underline")) { [weak self] _ in
                    self?.session.markUpSelection(.underline)
                },
                UIAction(title: String(localized: "Strikethrough"), image: UIImage(systemName: "strikethrough")) { [weak self] _ in
                    self?.session.markUpSelection(.strikeOut)
                },
                UIAction(title: String(localized: "Add Link…"), image: UIImage(systemName: "link")) { [weak self] _ in
                    self?.session.structureSheet = .link
                }
            ])
        }
        if session.canRedact {
            actions.append(UIAction(title: String(localized: "Redact"), image: UIImage(systemName: "rectangle.fill.badge.xmark"),
                                    attributes: .destructive) { [weak self] _ in
                self?.session.redactSelection()
            })
        }
        return UIMenu(children: actions)
    }

    func editMenuInteraction(_ interaction: UIEditMenuInteraction, targetRectFor configuration: UIEditMenuConfiguration) -> CGRect {
        selectionRect() ?? .null
    }
}

/// A grab handle at one end of a text selection, as the system draws them: a line along the text's height with a
/// round knob above the start or below the end. Its touch area is larger than it looks.
private final class SelectionHandle: UIView {
    let isStart: Bool
    private let stem = UIView()
    private let knob = UIView()

    init(isStart: Bool) {
        self.isStart = isStart
        super.init(frame: CGRect(x: 0, y: 0, width: 44, height: 44))
        stem.isUserInteractionEnabled = false
        knob.isUserInteractionEnabled = false
        knob.layer.cornerRadius = 5
        addSubview(stem)
        addSubview(knob)
        isAccessibilityElement = false
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    override func tintColorDidChange() {
        super.tintColorDidChange()
        stem.backgroundColor = tintColor
        knob.backgroundColor = tintColor
    }

    /// Lines the handle up with an edge of the selection, given as a zero-width rectangle the height of the line.
    func place(along edge: CGRect) {
        let height = max(edge.height, 12)
        frame = CGRect(x: edge.minX - 22, y: edge.minY - 16, width: 44, height: height + 32)
        stem.frame = CGRect(x: 21, y: 16, width: 2, height: height)
        knob.frame = CGRect(x: 17, y: isStart ? 6 : 16 + height, width: 10, height: 10)
        stem.backgroundColor = tintColor
        knob.backgroundColor = tintColor
    }
}
#endif
