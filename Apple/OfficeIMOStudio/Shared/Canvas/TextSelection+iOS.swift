#if os(iOS)
import UIKit

/// Text selection in the Read task, as in other iOS document views: press and hold on text to select a word, keep
/// dragging to extend, and the system edit menu offers Copy and text markup. A tap elsewhere clears it.
@MainActor
final class TextSelectionController: NSObject, @preconcurrency UIEditMenuInteractionDelegate, UIGestureRecognizerDelegate {
    private let session: DocumentSession
    private unowned let scrollView: PageScrollView
    private let press = UILongPressGestureRecognizer()
    private let tap = UITapGestureRecognizer()
    private lazy var editMenu = UIEditMenuInteraction(delegate: self)

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
    }

    @objc private func tapped(_ gesture: UITapGestureRecognizer) {
        editMenu.dismissMenu()
        if session.textSelection != nil { session.textSelection = nil }
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
                }
            ])
        }
        return UIMenu(children: actions)
    }

    func editMenuInteraction(_ interaction: UIEditMenuInteraction, targetRectFor configuration: UIEditMenuConfiguration) -> CGRect {
        selectionRect() ?? .null
    }
}
#endif
