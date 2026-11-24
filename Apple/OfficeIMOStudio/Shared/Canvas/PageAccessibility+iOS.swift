#if os(iOS)
import UIKit

/// The view the pages are laid out in. VoiceOver reads it as the pages around the one in view: each page, its lines of
/// text, form fields and comments, built when VoiceOver asks and kept while they stay the same, so its focus stays put.
final class AccessiblePagesView: UIView {
    weak var session: DocumentSession?
    var layout: () -> PageLayout = { PageLayout(pages: []) }
    /// Fills a form field as a tap on it does.
    var activateField: ((Int, FormField, FormField.Widget) -> Void)?
    private var cached: (items: [PageAccessibilityItem], elements: [PageElement])?

    override var isAccessibilityElement: Bool {
        get { false }
        set {}
    }

    override var accessibilityElements: [Any]? {
        get {
            guard let session else { return nil }
            let layout = layout()
            let items = session.accessibilityItems(around: session.currentPage, in: layout)
            if cached?.items != items {
                cached = (items, items.map { PageElement(container: self, item: $0) })
            }
            for element in cached?.elements ?? [] { element.place(in: layout) }
            // Views over the pages, such as selection handles, stay reachable after the pages' own elements.
            return (cached?.elements ?? []) + subviews.filter { !$0.isHidden && $0.alpha > 0 }
        }
        set {}
    }

    /// Catches up with what the pages show after page text arrives or a field's value changes, as the Mac does: the
    /// elements VoiceOver holds take their new values, and a different set of elements is announced as a new layout.
    /// Only elements already handed out need this, and only while VoiceOver runs.
    func refreshAccessibility() {
        guard UIAccessibility.isVoiceOverRunning, let session, let cached else { return }
        let layout = layout()
        let items = session.accessibilityItems(around: session.currentPage, in: layout)
        switch ElementRefresh.plan(from: cached.items.map { ($0.identity, $0) }, to: items.map { ($0.identity, $0) }) {
        case .none:
            return
        case .values(let changed):
            self.cached = (items, cached.elements)
            for index in changed { cached.elements[index].show(items[index]) }
            // The focused element's new value is read again; the others are read when VoiceOver reaches them.
            UIAccessibility.post(notification: .layoutChanged, argument: nil)
        case .layout:
            let elements = items.map { PageElement(container: self, item: $0) }
            for element in elements { element.place(in: layout) }
            self.cached = (items, elements)
            UIAccessibility.post(notification: .layoutChanged, argument: nil)
        }
    }

    fileprivate func activate(_ item: PageAccessibilityItem) -> Bool {
        guard let session else { return false }
        switch item.kind {
        case .field(_, let name, let number):
            guard let field = session.formFields(item.page)?.fields.first(where: { $0.name == name }),
                  let widget = field.widgets.first(where: { $0.objectNumber == number }), field.isFillable,
                  session.canFillFormFromPage else { return false }
            activateField?(item.page, field, widget)
            return true
        case .comment(let id):
            guard let thread = session.currentComments?.threads.first(where: { $0.id == id }) else { return false }
            session.selectComment(thread)
            session.showComments()
            return true
        case .page, .text:
            return false
        }
    }
}

/// One page, line, field or comment for VoiceOver, placed over the page in the zooming view.
private final class PageElement: UIAccessibilityElement {
    private(set) var item: PageAccessibilityItem

    init(container: AccessiblePagesView, item: PageAccessibilityItem) {
        self.item = item
        super.init(accessibilityContainer: container)
        show(item)
    }

    /// Takes an item's traits, label and value, also later for the same field or comment with a new value.
    func show(_ item: PageAccessibilityItem) {
        self.item = item
        accessibilityLabel = item.label
        accessibilityValue = item.value
        accessibilityHint = item.hint
        switch item.kind {
        case .page: accessibilityTraits = .header
        case .text: accessibilityTraits = .staticText
        case .field(let kind, _, _):
            accessibilityTraits = switch kind {
            case .checkbox, .radio: .toggleButton
            case .text, .combo, .list: .button
            default: .staticText
            }
        case .comment: accessibilityTraits = .button
        }
    }

    func place(in layout: PageLayout) {
        guard layout.shows(item.page) else { return }
        accessibilityFrameInContainerSpace = layout.layoutRect(item.rect, page: item.page)
    }

    override func accessibilityActivate() -> Bool {
        (accessibilityContainer as? AccessiblePagesView)?.activate(item) ?? false
    }
}
#endif
