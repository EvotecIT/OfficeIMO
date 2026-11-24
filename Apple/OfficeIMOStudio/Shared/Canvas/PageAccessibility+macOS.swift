#if os(macOS)
import AppKit

/// The view the pages are laid out in. VoiceOver reads it as the pages around the one in view: each page, its lines of
/// text, form fields and comments, built when VoiceOver asks and kept while they stay the same, so its focus stays put.
final class AccessiblePagesView: NSView {
    weak var session: DocumentSession?
    var layout: () -> PageLayout = { PageLayout(pages: []) }
    /// Fills a form field as a click on it does.
    var activateField: ((Int, FormField, FormField.Widget) -> Bool)?
    private var cached: (items: [PageAccessibilityItem], elements: [PageElement])?

    /// The pages' coordinates start at the top left, as the page layout does.
    override var isFlipped: Bool { true }

    override func isAccessibilityElement() -> Bool { false }

    override func accessibilityChildren() -> [Any]? {
        guard let session else { return super.accessibilityChildren() }
        let items = session.accessibilityItems(around: session.currentPage, in: layout())
        if cached?.items != items {
            cached = (items, items.map { PageElement(view: self, item: $0) })
        }
        // Views over the pages, such as a field being typed in, stay reachable after the pages' own elements.
        return (cached?.elements ?? []) + (super.accessibilityChildren() ?? [])
    }

    fileprivate func rect(of item: PageAccessibilityItem) -> CGRect? {
        let layout = layout()
        guard layout.shows(item.page) else { return nil }
        return layout.layoutRect(item.rect, page: item.page)
    }

    fileprivate func press(_ item: PageAccessibilityItem) -> Bool {
        guard let session else { return false }
        switch item.kind {
        case .field(_, let name, let number):
            guard session.canFillFormFromPage, layout().shows(item.page),
                  let field = session.formFields(item.page)?.fields.first(where: { $0.name == name }),
                  let widget = field.widgets.first(where: { $0.objectNumber == number }), field.isFillable else { return false }
            return activateField?(item.page, field, widget) ?? false
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

/// One page, line, field or comment for VoiceOver, placed over the page.
private final class PageElement: NSAccessibilityElement {
    let item: PageAccessibilityItem
    private weak var view: AccessiblePagesView?

    init(view: AccessiblePagesView, item: PageAccessibilityItem) {
        self.item = item
        self.view = view
        super.init()
        setAccessibilityParent(view)
        switch item.kind {
        case .page:
            setAccessibilityRole(.group)
            setAccessibilityLabel(item.label)
            setAccessibilityValue(item.value)
        case .text:
            setAccessibilityRole(.staticText)
            setAccessibilityValue(item.label)
        case .field(let kind, _, _):
            let role: NSAccessibility.Role = switch kind {
            case .text: .textField
            case .checkbox: .checkBox
            case .radio: .radioButton
            case .combo, .list: .popUpButton
            default: .staticText
            }
            setAccessibilityRole(role)
            setAccessibilityLabel(item.label)
            if let isOn = item.isOn { setAccessibilityValue(isOn ? 1 : 0) } else { setAccessibilityValue(item.value) }
            setAccessibilityHelp(item.hint)
        case .comment:
            setAccessibilityRole(.button)
            setAccessibilityLabel(item.label)
            setAccessibilityValue(item.value)
            setAccessibilityHelp(item.hint)
        }
    }

    // AppKit asks for accessibility on the main thread.
    override func accessibilityFrame() -> NSRect {
        let view = view, item = item
        return MainActor.assumeIsolated {
            guard let view, let rect = view.rect(of: item) else { return .zero }
            return NSAccessibility.screenRect(fromView: view, rect: rect)
        }
    }

    override func accessibilityPerformPress() -> Bool {
        let view = view, item = item
        return MainActor.assumeIsolated { view?.press(item) ?? false }
    }
}
#endif
