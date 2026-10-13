#if os(macOS)
import AppKit

/// Filling forms on the Mac in the Read task, as Preview does: click a text field to type in it, a check box or radio
/// button to set it, or a choice for its menu. Return or a click elsewhere adds the value, Tab moves on to the next
/// text field, Escape leaves the field as it was, and each value is one undoable change.
@MainActor
final class FormEditor: NSObject, NSTextFieldDelegate {
    private let session: DocumentSession
    private let layout: () -> PageLayout
    private weak var pages: NSView?
    private var editing: (page: Int, field: FormField, control: NSTextField)?

    init(session: DocumentSession, pages: NSView, layout: @escaping () -> PageLayout) {
        self.session = session
        self.pages = pages
        self.layout = layout
    }

    /// Handles a click on a field, in the pages' coordinates; false when no field is there.
    func click(at location: CGPoint) -> Bool {
        let frames = layout().frames
        guard !frames.isEmpty else { return false }
        let page = layout().page(at: location)
        let frame = frames[page - 1]
        let point = CGPoint(x: location.x - frame.minX, y: location.y - frame.minY)
        if session.pendingSignature != nil, frame.contains(location) {
            finish(commit: true)
            session.placeSignature(page: page, at: point)
            return true
        }
        guard let (field, widget) = session.formField(at: point, page: page) else {
            finish(commit: true)
            return false
        }
        if editing?.field.name != field.name { finish(commit: true) }
        switch field.kind {
        case .text: beginTyping(page: page, field: field, box: widget.bounds.cgRect.offsetBy(dx: frame.minX, dy: frame.minY))
        case .checkbox, .radio: session.toggle(field, widget: widget)
        case .combo, .list: showChoices(field, box: widget.bounds.cgRect.offsetBy(dx: frame.minX, dy: frame.minY))
        default: return false
        }
        return true
    }

    /// Whether a point of the pages is over a fillable field, for the pointer.
    func isField(at location: CGPoint) -> FormField.Kind? {
        let pages = layout()
        guard !pages.frames.isEmpty else { return nil }
        let page = pages.page(at: location)
        let frame = pages.frames[page - 1]
        return session.formField(at: CGPoint(x: location.x - frame.minX, y: location.y - frame.minY), page: page)?.0.kind
    }

    func setEnabled(_ enabled: Bool) {
        if !enabled { finish(commit: true) }
    }

    private func beginTyping(page: Int, field: FormField, box: CGRect) {
        guard editing?.field.name != field.name, let pages else { return }
        let control = field.password ? NSSecureTextField() : NSTextField()
        control.stringValue = field.text
        control.font = .systemFont(ofSize: max(8, min(box.height * (field.multiline ? 0.4 : 0.62), 24)))
        control.isBezeled = false
        control.drawsBackground = true
        control.backgroundColor = .white
        control.textColor = .black
        control.focusRingType = .exterior
        control.usesSingleLineMode = !field.multiline
        control.cell?.wraps = field.multiline
        control.cell?.isScrollable = !field.multiline
        control.delegate = self
        control.setAccessibilityLabel(field.displayName)
        control.contentType = switch field.detail {
        case .email: .emailAddress
        case .phone: .telephoneNumber
        case .givenName: .givenName
        case .familyName: .familyName
        case .fullName: .name
        case .organization: .organizationName
        case .street: .fullStreetAddress
        case .city: .addressCity
        case .postalCode: .postalCode
        case .country: .countryName
        case nil: nil
        }
        control.frame = box
        pages.addSubview(control)
        editing = (page, field, control)
        pages.window?.makeFirstResponder(control)
    }

    func control(_ control: NSControl, textView: NSTextView, doCommandBy selector: Selector) -> Bool {
        switch selector {
        case #selector(NSResponder.cancelOperation(_:)):
            finish(commit: false)
            return true
        case #selector(NSResponder.insertNewline(_:)) where editing?.field.multiline != true:
            finish(commit: true)
            return true
        case #selector(NSResponder.insertTab(_:)):
            moveToField(by: 1)
            return true
        case #selector(NSResponder.insertBacktab(_:)):
            moveToField(by: -1)
            return true
        default:
            return false
        }
    }

    /// Tab and Shift-Tab add the value and move to the next or previous text field, in reading order across the
    /// pages whose fields are loaded.
    private func moveToField(by step: Int) {
        guard let current = editing else { return }
        let order = (1...max(session.pageCount, 1)).flatMap { page in
            (session.formFields(page)?.fields ?? [])
                .filter { $0.isFillable && $0.kind == .text }
                .compactMap { field in field.widgets.first.map { (page: page, field: field, box: $0.bounds.cgRect) } }
                .sorted { ($0.box.minY, $0.box.minX) < ($1.box.minY, $1.box.minX) }
        }
        finish(commit: true)
        guard let index = order.firstIndex(where: { $0.page == current.page && $0.field.name == current.field.name }) else { return }
        let next = order[(index + step + order.count) % order.count]
        let frames = layout().frames
        guard frames.indices.contains(next.page - 1) else { return }
        let frame = frames[next.page - 1]
        beginTyping(page: next.page, field: next.field, box: next.box.offsetBy(dx: frame.minX, dy: frame.minY))
        pages?.scrollToVisible(next.box.offsetBy(dx: frame.minX, dy: frame.minY).insetBy(dx: 0, dy: -24))
    }

    func controlTextDidEndEditing(_ notification: Notification) {
        finish(commit: true)
    }

    func control(_ control: NSControl, textShouldEndEditing fieldEditor: NSText) -> Bool { true }

    func controlTextDidChange(_ notification: Notification) {
        guard let editing, let limit = editing.field.maxLength, limit > 0, editing.control.stringValue.count > limit else { return }
        editing.control.stringValue = String(editing.control.stringValue.prefix(limit))
    }

    /// Adds what was typed to the document and removes the control; the page's new images show the value.
    private func finish(commit: Bool) {
        guard let current = editing else { return }
        editing = nil
        let text = current.control.stringValue
        current.control.delegate = nil
        current.control.removeFromSuperview()
        if commit, text != current.field.text { session.setField(current.field, values: [text]) }
    }

    private func showChoices(_ field: FormField, box: CGRect) {
        guard let pages else { return }
        let menu = NSMenu(title: field.displayName)
        for option in field.options {
            let item = NSMenuItem(title: option.text, action: #selector(choose(_:)), keyEquivalent: "")
            item.target = self
            item.representedObject = ChoiceItem(field: field, value: option.value)
            item.state = field.values.contains(option.value) ? .on : .off
            menu.addItem(item)
        }
        menu.popUp(positioning: nil, at: CGPoint(x: box.minX, y: box.maxY + 2), in: pages)
    }

    @objc private func choose(_ sender: NSMenuItem) {
        guard let choice = sender.representedObject as? ChoiceItem else { return }
        let field = choice.field
        if field.multipleSelection {
            let selected = field.values.contains(choice.value)
            session.setField(field, values: selected ? field.values.filter { $0 != choice.value } : field.values + [choice.value])
        } else {
            session.setField(field, values: [choice.value])
        }
    }
}

private final class ChoiceItem: NSObject {
    let field: FormField
    let value: String

    init(field: FormField, value: String) {
        self.field = field
        self.value = value
    }
}
#endif
