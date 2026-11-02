#if os(iOS)
import UIKit

/// Filling forms on iPad and iPhone in the Read task, as Files and Preview do: tap a text field to type in it, with
/// the keyboard, Scribble with Apple Pencil and AutoFill; tap a check box or radio button to set it; tap a choice for
/// its options. Each committed value is one undoable change. A tap also places a signature waiting to be placed.
@MainActor
final class FormController: NSObject, UIGestureRecognizerDelegate, UITextFieldDelegate, UITextViewDelegate {
    private let session: DocumentSession
    private unowned let scrollView: PageScrollView
    private var keyboardFrame: CGRect?
    private let tap = UITapGestureRecognizer()
    /// The field being typed in: its page, field and widget, and the control over it.
    private var editing: (page: Int, field: FormField, widget: FormField.Widget, control: UIView)?

    init(session: DocumentSession, scrollView: PageScrollView) {
        self.session = session
        self.scrollView = scrollView
        super.init()
        tap.addTarget(self, action: #selector(tapped(_:)))
        tap.require(toFail: scrollView.doubleTap)
        tap.delegate = self
        scrollView.addGestureRecognizer(tap)
        NotificationCenter.default.addObserver(self, selector: #selector(keyboardShown(_:)), name: UIResponder.keyboardDidShowNotification, object: nil)
        NotificationCenter.default.addObserver(self, selector: #selector(keyboardHidden), name: UIResponder.keyboardDidHideNotification, object: nil)
    }

    func setEnabled(_ enabled: Bool) {
        tap.isEnabled = enabled
        if !enabled { finish(commit: true) }
    }

    /// Keeps the control over its field while the pages scroll and zoom.
    func layout() {
        guard let editing, let frame = screenFrame(page: editing.page, box: editing.widget.bounds.cgRect) else { return }
        editing.control.frame = frame
        let size = max(9, min(frame.height * (editing.field.multiline ? 0.4 : 0.62), 28))
        (editing.control as? UITextField)?.font = .systemFont(ofSize: size)
        (editing.control as? UITextView)?.font = .systemFont(ofSize: size)
    }

    // MARK: Gestures

    func gestureRecognizerShouldBegin(_ gestureRecognizer: UIGestureRecognizer) -> Bool {
        // Only a tap on a field is the form's; any other tap goes to the page as before.
        if editing != nil { return true }
        guard let (page, point) = pagePoint(gestureRecognizer) else { return false }
        return session.pendingSignature != nil || session.formField(at: point, page: page) != nil
    }

    @objc private func tapped(_ gesture: UITapGestureRecognizer) {
        if session.pendingSignature != nil, let (page, point) = pagePoint(gesture) {
            finish(commit: true)
            UIImpactFeedbackGenerator(style: .light).impactOccurred()
            session.placeSignature(page: page, at: point)
            return
        }
        let target = pagePoint(gesture).flatMap { page, point in session.formField(at: point, page: page).map { (page, $0.0, $0.1) } }
        if let editing, target?.1.name != editing.field.name || target?.0 != editing.page {
            finish(commit: true)
        }
        guard let (page, field, widget) = target else { return }
        switch field.kind {
        case .text: beginTyping(page: page, field: field, widget: widget)
        case .checkbox, .radio:
            UISelectionFeedbackGenerator().selectionChanged()
            session.toggle(field, widget: widget)
        case .combo, .list: showChoices(page: page, field: field, widget: widget)
        default: break
        }
    }

    private func pagePoint(_ gesture: UIGestureRecognizer) -> (Int, CGPoint)? {
        let location = gesture.location(in: scrollView.contentView)
        let layout = scrollView.tiles.layout
        guard !layout.frames.isEmpty else { return nil }
        let page = layout.page(at: location)
        let frame = layout.frames[page - 1]
        guard frame.contains(location) else { return nil }
        return (page, layout.pagePoint(location, page: page))
    }

    private func screenFrame(page: Int, box: CGRect) -> CGRect? {
        let layout = scrollView.tiles.layout
        guard layout.frames.indices.contains(page - 1) else { return nil }
        return scrollView.contentView.convert(layout.layoutRect(box, page: page), to: scrollView)
    }

    // MARK: Typing

    private func beginTyping(page: Int, field: FormField, widget: FormField.Widget) {
        guard editing?.field.name != field.name else { return }
        let control: UIView
        if field.multiline && !field.password {
            let view = UITextView()
            view.text = field.text
            view.delegate = self
            view.textContainerInset = UIEdgeInsets(top: 2, left: 1, bottom: 2, right: 1)
            configure(view, for: field)
            control = view
        } else {
            let text = UITextField()
            text.text = field.text
            text.delegate = self
            // UIKit has no secure UITextView. Password wins over the PDF's multiline flag.
            text.isSecureTextEntry = field.password
            text.returnKeyType = .done
            text.clearButtonMode = .whileEditing
            text.leftView = UIView(frame: CGRect(x: 0, y: 0, width: 3, height: 1))
            text.leftViewMode = .always
            configure(text, for: field)
            control = text
        }
        control.backgroundColor = .systemBackground
        control.layer.borderColor = scrollView.tintColor.cgColor
        control.layer.borderWidth = 1.5
        control.layer.cornerRadius = 3
        control.accessibilityLabel = field.displayName
        // Pages are white in both appearances, so the field is too.
        control.overrideUserInterfaceStyle = .light
        scrollView.addSubview(control)
        editing = (page, field, widget, control)
        layout()
        control.becomeFirstResponder()
        // Moving between fields keeps the keyboard up, so no new notification follows.
        liftAboveKeyboard()
    }

    /// Tells the keyboard and AutoFill what the field asks for.
    private func configure(_ input: UIView, for field: FormField) {
        let (type, keyboard): (UITextContentType?, UIKeyboardType) = switch field.detail {
        case .email: (.emailAddress, .emailAddress)
        case .phone: (.telephoneNumber, .phonePad)
        case .givenName: (.givenName, .default)
        case .familyName: (.familyName, .default)
        case .fullName: (.name, .default)
        case .organization: (.organizationName, .default)
        case .street: (.fullStreetAddress, .default)
        case .city: (.addressCity, .default)
        case .postalCode: (.postalCode, .numbersAndPunctuation)
        case .country: (.countryName, .default)
        case nil: (nil, .default)
        }
        if let field = input as? UITextField {
            field.textContentType = type
            field.keyboardType = keyboard
            field.autocapitalizationType = type == .emailAddress ? .none : .sentences
        } else if let view = input as? UITextView {
            view.textContentType = type
            view.keyboardType = keyboard
        }
    }

    func textField(_ textField: UITextField, shouldChangeCharactersIn range: NSRange, replacementString string: String) -> Bool {
        guard let limit = editing?.field.maxLength, limit > 0, let text = textField.text, let swiftRange = Range(range, in: text) else { return true }
        return text.replacingCharacters(in: swiftRange, with: string).count <= limit
    }

    func textFieldShouldReturn(_ textField: UITextField) -> Bool {
        finish(commit: true)
        return false
    }

    func textFieldDidEndEditing(_ textField: UITextField) { finish(commit: true) }

    func textViewDidEndEditing(_ textView: UITextView) { finish(commit: true) }

    /// Adds what was typed to the document and removes the control; the page's new images show the value.
    private func finish(commit: Bool) {
        guard let current = editing else { return }
        editing = nil
        let text = (current.control as? UITextField)?.text ?? (current.control as? UITextView)?.text ?? ""
        current.control.removeFromSuperview()
        if commit, text != current.field.text { session.setField(current.field, values: [text]) }
    }

    @objc private func keyboardShown(_ notification: Notification) {
        keyboardFrame = notification.userInfo?[UIResponder.keyboardFrameEndUserInfoKey] as? CGRect
        liftAboveKeyboard()
    }

    @objc private func keyboardHidden() {
        keyboardFrame = nil
        scrollView.keyboardInset = 0
    }

    /// The keyboard may cover the field. Room under the pages lets a field near the end scroll above it. The keyboard
    /// reports screen coordinates, and only a docked one reaches the bottom of the view; a floating one needs no room.
    private func liftAboveKeyboard() {
        guard let editing, let window = scrollView.window, let keyboardFrame else { return }
        let keyboard = window.screen.coordinateSpace.convert(keyboardFrame, to: window)
        let view = window.convert(scrollView.bounds, from: scrollView)
        scrollView.keyboardInset = keyboard.intersects(view) && keyboard.maxY >= view.maxY ? view.maxY - keyboard.minY : 0
        let field = window.convert(editing.control.bounds, from: editing.control)
        if field.intersects(keyboard) {
            let offset = CGPoint(x: scrollView.contentOffset.x, y: scrollView.contentOffset.y + field.maxY + 24 - keyboard.minY)
            scrollView.setContentOffset(offset, animated: true)
        }
    }

    // MARK: Choices

    private func showChoices(page: Int, field: FormField, widget: FormField.Widget) {
        guard let frame = screenFrame(page: page, box: widget.bounds.cgRect), var presenter = scrollView.window?.rootViewController else { return }
        while let presented = presenter.presentedViewController { presenter = presented }
        let sheet = UIAlertController(title: field.displayName, message: nil, preferredStyle: .actionSheet)
        for option in field.options {
            let selected = field.values.contains(option.value)
            let action = UIAlertAction(title: option.text, style: .default) { [weak self] _ in
                guard let self else { return }
                if field.multipleSelection {
                    let values = selected ? field.values.filter { $0 != option.value } : field.values + [option.value]
                    session.setField(field, values: values)
                } else {
                    session.setField(field, values: [option.value])
                }
            }
            action.setValue(selected, forKey: "checked")
            sheet.addAction(action)
        }
        if field.editable {
            sheet.addAction(UIAlertAction(title: String(localized: "Type a Value…"), style: .default) { [weak self] _ in
                self?.beginTyping(page: page, field: field, widget: widget)
            })
        }
        sheet.addAction(UIAlertAction(title: String(localized: "Cancel"), style: .cancel))
        sheet.popoverPresentationController?.sourceView = scrollView
        sheet.popoverPresentationController?.sourceRect = frame
        presenter.present(sheet, animated: true)
    }
}
#endif
