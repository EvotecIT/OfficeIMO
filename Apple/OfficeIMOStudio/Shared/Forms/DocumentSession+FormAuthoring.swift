import CoreGraphics
import Foundation

/// A field to add, as officeimo_document_edit_json takes it.
struct FormFieldAdd: Encodable, Sendable {
    let kind: String
    let name: String
    let page: Int
    let bounds: SearchResult.Rectangle
    var options: [String]?
    var combo: Bool?
    var required = false
    var readOnly = false
    var multiline = false
    var label: String?
    var caption: String?
}

/// A field's new name and flags, or alone its new box.
struct FormFieldUpdate: Encodable, Sendable {
    let name: String
    var newName: String?
    var required: Bool?
    var readOnly: Bool?
    var page: Int?
    var bounds: SearchResult.Rectangle?
}

/// The field selected while editing the form.
struct FormFieldSelection: Equatable {
    let name: String
    let page: Int
}

/// The sheet that names a field: a new one drawn on a page, or the selected one.
enum FormFieldSheet: Identifiable {
    case new(NewFormFieldKind, page: Int, rect: CGRect)
    case edit(FormField, page: Int)

    var id: String {
        switch self {
        case .new(let kind, let page, let rect): "new-\(kind.rawValue)-\(page)-\(rect.minX)-\(rect.minY)"
        case .edit(let field, _): "edit-\(field.name)"
        }
    }
}

extension DocumentSession {
    var canAuthorForms: Bool { isReady && facts?.canAuthorForms == true }

    /// Whether a drag on the page adds and arranges form fields.
    var isEditingForm: Bool { formFieldTool != nil && canAuthorForms && task == .read }

    /// Starts editing the form in the Read task, adding fields of the kind chosen.
    func beginEditingForm(_ kind: NewFormFieldKind = .text) {
        guard canAuthorForms else { return }
        if task != .read { task = .read }
        textSelection = nil
        redactionAreas = nil
        pendingSignature = nil
        formFieldTool = kind
    }

    func endEditingForm() {
        formFieldTool = nil
        selectedFormField = nil
    }

    /// Every field name the loaded pages show, to name new fields and check renames.
    var knownFormFieldNames: Set<String> {
        Set(formPages.values.filter { $0.revision == revision }.flatMap { $0.fields.map(\.name) })
    }

    /// The selected field, read for the revision shown.
    var selectedFormFieldValue: (field: FormField, widget: FormField.Widget)? {
        guard let selection = selectedFormField, let fields = formFields(selection.page),
              let field = fields.fields.first(where: { $0.name == selection.name }), let widget = field.widgets.first else { return nil }
        return (field, widget)
    }

    /// The topmost field box under a point of a page, in its displayed points, for selecting it.
    func authoredField(at point: CGPoint, page: Int) -> (FormField, FormField.Widget)? {
        guard let fields = formFields(page) else { return nil }
        for field in fields.fields.reversed() {
            if let widget = field.widgets.first(where: { $0.bounds.cgRect.insetBy(dx: -3, dy: -3).contains(point) }) { return (field, widget) }
        }
        return nil
    }

    /// Asks for the name of a field drawn from one point to another of a page, in its displayed points.
    func drawFormField(page: Int, from start: CGPoint, to end: CGPoint) {
        guard isEditingForm, let kind = formFieldTool, let facts, facts.pages.indices.contains(page - 1) else { return }
        let rect = FormAuthoring.box(from: start, to: end, kind: kind, page: facts.pages[page - 1].displaySize)
        formFieldSheet = .new(kind, page: page, rect: rect)
    }

    /// Adds a field as one undoable change and selects it.
    func addFormField(_ field: FormFieldAdd) {
        guard canAuthorForms else { return }
        edit(String(localized: "Add Form Field"), { try await $0.edit(EditRequest(addFormField: field)) }, then: { session in
            session.selectedFormField = FormFieldSelection(name: field.name, page: field.page)
        })
    }

    /// Renames a field or changes whether it is required or read-only, as one undoable change.
    func updateFormField(_ name: String, newName: String, required: Bool, readOnly: Bool, page: Int) {
        guard canAuthorForms else { return }
        let update = FormFieldUpdate(name: name, newName: newName == name ? nil : newName, required: required, readOnly: readOnly)
        edit(String(localized: "Change Form Field"), { try await $0.edit(EditRequest(updateFormField: update)) }, then: { session in
            session.selectedFormField = FormFieldSelection(name: newName, page: page)
        })
    }

    /// Moves a field's box, or resizes it, on its page, as one undoable change.
    func setFormFieldBounds(_ name: String, page: Int, _ rect: CGRect, resizing: Bool, failed: (() -> Void)? = nil) {
        guard canAuthorForms, let facts, facts.pages.indices.contains(page - 1) else {
            failed?()
            return
        }
        let rect = FormAuthoring.clamp(rect, to: facts.pages[page - 1].displaySize)
        let update = FormFieldUpdate(name: name, page: page,
                                     bounds: .init(x: rect.minX, y: rect.minY, width: rect.width, height: rect.height))
        edit(resizing ? String(localized: "Resize Form Field") : String(localized: "Move Form Field"),
             { try await $0.edit(EditRequest(updateFormField: update)) }, failed: failed)
    }

    /// Removes a field and its boxes as one undoable change.
    func removeFormField(_ name: String) {
        guard canAuthorForms else { return }
        if selectedFormField?.name == name { selectedFormField = nil }
        edit(String(localized: "Delete Form Field")) { try await $0.edit(EditRequest(removeFormField: name)) }
    }

    func removeSelectedFormField() {
        if let name = selectedFormField?.name { removeFormField(name) }
    }
}
