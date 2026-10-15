import CoreGraphics
import Foundation
import OfficeIMOStudioEngine

/// A page's form fields from officeimo_document_form_fields_json. Boxes are in points of the page as displayed.
struct FormFields: Decodable, Sendable {
    let revision: Int64
    let page: Int
    let canFill: Bool
    let fields: [FormField]

    /// The fillable field and widget under a point of the page, if any.
    func hit(_ point: CGPoint) -> (FormField, FormField.Widget)? {
        for field in fields where field.isFillable {
            if let widget = field.widgets.first(where: { $0.bounds.cgRect.insetBy(dx: -3, dy: -3).contains(point) }) {
                return (field, widget)
            }
        }
        return nil
    }

    /// The same fields at a later revision where one field has new values; filling changes nothing else.
    func filling(_ name: String, with values: [String], revision: Int64) -> FormFields {
        FormFields(revision: revision, page: page, canFill: canFill,
                   fields: fields.map { $0.name == name ? $0.with(values: values) : $0 })
    }

    /// The boxes of the fields that can be filled, to show where they are.
    var fillableBoxes: [CGRect] {
        canFill ? fields.filter(\.isFillable).flatMap { $0.widgets.map(\.bounds.cgRect) } : []
    }
}

struct FormField: Decodable, Sendable, Identifiable {
    struct Option: Decodable, Sendable, Hashable {
        let value: String
        let text: String
    }

    struct Widget: Decodable, Sendable, Hashable {
        let objectNumber: Int
        let bounds: SearchResult.Rectangle
        /// The value that turns a check box or radio button on.
        let onState: String?
        let isOn: Bool
    }

    enum Kind: String, Decodable, Sendable {
        case text, checkbox, radio, combo, list, button, signature, other
    }

    let name: String
    let kind: Kind
    let values: [String]
    let options: [Option]
    let readOnly: Bool
    let required: Bool
    let multiline: Bool
    let password: Bool
    let comb: Bool
    let maxLength: Int?
    let multipleSelection: Bool
    let editable: Bool
    /// The field's tooltip name, which forms use as its label.
    let label: String?
    let widgets: [Widget]

    var id: String { name }
    var text: String { values.first ?? "" }

    func with(values: [String]) -> FormField {
        let widgets = widgets.map { widget in
            Widget(objectNumber: widget.objectNumber, bounds: widget.bounds, onState: widget.onState,
                   isOn: widget.onState.map(values.contains) ?? false)
        }
        return FormField(name: name, kind: kind, values: values, options: options, readOnly: readOnly, required: required,
                         multiline: multiline, password: password, comb: comb, maxLength: maxLength,
                         multipleSelection: multipleSelection, editable: editable, label: label, widgets: widgets)
    }

    var isFillable: Bool {
        !readOnly && [.text, .checkbox, .radio, .combo, .list].contains(kind)
    }

    /// What the field is called for people: its label, or its name with separators read as spaces.
    var displayName: String {
        if let label, !label.isEmpty { return label }
        return name.replacingOccurrences(of: "_", with: " ").replacingOccurrences(of: ".", with: " ")
    }

    /// What kind of personal detail a text field asks for, from its name and label, so AutoFill can offer it.
    var detail: FormDetail? {
        let words = "\(name) \(label ?? "")".lowercased()
        func has(_ parts: String...) -> Bool { parts.contains { words.contains($0) } }
        if has("email", "e-mail") { return .email }
        if has("phone", "tel", "mobile") { return .phone }
        if has("first", "given", "forename") { return .givenName }
        if has("last", "surname", "family") { return .familyName }
        if has("company", "organi") { return .organization }
        if has("street", "address") { return .street }
        if has("city", "town") { return .city }
        if has("zip", "postal", "postcode") { return .postalCode }
        if has("country") { return .country }
        if has("name") { return .fullName }
        return nil
    }
}

/// Personal details AutoFill knows.
enum FormDetail {
    case fullName, givenName, familyName, email, phone, organization, street, city, postalCode, country
}

private struct FormFill: Encodable {
    let values: [String: [String]]
}

extension EngineDocument {
    func formFields(page: Int) async throws -> FormFields {
        let json: String = try await withEngineCancellation { cancellation in
            var result: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_form_fields_json(handle, Int32(page), cancellation, &result, &error), error)
            return takeString(result) ?? "{}"
        }
        return try JSONDecoder().decode(FormFields.self, from: Data(json.utf8))
    }

    /// Sets field values as one undoable change.
    func fillForm(_ values: [String: [String]]) async throws {
        let request = String(decoding: try JSONEncoder().encode(FormFill(values: values)), as: UTF8.self)
        try await withEngineCancellation { cancellation in
            var error: UnsafeMutablePointer<CChar>?
            try check(request.withCString { officeimo_document_fill_form_json(handle, $0, cancellation, &error) }, error)
        }
    }
}
