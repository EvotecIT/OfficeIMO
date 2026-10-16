import CoreGraphics
import Foundation

extension DocumentSession {
    var canFillForms: Bool { isReady && facts?.canFillForms == true }

    /// A page's form fields, if loaded for the revision shown.
    func formFields(_ page: Int) -> FormFields? {
        guard let fields = formPages[page], fields.revision == revision else { return nil }
        return fields
    }

    /// Loads a page's form fields once per revision, keeping a few pages.
    func loadFormFields(_ page: Int) async {
        guard formFields(page) == nil, let engine, canFillForms, (1...max(pageCount, 1)).contains(page) else { return }
        let load = formLoads[page] ?? RevisionRead<FormFields>()
        formLoads[page] = load
        defer {
            if formLoads[page] === load, !load.isLoading { formLoads[page] = nil }
        }
        do {
            let fields = try await load.read(source: engine, revision: { self.revision }) { try await engine.formFields(page: page) }
            guard !Task.isCancelled, isReady, self.engine === engine, fields.revision == revision else { return }
            formPages[page] = fields
            if formPages.count > 8, let far = formPages.keys.max(by: { abs($0 - page) < abs($1 - page) }) { formPages[far] = nil }
            formFieldsVersion += 1
        } catch {
            if !Task.isCancelled, self.engine === engine { errorMessage = error.localizedDescription }
        }
    }

    /// The boxes of fillable fields on loaded pages, in each page's displayed points, to show where they are.
    var fillableFieldBoxes: [Int: [CGRect]] {
        guard canFillForms, task == .read else { return [:] }
        return formPages.filter { $0.value.revision == revision }.mapValues(\.fillableBoxes).filter { !$0.value.isEmpty }
    }

    /// The fillable field under a point of a page.
    func formField(at point: CGPoint, page: Int) -> (FormField, FormField.Widget)? {
        guard canFillForms, task == .read else { return nil }
        return formFields(page)?.hit(point)
    }

    /// Sets a field's values as one undoable change.
    func setField(_ field: FormField, values: [String]) {
        guard canFillForms, values != field.values else { return }
        let before = revision
        edit(String(localized: "Fill Form"), { try await $0.fillForm([field.name: values]) }, then: { session in
            session.carryFormFields(from: before, filling: field.name, with: values)
        })
    }

    /// The fields loaded before a fill stay right after it but for the field's values, so the next tap finds them at
    /// once rather than after they load again.
    private func carryFormFields(from before: Int64, filling name: String, with values: [String]) {
        for (page, fields) in formPages where fields.revision == before {
            formPages[page] = fields.filling(name, with: values, revision: revision)
        }
        formFieldsVersion += 1
    }

    /// Turns a check box or radio button on, or a check box off again.
    func toggle(_ field: FormField, widget: FormField.Widget) {
        guard let on = widget.onState else { return }
        let value = field.kind == .checkbox && widget.isOn ? "Off" : on
        setField(field, values: [value])
    }
}
