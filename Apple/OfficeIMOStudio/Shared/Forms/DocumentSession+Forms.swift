import CoreGraphics
import Foundation

extension DocumentSession {
    var canFillForms: Bool { isReady && facts?.canFillForms == true }

    /// Whether a field on the page fills when activated, by a tap or VoiceOver: only in Read, with no page tool such as
    /// Edit Form, a shape or redaction marking in use.
    var canFillFormFromPage: Bool { canFillForms && task == .read && pageDragTool == nil }

    /// A page's form fields, if loaded for the revision shown.
    func formFields(_ page: Int) -> FormFields? {
        formPages.value(page: page, revision: revision)
    }

    /// Loads a page's form fields once per revision, keeping a few pages.
    func loadFormFields(_ page: Int) async {
        guard formFields(page) == nil, let engine, canFillForms || canAuthorForms, (1...max(pageCount, 1)).contains(page) else { return }
        let load = formLoads[page] ?? RevisionRead<FormFields>()
        formLoads[page] = load
        defer {
            if formLoads[page] === load, !load.isLoading { formLoads[page] = nil }
        }
        do {
            let fields = try await load.read(source: engine, revision: { self.revision }) { try await engine.formFields(page: page) }
            guard !Task.isCancelled, isReady, self.engine === engine, fields.revision == revision else { return }
            formPages.store(fields, page: page, revision: fields.revision)
        } catch {
            if !Task.isCancelled, self.engine === engine { errorMessage = error.localizedDescription }
        }
    }

    /// The boxes of fillable fields on loaded pages, in each page's displayed points, to show where they are; while
    /// editing the form, every field's boxes.
    var fillableFieldBoxes: [Int: [CGRect]] {
        if isEditingForm {
            return formPages.values(revision: revision)
                .mapValues { $0.fields.flatMap { $0.widgets.map(\.bounds.cgRect) } }.filter { !$0.value.isEmpty }
        }
        guard canFillForms, task == .read else { return [:] }
        return formPages.values(revision: revision).mapValues(\.fillableBoxes).filter { !$0.value.isEmpty }
    }

    /// The fillable field under a point of a page.
    func formField(at point: CGPoint, page: Int) -> (FormField, FormField.Widget)? {
        guard canFillForms, task == .read else { return nil }
        return formFields(page)?.hit(point)
    }

    /// Sets a field's values as one undoable change; `failed` runs when the document refuses them.
    func setField(_ field: FormField, values: [String], failed: (() -> Void)? = nil) {
        guard canFillForms, values != field.values else { return }
        let before = revision
        // Edits run in order, so with another one queued, such as a field being moved, the fields read before it may no
        // longer be where they were; they are read again instead of carried.
        let isFirstInLine = !isWorking
        edit(String(localized: "Fill Form"), { try await $0.fillForm([field.name: values]) }, then: { session in
            if isFirstInLine { session.carryFormFields(from: before, filling: field.name, with: values) }
        }, failed: failed)
    }

    /// The fields loaded before a fill stay right after it but for the field's values, so the next tap finds them at
    /// once rather than after they load again.
    private func carryFormFields(from before: Int64, filling name: String, with values: [String]) {
        let after = revision
        formPages.carry(from: before, to: after) { $0.filling(name, with: values, revision: after) }
    }

    /// Turns a check box or radio button on, or a check box off again.
    func toggle(_ field: FormField, widget: FormField.Widget) {
        guard let on = widget.onState else { return }
        let value = field.kind == .checkbox && widget.isOn ? "Off" : on
        setField(field, values: [value])
    }
}
