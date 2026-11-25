import SwiftUI

extension NewFormFieldKind {
    var title: LocalizedStringKey {
        switch self {
        case .text: "Text Field"
        case .checkbox: "Check Box"
        case .choice: "Choice"
        case .radio: "Radio Buttons"
        case .button: "Push Button"
        case .signature: "Signature Field"
        }
    }

    var systemImage: String {
        switch self {
        case .text: "character.cursor.ibeam"
        case .checkbox: "checkmark.square"
        case .choice: "list.bullet"
        case .radio: "circle.inset.filled"
        case .button: "capsule"
        case .signature: "signature"
        }
    }
}

/// The form fields Edit Form adds, as a menu for the Tools and task menus.
struct EditFormMenu: View {
    let session: DocumentSession

    var body: some View {
        Menu("Edit Form", systemImage: "rectangle.and.pencil.and.ellipsis") {
            ForEach(NewFormFieldKind.allCases) { kind in
                Button(kind.title, systemImage: kind.systemImage) { session.beginEditingForm(kind) }
            }
        }
        .disabled(!session.canAuthorForms || session.isWorking)
    }
}

/// Over the pages while editing the form: the kind of field a drag adds, the selected field's Edit and Delete, and
/// Done.
struct FormEditingBanner: View {
    @Bindable var session: DocumentSession

    var body: some View {
        let selected = session.selectedFormFieldValue
        HStack(spacing: 10) {
            Picker("Field", selection: kind) {
                ForEach(NewFormFieldKind.allCases) { Label($0.title, systemImage: $0.systemImage).tag($0) }
            }
            .pickerStyle(.menu)
            .fixedSize()
            if let selected, let page = session.selectedFormField?.page {
                Text(selected.field.name)
                    .lineLimit(1)
                Button("Edit…", systemImage: "pencil") { session.formFieldSheet = .edit(selected.field, page: page) }
                    .labelStyle(.iconOnly)
                Button("Delete", systemImage: "trash", role: .destructive) { session.removeSelectedFormField() }
                    .labelStyle(.iconOnly)
            } else {
                Text(hint)
            }
            Button("Done") { session.endEditingForm() }
                .buttonStyle(.borderless)
        }
        .disabled(session.isWorking)
        .font(.callout.weight(.medium))
        .padding(.horizontal, 14)
        .padding(.vertical, 8)
        .glassEffect(.regular, in: .capsule)
    }

    private var kind: Binding<NewFormFieldKind> {
        Binding { session.formFieldTool ?? .text } set: { session.formFieldTool = $0 }
    }

    private var hint: LocalizedStringKey {
        #if os(iOS)
        "Drag where the field goes, or tap a field to select it."
        #else
        "Drag where the field goes, or click a field to select it."
        #endif
    }
}

/// Names a new field and sets its options, or changes the selected field's name and flags.
struct FormFieldSheetView: View {
    let session: DocumentSession
    let sheet: FormFieldSheet
    @Environment(\.dismiss) private var dismiss
    @State private var name = ""
    @State private var label = ""
    @State private var required = false
    @State private var readOnly = false
    @State private var multiline = false
    @State private var options = ""
    @State private var dropDown = true
    @State private var caption = ""

    var body: some View {
        NavigationStack {
            Form {
                Section {
                    TextField("Name", text: $name)
                        .autocorrectionDisabled()
                    if case .new = sheet {
                        TextField("Label", text: $label, prompt: Text("Shown as the field’s tooltip"))
                    }
                } footer: {
                    if let problem { Text(problem).foregroundStyle(.red) }
                }
                Section {
                    Toggle("Required", isOn: $required)
                    Toggle("Read Only", isOn: $readOnly)
                    if case .new(.text, _, _) = sheet {
                        Toggle("Multiple Lines", isOn: $multiline)
                    }
                }
                if takesOptions {
                    Section {
                        TextField("Options", text: $options, prompt: Text("One option per line"), axis: .vertical)
                            .lineLimit(3...8)
                        if isChoice { Toggle("Drop-Down List", isOn: $dropDown) }
                    } header: {
                        Text("Options")
                    } footer: {
                        if isChoice {
                            Text("A drop-down list shows one option until it is opened; a list shows them all.")
                        } else {
                            Text("Each option is a button, one under another; choosing one clears the others.")
                        }
                    }
                }
                if case .new(.button, _, _) = sheet {
                    Section {
                        TextField("Caption", text: $caption, prompt: Text("Button"))
                    }
                }
                if case .edit(let field, _) = sheet {
                    Section {
                        Button("Delete Field", systemImage: "trash", role: .destructive) {
                            session.removeFormField(field.name)
                            dismiss()
                        }
                    }
                }
            }
            .formStyle(.grouped)
            .navigationTitle(title)
            #if os(iOS)
            .navigationBarTitleDisplayMode(.inline)
            #endif
            .toolbar {
                ToolbarItem(placement: .cancellationAction) {
                    Button("Cancel") { dismiss() }
                }
                ToolbarItem(placement: .confirmationAction) {
                    Button(isNew ? "Add" : "Save") { save() }
                        .disabled(problem != nil || (takesOptions && FormAuthoring.options(from: options).isEmpty))
                }
            }
        }
        #if os(macOS)
        .frame(minWidth: 380, minHeight: takesOptions ? 420 : 280)
        #endif
        .onAppear(perform: fill)
    }

    private var isNew: Bool { if case .new = sheet { true } else { false } }
    private var isChoice: Bool { if case .new(.choice, _, _) = sheet { true } else { false } }
    private var takesOptions: Bool { if case .new(let kind, _, _) = sheet { FormAuthoring.takesOptions(kind) } else { false } }

    private var title: LocalizedStringKey {
        switch sheet {
        case .new(let kind, _, _): "Add \(Text(kind.title))"
        case .edit: "Form Field"
        }
    }

    private var problem: String? {
        var current: String?
        if case .edit(let field, _) = sheet { current = field.name }
        return FormAuthoring.nameProblem(name, existing: session.knownFormFieldNames, renaming: current)
    }

    private func fill() {
        switch sheet {
        case .new:
            name = FormAuthoring.nextName(existing: session.knownFormFieldNames)
        case .edit(let field, _):
            name = field.name
            required = field.required
            readOnly = field.readOnly
        }
    }

    private func save() {
        let name = name.trimmingCharacters(in: .whitespacesAndNewlines)
        switch sheet {
        case .new(let kind, let page, let rect):
            let label = label.trimmingCharacters(in: .whitespacesAndNewlines)
            session.addFormField(FormFieldAdd(
                kind: kind.rawValue, name: name, page: page,
                bounds: .init(x: rect.minX, y: rect.minY, width: rect.width, height: rect.height),
                options: FormAuthoring.takesOptions(kind) ? FormAuthoring.options(from: options) : nil, combo: kind == .choice ? dropDown : nil,
                required: required, readOnly: readOnly, multiline: kind == .text && multiline, label: label.isEmpty ? nil : label,
                caption: kind == .button ? caption.trimmingCharacters(in: .whitespacesAndNewlines) : nil))
        case .edit(let field, let page):
            if name != field.name || required != field.required || readOnly != field.readOnly {
                session.updateFormField(field.name, newName: name, required: required, readOnly: readOnly, page: page)
            }
        }
        dismiss()
    }
}
