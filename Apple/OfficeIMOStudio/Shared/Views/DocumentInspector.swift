import SwiftUI

/// The document's facts from the engine, in the system's grouped form.
struct DocumentInspector: View {
    let session: DocumentSession

    var body: some View {
        Form {
            if let facts = session.facts {
                Section("Document") {
                    LabeledContent("Name", value: facts.fileName)
                    LabeledContent("Pages", value: facts.pages.count, format: .number)
                    LabeledContent("File Size", value: Measurement(value: Double(facts.fileSize), unit: UnitInformationStorage.bytes),
                                   format: .byteCount(style: .file))
                }
                PropertiesSection(session: session)
                if facts.pages.indices.contains(session.currentPage - 1) {
                    let page = facts.pages[session.currentPage - 1]
                    Section("Page \(page.number)") {
                        LabeledContent("Size", value: Self.paperSize(page.displaySize))
                        LabeledContent("Rotation", value: Measurement(value: Double(page.rotation), unit: UnitAngle.degrees),
                                       format: .measurement(width: .narrow))
                    }
                }
                AnnotationSection(session: session)
                AttachmentsSection(session: session)
                Section("Permissions") {
                    LabeledContent("Copy Text", value: facts.canExtractText ? "Allowed" : "Not Allowed")
                    LabeledContent("Extract Content", value: facts.canExtractContent ? "Allowed" : "Not Allowed")
                }
            }
        }
        .formStyle(.grouped)
        .task(id: "\(session.currentPage)-\(session.revision)") { await session.loadAnnotations() }
    }

    /// The size in the reader's units: millimetres, or inches where the US measurement system is used.
    static func paperSize(_ size: CGSize) -> String {
        let metric = Locale.current.measurementSystem != .us
        let unit = metric ? UnitLength.millimeters : UnitLength.inches
        let number = FloatingPointFormatStyle<Double>.number.precision(.fractionLength(0...(metric ? 0 : 2)))
        let width = Measurement(value: size.width / 72, unit: UnitLength.inches).converted(to: unit)
        let height = Measurement(value: size.height / 72, unit: UnitLength.inches).converted(to: unit)
        return "\(width.value.formatted(number)) × \(height.formatted(.measurement(width: .abbreviated, usage: .asProvided, numberFormatStyle: number)))"
    }
}

/// The current page's annotations: what each is, who made it and what it says, with Delete and, for notes, Edit.
private struct AnnotationSection: View {
    let session: DocumentSession
    @State private var editing: AnnotationList.Annotation?
    @State private var text = ""

    var body: some View {
        let annotations = session.pageAnnotations?.page == session.currentPage ? session.pageAnnotations?.annotations ?? [] : []
        Section("Annotations on Page \(session.currentPage)") {
            if annotations.isEmpty {
                Text("None").foregroundStyle(.secondary)
            }
            ForEach(annotations) { annotation in
                row(annotation)
                    .contextMenu {
                        if annotation.subtype == "Text" || annotation.subtype == "FreeText" {
                            Button("Edit Text…", systemImage: "pencil") { edit(annotation) }
                        }
                        Button("Delete", systemImage: "trash", role: .destructive) { session.removeAnnotation(annotation) }
                    }
                    #if os(iOS)
                    .swipeActions {
                        Button("Delete", systemImage: "trash", role: .destructive) { session.removeAnnotation(annotation) }
                    }
                    #endif
            }
        }
        .alert("Edit Note", isPresented: Binding { editing != nil } set: { if !$0 { editing = nil } }) {
            TextField("Note", text: $text, axis: .vertical)
            Button("Save") {
                if let editing { session.updateNote(editing, text: text) }
                editing = nil
            }
            Button("Cancel", role: .cancel) { editing = nil }
        }
    }

    private func row(_ annotation: AnnotationList.Annotation) -> some View {
        HStack(alignment: .firstTextBaseline, spacing: 10) {
            Image(systemName: Self.symbol(annotation.subtype))
                .foregroundStyle(annotation.color.map { Color(red: $0.red, green: $0.green, blue: $0.blue) } ?? .secondary)
                .frame(width: 18)
            VStack(alignment: .leading, spacing: 2) {
                Text(Self.name(annotation.subtype))
                if let contents = annotation.contents, !contents.isEmpty {
                    Text(contents).font(.callout).foregroundStyle(.secondary).lineLimit(3)
                }
                if let author = annotation.author, !author.isEmpty {
                    Text(author).font(.caption).foregroundStyle(.tertiary)
                }
            }
        }
        .accessibilityElement(children: .combine)
    }

    private func edit(_ annotation: AnnotationList.Annotation) {
        text = annotation.contents ?? ""
        editing = annotation
    }

    static func name(_ subtype: String) -> LocalizedStringKey {
        switch subtype {
        case "Ink": "Drawing"
        case "Text": "Note"
        case "FreeText": "Text Box"
        case "Highlight": "Highlight"
        case "Underline": "Underline"
        case "StrikeOut": "Strikethrough"
        case "Squiggly": "Squiggly Underline"
        case "Square": "Rectangle"
        case "Circle": "Oval"
        case "Line": "Line"
        case "Polygon": "Shape"
        case "Stamp": "Stamp"
        case "Link": "Link"
        default: LocalizedStringKey(subtype)
        }
    }

    static func symbol(_ subtype: String) -> String {
        switch subtype {
        case "Ink": "scribble"
        case "Text": "note.text"
        case "FreeText": "character.textbox"
        case "Highlight": "highlighter"
        case "Underline", "Squiggly": "underline"
        case "StrikeOut": "strikethrough"
        case "Square": "rectangle"
        case "Circle": "circle"
        case "Line": "line.diagonal"
        case "Link": "link"
        case "Stamp": "seal"
        default: "pencil.tip.crop.circle"
        }
    }
}
