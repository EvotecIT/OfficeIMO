import SwiftUI

/// The document's title, author, subject and keywords. Editing a field changes the document when the field is
/// submitted or loses focus, as one undoable change.
struct PropertiesSection: View {
    let session: DocumentSession
    @State private var draft = DocumentStructure.Properties()
    @State private var loadedRevision: Int64 = -1
    @FocusState private var focused: Field?

    private enum Field { case title, author, subject, keywords }

    var body: some View {
        Section("Properties") {
            if let structure = session.currentStructure {
                field("Title", text: $draft.title, field: .title)
                field("Author", text: $draft.author, field: .author)
                field("Subject", text: $draft.subject, field: .subject)
                field("Keywords", text: $draft.keywords, field: .keywords)
                if let created = Self.date(structure.properties.created) {
                    LabeledContent("Created", value: created, format: .dateTime)
                }
                if let modified = Self.date(structure.properties.modified) {
                    LabeledContent("Modified", value: modified, format: .dateTime)
                }
            } else {
                ProgressView()
            }
        }
        .disabled(!session.canEditProperties)
        .task(id: session.revision) {
            await session.loadStructure()
            if let structure = session.currentStructure, structure.revision != loadedRevision, focused == nil {
                draft = structure.properties
                loadedRevision = structure.revision
            }
        }
        .onChange(of: focused) { previous, _ in
            if previous != nil { commit() }
        }
    }

    private func field(_ title: LocalizedStringKey, text: Binding<String?>, field: Field) -> some View {
        TextField(title, text: Binding { text.wrappedValue ?? "" } set: { text.wrappedValue = $0 })
            .focused($focused, equals: field)
            .onSubmit(commit)
    }

    private func commit() {
        guard var current = session.currentStructure?.properties else { return }
        current.title = draft.title ?? ""
        current.author = draft.author ?? ""
        current.subject = draft.subject ?? ""
        current.keywords = draft.keywords ?? ""
        let saved = session.currentStructure?.properties
        guard (current.title, current.author, current.subject, current.keywords)
                != (saved?.title ?? "", saved?.author ?? "", saved?.subject ?? "", saved?.keywords ?? "") else { return }
        session.setProperties(DocumentStructure.Properties(title: current.title, author: current.author, subject: current.subject,
                                                           keywords: current.keywords))
    }

    private static func date(_ value: String?) -> Date? {
        value.flatMap { try? Date($0, strategy: .iso8601) }
    }
}

/// Files attached to the document: save one where you choose, attach another, or delete one.
struct AttachmentsSection: View {
    let session: DocumentSession

    var body: some View {
        Section("Attachments") {
            let attachments = session.currentStructure?.attachments ?? []
            if attachments.isEmpty {
                Text("None").foregroundStyle(.secondary)
            }
            ForEach(attachments) { attachment in
                LabeledContent {
                    Text(Measurement(value: Double(attachment.size), unit: UnitInformationStorage.bytes), format: .byteCount(style: .file))
                } label: {
                    Label(attachment.fileName, systemImage: "paperclip")
                }
                .contextMenu {
                    Button("Save…", systemImage: "square.and.arrow.down") { session.saveAttachment(attachment) }
                    if session.canEditAttachments {
                        Button("Delete", systemImage: "trash", role: .destructive) { session.removeAttachment(attachment) }
                    }
                }
                #if os(iOS)
                .swipeActions {
                    if session.canEditAttachments {
                        Button("Delete", systemImage: "trash", role: .destructive) { session.removeAttachment(attachment) }
                    }
                    Button("Save", systemImage: "square.and.arrow.down") { session.saveAttachment(attachment) }
                }
                #endif
            }
            if session.canEditAttachments {
                Button("Attach File…", systemImage: "paperclip.badge.ellipsis") { session.isChoosingAttachment = true }
            }
        }
    }
}
