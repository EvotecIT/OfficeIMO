import SwiftUI

/// The document's comments in the inspector: a sidebar beside the pages on the Mac and iPad, a sheet on iPhone. Each
/// thread shows the comment and its replies; the selected one takes a reply and can be resolved or opened again, as in
/// the desktop Studio's Comments panel.
struct CommentsInspector: View {
    @Bindable var session: DocumentSession
    @AppStorage(CommentAuthor.defaultsKey) private var authorName = ""

    var body: some View {
        let threads = session.shownCommentThreads
        Form {
            Section {
                Picker("Show", selection: $session.commentFilter) {
                    Text("All").tag(CommentFilter.all)
                    Text("Open").tag(CommentFilter.open)
                    Text("Resolved").tag(CommentFilter.resolved)
                }
                .pickerStyle(.segmented)
                TextField("Author", text: $session.commentAuthorFilter, prompt: Text("Filter by Author"))
                    .autocorrectionDisabled()
                Button("Next Open Comment", systemImage: "arrow.down.circle") { session.showNextOpenComment() }
                    .disabled(!threads.contains { !$0.resolved })
            }
            if threads.isEmpty {
                Section {
                    Text(session.comments == nil ? "Reading comments…" : emptyMessage)
                        .foregroundStyle(.secondary)
                }
            }
            ForEach(threads) { thread in
                Section {
                    CommentThreadView(session: session, thread: thread)
                } header: {
                    if let page = thread.page { Text("Page \(page)") }
                }
            }
            Section {
                TextField("Your Name", text: $authorName, prompt: DocumentSession.author.isEmpty ? Text("Name") : Text(verbatim: DocumentSession.author))
            } header: {
                Text("Your Name")
            } footer: {
                Text("Comments, replies and markup you add carry this name.")
            }
        }
        .formStyle(.grouped)
        .task(id: "\(session.isReady)-\(session.revision)") { await session.loadComments() }
    }

    private var emptyMessage: LocalizedStringKey {
        switch session.commentFilter {
        case .all: "This document has no comments. Add a note in Markup to start one."
        case .open: "No open comments."
        case .resolved: "No resolved comments."
        }
    }
}

/// A comment, its replies, and for the selected thread a reply field with Reply and Resolve.
private struct CommentThreadView: View {
    @Bindable var session: DocumentSession
    let thread: CommentThread

    private var isSelected: Bool { session.selectedComment == thread.id }

    var body: some View {
        Button { session.selectComment(thread) } label: {
            VStack(alignment: .leading, spacing: 8) {
                ForEach(Array(thread.entries.enumerated()), id: \.offset) { _, entry in
                    entryView(entry)
                }
                if thread.resolved {
                    Label("Resolved", systemImage: "checkmark.circle.fill")
                        .font(.caption)
                        .foregroundStyle(.green)
                }
                if thread.orphaned {
                    Text("This reply’s comment is no longer in the document.")
                        .font(.caption)
                        .foregroundStyle(.secondary)
                }
            }
            .frame(maxWidth: .infinity, alignment: .leading)
            .contentShape(.rect)
        }
        .buttonStyle(.plain)
        .accessibilityElement(children: .combine)
        .accessibilityAddTraits(isSelected ? [.isSelected] : [])
        .accessibilityHint(Text("Shows the comment on its page."))
        .contextMenu {
            resolveButton
            if thread.page != nil {
                Button("Show on Page", systemImage: "doc.text.magnifyingglass") { session.selectComment(thread) }
            }
        }
        .listRowBackground(isSelected ? Color.accentColor.opacity(0.12) : nil)
        if isSelected {
            TextField("Reply", text: $session.commentDrafts[thread.id], prompt: Text("Write a reply"), axis: .vertical)
                .lineLimit(1...5)
                .onSubmit { session.reply(to: thread) }
            HStack {
                Button("Reply", systemImage: "arrowshape.turn.up.left") { session.reply(to: thread) }
                    .disabled(!session.canReplyToComments
                              || session.commentDrafts[thread.id].trimmingCharacters(in: .whitespacesAndNewlines).isEmpty)
                Spacer()
                resolveButton
            }
            .buttonStyle(.borderless)
            if session.currentComments?.canReply == false {
                Text("This document’s comments can’t be changed.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }
        }
    }

    @ViewBuilder private var resolveButton: some View {
        if thread.resolved {
            Button("Reopen", systemImage: "arrow.uturn.backward.circle") { session.setResolved(false, thread) }
                .disabled(!session.canReplyToComments)
        } else {
            Button("Resolve", systemImage: "checkmark.circle") { session.setResolved(true, thread) }
                .disabled(!session.canReplyToComments)
        }
    }

    private func entryView(_ entry: CommentThread.Entry) -> some View {
        let author = entry.author.flatMap { $0.isEmpty ? nil : $0 } ?? String(localized: "Unknown Author")
        return VStack(alignment: .leading, spacing: 2) {
            Text(author)
                .font(.subheadline.weight(.semibold))
            if let contents = entry.contents, !contents.isEmpty {
                Text(contents)
                    .font(.callout)
                    .fixedSize(horizontal: false, vertical: true)
            } else if entry.depth == 0 {
                Text(AnnotationKind.name(thread.subtype))
                    .font(.callout)
                    .foregroundStyle(.secondary)
            }
        }
        .padding(.leading, CGFloat(min(entry.depth, 4)) * 14)
        // VoiceOver says who wrote each entry and whether it replies, which the indent shows on screen.
        .accessibilityElement(children: .ignore)
        .accessibilityLabel(entry.depth == 0 ? Text("Comment from \(author)") : Text("Reply from \(author)"))
        .accessibilityValue(Text(entry.contents ?? ""))
    }
}
