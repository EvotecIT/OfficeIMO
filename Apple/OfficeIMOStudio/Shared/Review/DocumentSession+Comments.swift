import CoreGraphics
import Foundation
import UniformTypeIdentifiers
import OfficeIMOStudioEngine

/// A reply to the comment with this object number, as officeimo_document_edit_json takes it.
struct CommentReply: Encodable, Sendable {
    let objectNumber: Int
    let text: String
    let author: String
    var color: EngineColor?
}

/// A comment's review state: completed resolves it and none opens it again.
struct CommentStateChange: Encodable, Sendable {
    let objectNumber: Int
    let state: String
}

/// What the inspector shows.
enum InspectorPane: String, CaseIterable, Identifiable {
    case details
    case comments

    var id: Self { self }
}

extension EngineDocument {
    /// The document's comment threads, each with its replies and review state.
    func comments() async throws -> CommentList {
        let json: String = try await withEngineCancellation { cancellation in
            var json: UnsafeMutablePointer<CChar>?
            var error: UnsafeMutablePointer<CChar>?
            try check(officeimo_document_comments_json(handle, cancellation, &json, &error), error)
            return takeString(json) ?? "{}"
        }
        return try JSONDecoder().decode(CommentList.self, from: Data(json.utf8))
    }
}

extension DocumentSession {
    /// The comments of the revision shown, once read.
    var currentComments: CommentList? {
        guard let comments, comments.revision == revision else { return nil }
        return comments
    }

    var canReplyToComments: Bool { isReady && !isWorking && currentComments?.canReply == true }

    /// The threads the filter keeps.
    var shownCommentThreads: [CommentThread] {
        commentFilter.apply(to: currentComments?.threads ?? comments?.threads ?? [], author: commentAuthorFilter)
    }

    var selectedCommentThread: CommentThread? {
        guard let selectedComment else { return nil }
        return (currentComments ?? comments)?.threads.first { $0.id == selectedComment }
    }

    /// Reads the comment threads once per revision.
    func loadComments() async {
        guard let engine, isReady, currentComments == nil else { return }
        do {
            let read = try await commentsLoad.read(source: engine, revision: { self.revision }) { try await engine.comments() }
            if !Task.isCancelled, isReady, self.engine === engine, read.revision == revision { comments = read }
        } catch {
            if !Task.isCancelled, self.engine === engine { errorMessage = error.localizedDescription }
        }
    }

    /// Opens the inspector on the comments, as the toolbar's Comments button and the View menu do.
    func showComments() {
        inspectorPane = .comments
        isInspectorShown = true
    }

    /// Selects a thread and shows its page, with the comment brought into view.
    func selectComment(_ thread: CommentThread) {
        selectedComment = thread.id
        guard let page = thread.page, (1...max(pageCount, 1)).contains(page) else { return }
        if let box = thread.bounds {
            currentPage = page
            reveal(page: page, rect: CGRect(x: box.x, y: box.y, width: box.width, height: box.height))
        } else {
            show(page: page)
        }
    }

    /// Selects the next open thread after the selected one.
    func showNextOpenComment() {
        if let next = CommentNavigation.nextOpen(after: selectedComment, in: shownCommentThreads) { selectComment(next) }
    }

    /// Adds the draft reply to the selected thread as one undoable change.
    func reply(to thread: CommentThread) {
        let text = commentDrafts[thread].trimmingCharacters(in: .whitespacesAndNewlines)
        guard canReplyToComments, !text.isEmpty else { return }
        let draft = commentDrafts[thread]
        let reply = CommentReply(objectNumber: thread.objectNumber, text: text, author: Self.author,
                                 color: EngineColor(red: 1, green: 0.8, blue: 0.1))
        edit(String(localized: "Reply to Comment"), { try await $0.edit(EditRequest(reply: reply)) }, then: { session in
            session.commentDrafts.sent(draft, to: thread.id)
            Task { await session.loadComments() }
        })
    }

    /// Reply drafts whose thread is no longer in the revision shown, once its comments are read.
    var retainedCommentDrafts: [CommentDrafts.Draft] {
        guard let comments = currentComments else { return [] }
        return commentDrafts.unassigned(keeping: comments.threads)
    }

    /// Writes the comment threads as a Markdown summary, grouped by page, for the export panel.
    func commentSummaryFile() throws -> ExportedFile? {
        guard let comments = currentComments, !comments.threads.isEmpty, let facts else { return nil }
        let base = (facts.fileName as NSString).deletingPathExtension
        let name = try StudioTemporaryFiles.outputName(String(localized: "\(base) Comments") + ".md")
        let folder = StudioTemporaryFiles.root.appending(path: "Exported/\(UUID().uuidString)", directoryHint: .isDirectory)
        let url = folder.appending(path: name)
        do {
            try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
            try Data(CommentSummary.markdown(comments.threads, documentName: facts.fileName).utf8).write(to: url)
            return ExportedFile(document: try ExportedDocument(url: url, stagingFolder: folder), contentType: .markdown, name: name, folder: folder)
        } catch {
            try? FileManager.default.removeItem(at: folder)
            throw error
        }
    }

    /// Resolves an open thread or opens a resolved one again, as one undoable change.
    func setResolved(_ resolved: Bool, _ thread: CommentThread) {
        guard canReplyToComments, thread.resolved != resolved else { return }
        let change = CommentStateChange(objectNumber: thread.objectNumber, state: resolved ? "completed" : "none")
        edit(resolved ? String(localized: "Resolve Comment") : String(localized: "Reopen Comment"),
             { try await $0.edit(EditRequest(reviewState: change)) }, then: { session in
                 Task { await session.loadComments() }
             })
    }
}
