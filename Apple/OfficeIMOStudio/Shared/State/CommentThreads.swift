import Foundation

/// The document's comment threads from officeimo_document_comments_json: each comment with its replies, nested by
/// depth, and its review state. A thread's id stays the same across revisions, so the selection and a draft reply
/// survive a reply or a change of state that rewrites the page.
struct CommentList: Decodable, Sendable, Equatable {
    let revision: Int64
    let canReply: Bool
    let threads: [CommentThread]
}

struct CommentThread: Decodable, Sendable, Identifiable, Equatable {
    /// Where the comment is, in points of the page as displayed from its top-left corner.
    struct Box: Decodable, Sendable, Equatable {
        let x: Double
        let y: Double
        let width: Double
        let height: Double
    }

    struct Entry: Decodable, Sendable, Equatable {
        let objectNumber: Int?
        let author: String?
        let contents: String?
        let modified: String?
        /// 0 for the comment, 1 for a reply to it, 2 for a reply to that reply.
        let depth: Int
    }

    let id: String
    let objectNumber: Int
    let page: Int?
    let subtype: String
    let bounds: Box?
    let state: String
    let resolved: Bool
    /// A reply whose comment is gone.
    let orphaned: Bool
    let entries: [Entry]

    var comment: Entry? { entries.first }
    var replies: ArraySlice<Entry> { entries.dropFirst() }

    /// Whether any entry was written by someone whose name contains the text, ignoring case and accents.
    func hasAuthor(matching text: String) -> Bool {
        let text = text.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !text.isEmpty else { return true }
        return entries.contains { $0.author?.range(of: text, options: [.caseInsensitive, .diacriticInsensitive]) != nil }
    }
}

/// Which comments the list shows, as the desktop's Comments filter does.
enum CommentFilter: String, CaseIterable, Identifiable, Sendable {
    case all
    case open
    case resolved

    var id: Self { self }

    func includes(_ thread: CommentThread) -> Bool {
        switch self {
        case .all: true
        case .open: !thread.resolved
        case .resolved: thread.resolved
        }
    }

    /// The threads the filter and the author text keep, in document order.
    func apply(to threads: [CommentThread], author: String = "") -> [CommentThread] {
        threads.filter { includes($0) && $0.hasAuthor(matching: author) }
    }
}

enum CommentNavigation {
    /// The next open thread after the one given, wrapping around to the first; nil when every thread is resolved.
    static func nextOpen(after id: String?, in threads: [CommentThread]) -> CommentThread? {
        guard !threads.isEmpty else { return nil }
        let start = id.flatMap { id in threads.firstIndex { $0.id == id } } ?? -1
        for offset in 1...threads.count {
            let thread = threads[(start + offset + threads.count) % threads.count]
            if !thread.resolved { return thread }
        }
        return nil
    }
}

/// Replies being written, one per thread, so moving to another comment and back keeps what was typed.
struct CommentDrafts: Sendable, Equatable {
    private var texts: [String: String] = [:]

    subscript(thread: String) -> String {
        get { texts[thread] ?? "" }
        set { texts[thread] = newValue.isEmpty ? nil : newValue }
    }

    /// Drops the draft once its reply is in the document, unless it changed while the reply was added.
    mutating func sent(_ text: String, to thread: String) {
        if texts[thread] == text { texts[thread] = nil }
    }
}

/// The name comments, replies and markup carry as their author. People set it once; until then the Mac offers their
/// account's full name, and iPad and iPhone, which have no such name, leave it empty.
enum CommentAuthor {
    static let defaultsKey = "CommentAuthor"

    static func resolve(stored: String?, system: String) -> String {
        let stored = stored?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
        return stored.isEmpty ? system.trimmingCharacters(in: .whitespacesAndNewlines) : stored
    }
}
