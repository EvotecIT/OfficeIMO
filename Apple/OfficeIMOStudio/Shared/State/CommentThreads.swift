import Foundation

/// The document's comment threads from officeimo_document_comments_json: each comment with its replies, nested by
/// depth, and its review state. A thread's id stays the same across revisions, so the selection and a draft reply
/// survive a reply or a change of state that rewrites the page.
struct CommentList: Decodable, Sendable, Equatable {
    let revision: Int64
    let canReply: Bool
    let threads: [CommentThread]
}

struct CommentThread: Decodable, Sendable, Identifiable, Hashable {
    /// Where the comment is, in points of the page as displayed from its top-left corner.
    struct Box: Decodable, Sendable, Hashable {
        let x: Double
        let y: Double
        let width: Double
        let height: Double
    }

    struct Entry: Decodable, Sendable, Hashable {
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

/// Replies being written, one per thread, so moving to another comment and back keeps what was typed. A draft whose
/// thread is no longer in the document, because the comment was removed or an edit replaced it, is kept with what it
/// replied to, so it can be restored to another thread or copied, as the desktop's retained drafts are.
struct CommentDrafts: Sendable, Equatable {
    struct Draft: Sendable, Equatable, Identifiable {
        let thread: String
        /// What the draft replies to: the page and author of the comment.
        let label: String
        var text: String
        var id: String { thread }
    }

    private var drafts: [String: Draft] = [:]

    subscript(thread: CommentThread) -> String {
        get { drafts[thread.id]?.text ?? "" }
        set { drafts[thread.id] = newValue.isEmpty ? nil : Draft(thread: thread.id, label: Self.label(thread), text: newValue) }
    }

    func text(for thread: String) -> String { drafts[thread]?.text ?? "" }

    /// Drops the draft once its reply is in the document, unless it changed while the reply was added.
    mutating func sent(_ text: String, to thread: String) {
        if drafts[thread]?.text == text { drafts[thread] = nil }
    }

    /// Drafts whose thread the document no longer has, in the order they were started.
    func unassigned(keeping threads: [CommentThread]) -> [Draft] {
        let ids = Set(threads.map(\.id))
        return drafts.values.filter { !ids.contains($0.thread) }.sorted { $0.label < $1.label }
    }

    /// Moves a retained draft to a thread that has no draft of its own; false when it has one.
    @discardableResult
    mutating func restore(_ draft: String, to thread: CommentThread) -> Bool {
        guard let retained = drafts[draft], self[thread].isEmpty else { return false }
        drafts[draft] = nil
        self[thread] = retained.text
        return true
    }

    mutating func discard(_ draft: String) { drafts[draft] = nil }

    private static func label(_ thread: CommentThread) -> String {
        let author = thread.comment?.author.flatMap { $0.isEmpty ? nil : $0 } ?? String(localized: "Unknown Author")
        guard let page = thread.page else { return author }
        return String(localized: "Page \(page), \(author)")
    }
}

/// The comment threads as a Markdown summary, grouped by page with each thread's state and replies, as the desktop's
/// Export Summary writes it.
enum CommentSummary {
    static func markdown(_ threads: [CommentThread], documentName: String) -> String {
        var text = "# " + String(localized: "Comments on \(escape(documentName))") + "\n\n"
        let open = threads.filter { !$0.resolved }.count
        text += String(localized: "\(threads.count) threads, \(open) open.") + "\n\n"
        let pages = Dictionary(grouping: threads) { $0.page ?? 0 }
        for page in pages.keys.sorted() {
            text += "## " + (page == 0 ? String(localized: "No Page") : String(localized: "Page \(page)")) + "\n\n"
            for thread in pages[page] ?? [] {
                for (index, entry) in thread.entries.enumerated() {
                    let author = escape(entry.author.flatMap { $0.isEmpty ? nil : $0 } ?? String(localized: "Unknown Author"))
                    let contents = entry.contents.flatMap { $0.isEmpty ? nil : $0 }.map(escape) ?? String(localized: "(no text)")
                    if index == 0 {
                        text += "- **\(author)** (\(escape(stateName(thread.state)))): \(contents)\n"
                    } else {
                        text += "  - \(author): \(contents)\n"
                    }
                }
            }
            text += "\n"
        }
        return text
    }

    /// What people call a review state.
    static func stateName(_ state: String) -> String {
        switch state {
        case "completed": String(localized: "Resolved")
        case "accepted": String(localized: "Accepted")
        case "rejected": String(localized: "Rejected")
        case "cancelled": String(localized: "Cancelled")
        case "marked": String(localized: "Marked")
        case "unmarked": String(localized: "Unmarked")
        default: String(localized: "Unresolved")
        }
    }

    /// Text that reads the same in Markdown: whitespace as spaces, HTML characters as entities and Markdown punctuation
    /// escaped.
    static func escape(_ value: String) -> String {
        var result = ""
        for character in value {
            if character.isWhitespace { result.append(" ") }
            else if character == "&" { result += "&amp;" }
            else if character == "<" { result += "&lt;" }
            else if character == ">" { result += "&gt;" }
            else {
                if "\\`*_{}[]()#+-.!|~".contains(character) { result.append("\\") }
                result.append(character)
            }
        }
        return result
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
