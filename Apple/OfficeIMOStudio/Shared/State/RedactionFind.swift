import Foundation

struct RedactionFind: Encodable, Sendable, Equatable {
    var text: String
    var regex = false
    var matchCase = false
    /// Removes the whole block of text around each match, for matches too close to their neighbours to remove alone.
    var wholeBlocks = false

    var normalized: Self {
        Self(text: text.trimmingCharacters(in: .whitespaces), regex: regex, matchCase: matchCase, wholeBlocks: wholeBlocks)
    }
}

/// Search matches belong to one normalized request and one document revision.
struct RedactionSearchBinding: Sendable, Equatable {
    let find: RedactionFind
    let revision: Int64

    init(find: RedactionFind, revision: Int64) {
        self.find = find.normalized
        self.revision = revision
    }

    func matches(_ current: RedactionFind, revision: Int64) -> Bool {
        self.revision == revision && find == current.normalized
    }
}
