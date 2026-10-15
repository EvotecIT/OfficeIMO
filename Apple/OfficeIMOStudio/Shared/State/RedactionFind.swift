import Foundation

struct RedactionFind: Encodable, Sendable, Equatable {
    var text: String
    var regex = false
    var matchCase = false

    var normalized: Self {
        Self(text: text.trimmingCharacters(in: .whitespaces), regex: regex, matchCase: matchCase)
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
