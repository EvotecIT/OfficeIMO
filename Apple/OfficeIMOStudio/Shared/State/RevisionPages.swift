import Foundation

/// Values read for a few pages of a document, such as their form fields, each kept with the revision it was read from.
/// Only values of the revision shown count, so an edit makes every page read again unless a value is carried over. The
/// session keeps it observed, so a view that shows a page's value, like the selected form field, updates as it arrives.
struct RevisionPages<Value> {
    private var values: [Int: (revision: Int64, value: Value)] = [:]
    /// How many pages are kept; storing another drops the page farthest from it.
    let capacity: Int

    init(capacity: Int = 8) {
        self.capacity = capacity
    }

    func value(page: Int, revision: Int64) -> Value? {
        guard let entry = values[page], entry.revision == revision else { return nil }
        return entry.value
    }

    /// The pages read for a revision.
    func values(revision: Int64) -> [Int: Value] {
        values.compactMapValues { $0.revision == revision ? $0.value : nil }
    }

    mutating func store(_ value: Value, page: Int, revision: Int64) {
        values[page] = (revision, value)
        while values.count > capacity, let far = values.keys.max(by: { abs($0 - page) < abs($1 - page) }) {
            values[far] = nil
        }
    }

    /// Moves the values of one revision to the next, changed as an edit changed them, so they stay in use at once.
    mutating func carry(from old: Int64, to new: Int64, _ change: (Value) -> Value) {
        for (page, entry) in values where entry.revision == old {
            values[page] = (new, change(entry.value))
        }
    }
}
