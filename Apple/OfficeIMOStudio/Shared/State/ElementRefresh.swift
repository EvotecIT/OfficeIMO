import Foundation

/// How an accessibility element list catches up with what the page shows. Elements a screen reader already holds keep
/// their identity when only their values change, so its focus stays put and it hears the new value; a different set of
/// elements is a new layout.
enum ElementRefresh: Equatable {
    /// Nothing changed.
    case none
    /// The same elements, with new values at these positions.
    case values([Int])
    /// Elements were added, removed or reordered.
    case layout

    /// Compares elements by identity and value, in order.
    static func plan<Identity: Equatable, Value: Equatable>(from old: [(Identity, Value)], to new: [(Identity, Value)]) -> ElementRefresh {
        guard old.count == new.count, zip(old, new).allSatisfy({ $0.0 == $1.0 }) else { return .layout }
        let changed = zip(old, new).enumerated().filter { $0.element.0.1 != $0.element.1.1 }.map(\.offset)
        return changed.isEmpty ? .none : .values(changed)
    }
}
