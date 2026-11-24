import CoreGraphics
import Foundation

/// The fields Edit Form adds, as the desktop Studio's form authoring does: a text field, a check box, a choice of
/// options or a place to sign.
enum NewFormFieldKind: String, CaseIterable, Identifiable, Sendable {
    case text
    case checkbox
    case choice
    case signature

    var id: Self { self }

    /// The size a click or tap without a drag gives the field, in points.
    var standardSize: CGSize {
        switch self {
        case .text: CGSize(width: 180, height: 24)
        case .checkbox: CGSize(width: 16, height: 16)
        case .choice: CGSize(width: 180, height: 24)
        case .signature: CGSize(width: 200, height: 48)
        }
    }
}

enum FormAuthoring {
    /// The smallest box a field can have, in points, which the engine also requires.
    static let minimumSide: CGFloat = 4

    /// The first of Field1, Field2 and so on that the form doesn't use, as the desktop names new fields.
    static func nextName(existing: Set<String>) -> String {
        var number = 1
        while existing.contains("Field\(number)") { number += 1 }
        return "Field\(number)"
    }

    /// A choice field's options from text with one option per line: trimmed, without empty lines or repeats.
    static func options(from text: String) -> [String] {
        var seen = Set<String>()
        return text.split(whereSeparator: \.isNewline)
            .map { $0.trimmingCharacters(in: .whitespaces) }
            .filter { !$0.isEmpty && seen.insert($0).inserted }
    }

    /// The box a drag draws on a page of the given size, or the kind's standard size centred on a click; kept on the
    /// page. A check box stays square.
    static func box(from start: CGPoint, to end: CGPoint, kind: NewFormFieldKind, page: CGSize) -> CGRect {
        var rect = CGRect(x: min(start.x, end.x), y: min(start.y, end.y), width: abs(end.x - start.x), height: abs(end.y - start.y))
        if rect.width < minimumSide && rect.height < minimumSide {
            let size = kind.standardSize
            rect = CGRect(x: start.x - size.width / 2, y: start.y - size.height / 2, width: size.width, height: size.height)
        }
        if kind == .checkbox {
            let side = max(rect.width, rect.height)
            rect.size = CGSize(width: side, height: side)
        }
        return clamp(rect, to: page)
    }

    /// A box kept on its page and at least the minimum size.
    static func clamp(_ rect: CGRect, to page: CGSize) -> CGRect {
        var rect = rect
        rect.size.width = min(max(rect.width, minimumSide), max(page.width, minimumSide))
        rect.size.height = min(max(rect.height, minimumSide), max(page.height, minimumSide))
        rect.origin.x = min(max(rect.minX, 0), max(page.width - rect.width, 0))
        rect.origin.y = min(max(rect.minY, 0), max(page.height - rect.height, 0))
        return rect
    }

    /// Why a name can't be used for a field, or nil when it can: it must not be empty, and only the field being
    /// renamed may already have it.
    static func nameProblem(_ name: String, existing: Set<String>, renaming current: String? = nil) -> String? {
        let name = name.trimmingCharacters(in: .whitespacesAndNewlines)
        if name.isEmpty { return String(localized: "Enter a name for the field.") }
        if name != current && existing.contains(name) { return String(localized: "The form already has a field named “\(name)”.") }
        return nil
    }
}
