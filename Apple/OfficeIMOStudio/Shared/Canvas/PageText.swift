import CoreGraphics
import Foundation

/// A page's text from officeimo_document_page_text_json: its elements in extraction order, each one's box in the
/// displayed page's points and how it follows the one before. Selection runs along the elements, as reading does; the
/// engine decides where words and lines break, as it does for every OfficeIMO reader.
struct PageText: Decodable, Sendable {
    let revision: Int64
    let page: Int
    let texts: [String]
    let boxes: [Double]
    let breaks: [Break]

    /// How an element follows the one before it.
    enum Break: Int, Decodable, Sendable {
        /// Its glyphs touch the previous element's.
        case none = 0
        /// It is further along the same line after a gap, as between a scan's recognized words.
        case space = 1
        /// It starts another line.
        case line = 2
    }

    var count: Int { texts.count }

    func box(_ index: Int) -> CGRect {
        CGRect(x: boxes[index * 4], y: boxes[index * 4 + 1], width: boxes[index * 4 + 2], height: boxes[index * 4 + 3])
    }

    /// The element under a point, or the nearest one within reach: on the line the point is beside, then any.
    func index(nearest point: CGPoint, reach: CGFloat = 24) -> Int? {
        guard count > 0 else { return nil }
        var best: (index: Int, distance: CGFloat)?
        for index in 0..<count where !texts[index].isEmpty {
            let rect = box(index)
            if rect.contains(point) { return index }
            let dx = max(rect.minX - point.x, 0, point.x - rect.maxX)
            let dy = max(rect.minY - point.y, 0, point.y - rect.maxY)
            // Vertical distance weighs more, so a point beside a line picks that line.
            let distance = hypot(dx, dy * 3)
            if distance < (best?.distance ?? .infinity) { best = (index, distance) }
        }
        guard let best, best.distance <= reach * 3 else { return nil }
        return best.index
    }

    /// The word around an element: the touching elements around it, up to whitespace or a gap.
    func word(at index: Int) -> ClosedRange<Int> {
        guard isWordElement(index) else { return index...index }
        var first = index
        var last = index
        while first > 0, isWordElement(first - 1), breaks[first] == .none { first -= 1 }
        while last + 1 < count, isWordElement(last + 1), breaks[last + 1] == .none { last += 1 }
        return first...last
    }

    /// One rectangle per line of a range, for the selection and for text markup.
    func lineRects(_ range: ClosedRange<Int>) -> [CGRect] {
        var rects: [CGRect] = []
        var current: CGRect?
        var newLine = false
        for index in range where index < count {
            if index > range.lowerBound, breaks[index] == .line { newLine = true }
            let rect = box(index)
            // Spaces join the words around them but never extend a line's mark past its last word.
            guard rect.width > 0, rect.height > 0, !texts[index].allSatisfy(\.isWhitespace) else { continue }
            if let line = current, !newLine {
                current = line.union(rect)
            } else {
                if let line = current { rects.append(line) }
                current = rect
            }
            newLine = false
        }
        if let line = current { rects.append(line) }
        return rects
    }

    /// The text of a range, with a space across a gap between words and a line break where the selection moves to
    /// another line, as OfficeIMO's readers copy it.
    func string(_ range: ClosedRange<Int>) -> String {
        var result = ""
        for index in range where index < count {
            let text = texts[index]
            if index > range.lowerBound, let last = result.last {
                switch breaks[index] {
                case .space:
                    if !last.isWhitespace, text.first?.isWhitespace != true { result.append(" ") }
                case .line:
                    if !last.isNewline, text.first?.isNewline != true { result.append("\n") }
                case .none: break
                }
            }
            result += text
        }
        return result.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    private func isWordElement(_ index: Int) -> Bool {
        let text = texts[index]
        return !text.isEmpty && !text.contains(where: \.isWhitespace) && !text.allSatisfy(\.isPunctuation)
    }
}

/// Text selected on one page: a range of its elements.
struct TextSelection: Equatable {
    let page: Int
    let revision: Int64
    /// Where the selection started, which stays put while it is extended.
    let anchor: ClosedRange<Int>
    var range: ClosedRange<Int>

    /// Extends from the anchor to an element, either way.
    func extended(to index: Int) -> TextSelection {
        var selection = self
        selection.range = index < anchor.lowerBound ? index...anchor.upperBound : anchor.lowerBound...max(anchor.upperBound, index)
        return selection
    }
}
