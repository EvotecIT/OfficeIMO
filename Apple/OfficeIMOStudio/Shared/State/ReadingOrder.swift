import CoreGraphics
import Foundation

/// The order VoiceOver reads a page in: its lines of text as the engine extracts them, with form fields and comments
/// placed before the first line that starts below them, so each is read where it sits on the page.
enum ReadingOrder {
    /// One item VoiceOver reads: a line of text, by its index among the lines, or another element, by its index.
    enum Item: Equatable {
        case line(Int)
        case other(Int)
    }

    /// The ranges of text elements that make lines, given how each element follows the one before it: 2 starts
    /// another line. The first element always starts one.
    static func lines(breaks: [Int]) -> [ClosedRange<Int>] {
        guard !breaks.isEmpty else { return [] }
        var ranges: [ClosedRange<Int>] = []
        var start = 0
        for index in 1..<breaks.count where breaks[index] == 2 {
            ranges.append(start...(index - 1))
            start = index
        }
        ranges.append(start...(breaks.count - 1))
        return ranges
    }

    /// Lines in reading order with the other elements, ordered top to bottom and left to right among themselves, each
    /// placed before the first line that starts below its top.
    static func merge(lines: [CGRect], others: [CGRect]) -> [Item] {
        let sorted = others.indices.sorted { a, b in
            let first = others[a], second = others[b]
            return abs(first.minY - second.minY) > 1 ? first.minY < second.minY : first.minX < second.minX
        }
        var result: [Item] = []
        var next = 0
        for (index, line) in lines.enumerated() {
            while next < sorted.count, others[sorted[next]].minY < line.minY {
                result.append(.other(sorted[next]))
                next += 1
            }
            result.append(.line(index))
        }
        result += sorted[next...].map { .other($0) }
        return result
    }
}
