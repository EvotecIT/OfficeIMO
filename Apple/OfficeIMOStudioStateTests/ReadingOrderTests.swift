import CoreGraphics
import Foundation
import XCTest
@testable import OfficeIMOStudioState

/// The order VoiceOver reads a page's lines, fields and comments in.
final class ReadingOrderTests: XCTestCase {
    func testLinesStartAtEachLineBreak() {
        XCTAssertEqual(ReadingOrder.lines(breaks: []), [])
        XCTAssertEqual(ReadingOrder.lines(breaks: [0]), [0...0])
        // A gap between words (1) stays on the line; 2 starts another, even on the first element.
        XCTAssertEqual(ReadingOrder.lines(breaks: [2, 0, 1, 0, 2, 0, 2]), [0...3, 4...5, 6...6])
    }

    func testFieldsAndCommentsAreReadBeforeTheFirstLineBelowThem() {
        let lines = [CGRect(x: 40, y: 40, width: 300, height: 12), CGRect(x: 40, y: 100, width: 300, height: 12),
                     CGRect(x: 40, y: 160, width: 300, height: 12)]
        let others = [CGRect(x: 40, y: 200, width: 100, height: 20), CGRect(x: 200, y: 70, width: 100, height: 20),
                      CGRect(x: 40, y: 70, width: 100, height: 20), CGRect(x: 10, y: 10, width: 20, height: 20)]
        XCTAssertEqual(ReadingOrder.merge(lines: lines, others: others),
                       [.other(3), .line(0), .other(2), .other(1), .line(1), .line(2), .other(0)])
    }

    func testAPageWithoutTextReadsItsOtherElementsInPlaceOrder() {
        XCTAssertEqual(ReadingOrder.merge(lines: [], others: [CGRect(x: 0, y: 50, width: 1, height: 1), CGRect(x: 0, y: 5, width: 1, height: 1)]),
                       [.other(1), .other(0)])
        XCTAssertEqual(ReadingOrder.merge(lines: [CGRect(x: 0, y: 0, width: 1, height: 1)], others: []), [.line(0)])
    }
}
