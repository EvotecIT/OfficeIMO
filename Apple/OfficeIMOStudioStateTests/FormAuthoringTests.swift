import CoreGraphics
import Foundation
import XCTest
@testable import OfficeIMOStudioState

/// Naming, option lists and boxes for fields added with Edit Form.
final class FormAuthoringTests: XCTestCase {
    private let page = CGSize(width: 400, height: 300)

    func testNewFieldsTakeTheFirstUnusedNumberedName() {
        XCTAssertEqual(FormAuthoring.nextName(existing: []), "Field1")
        XCTAssertEqual(FormAuthoring.nextName(existing: ["Field1", "Field3", "Name"]), "Field2")
    }

    func testOptionsComeOneALineTrimmedWithoutBlanksOrRepeats() {
        XCTAssertEqual(FormAuthoring.options(from: " Poland \n\nGermany\r\nPoland\n  \nFrance"), ["Poland", "Germany", "France"])
        XCTAssertEqual(FormAuthoring.options(from: "   "), [])
    }

    func testADragDrawsItsBoxAndAClickGivesTheStandardSizeCentredOnIt() {
        XCTAssertEqual(FormAuthoring.box(from: CGPoint(x: 200, y: 80), to: CGPoint(x: 40, y: 50), kind: .text, page: page),
                       CGRect(x: 40, y: 50, width: 160, height: 30))
        XCTAssertEqual(FormAuthoring.box(from: CGPoint(x: 100, y: 100), to: CGPoint(x: 101, y: 101), kind: .signature, page: page),
                       CGRect(x: 0, y: 76, width: 200, height: 48))
        XCTAssertEqual(FormAuthoring.box(from: CGPoint(x: 100, y: 100), to: CGPoint(x: 100, y: 100), kind: .checkbox, page: page),
                       CGRect(x: 92, y: 92, width: 16, height: 16))
    }

    func testCheckBoxesStaySquareAndBoxesStayOnThePage() {
        let square = FormAuthoring.box(from: CGPoint(x: 10, y: 10), to: CGPoint(x: 30, y: 22), kind: .checkbox, page: page)
        XCTAssertEqual(square.width, square.height)
        let clamped = FormAuthoring.clamp(CGRect(x: 380, y: -20, width: 100, height: 1), to: page)
        XCTAssertEqual(clamped, CGRect(x: 300, y: 0, width: 100, height: FormAuthoring.minimumSide))
    }

    func testNamesMustBeGivenAndUnusedExceptByTheFieldBeingRenamed() {
        XCTAssertNotNil(FormAuthoring.nameProblem("  ", existing: []))
        XCTAssertNotNil(FormAuthoring.nameProblem("Name", existing: ["Name"]))
        XCTAssertNil(FormAuthoring.nameProblem("Name", existing: ["Name"], renaming: "Name"))
        XCTAssertNil(FormAuthoring.nameProblem(" Email ", existing: ["Name"]))
    }
}
