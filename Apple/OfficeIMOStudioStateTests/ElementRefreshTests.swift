import Foundation
import XCTest
@testable import OfficeIMOStudioState

/// A filled check box keeps its element and reports its new value; text arriving for a page is a new layout.
final class ElementRefreshTests: XCTestCase {
    func testOnlyChangedValuesAreRefreshedWhenTheElementsStayTheSame() {
        let before = [("page-1", "Page 1"), ("field-Agree", "Off"), ("field-Name", "")]
        XCTAssertEqual(ElementRefresh.plan(from: before, to: before), .none)
        XCTAssertEqual(ElementRefresh.plan(from: before, to: [("page-1", "Page 1"), ("field-Agree", "On"), ("field-Name", "")]), .values([1]))
    }

    func testAddedRemovedOrReorderedElementsAreANewLayout() {
        let before = [("page-1", "Page 1"), ("field-Agree", "Off")]
        XCTAssertEqual(ElementRefresh.plan(from: before, to: before + [("text-1-0", "Invoice")]), .layout)
        XCTAssertEqual(ElementRefresh.plan(from: before, to: Array(before.reversed())), .layout)
        XCTAssertEqual(ElementRefresh.plan(from: before, to: [("page-1", "Page 1")]), .layout)
    }
}
