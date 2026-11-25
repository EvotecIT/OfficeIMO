import Foundation
import XCTest
@testable import OfficeIMOStudioState

/// Page values, such as form fields, are used only for the revision they were read from.
final class RevisionPagesTests: XCTestCase {
    func testAValueCountsOnlyForItsRevisionSoAnEditShowsTheFieldsReadAfterIt() {
        var pages = RevisionPages<[String]>()
        pages.store(["Name"], page: 1, revision: 4)
        XCTAssertEqual(pages.value(page: 1, revision: 4), ["Name"])
        // After adding a field the revision moves on; the old fields no longer count until the new ones are stored.
        XCTAssertNil(pages.value(page: 1, revision: 5))
        pages.store(["Name", "Field1"], page: 1, revision: 5)
        XCTAssertEqual(pages.value(page: 1, revision: 5), ["Name", "Field1"])
        XCTAssertNil(pages.value(page: 2, revision: 5))
    }

    func testValuesOfOneRevisionAreListedAndCarriedToTheNext() {
        var pages = RevisionPages<String>()
        pages.store("one", page: 1, revision: 1)
        pages.store("two", page: 2, revision: 1)
        pages.store("old", page: 3, revision: 0)
        XCTAssertEqual(pages.values(revision: 1), [1: "one", 2: "two"])
        pages.carry(from: 1, to: 2) { $0.uppercased() }
        XCTAssertEqual(pages.values(revision: 2), [1: "ONE", 2: "TWO"])
        XCTAssertEqual(pages.values(revision: 1), [:])
    }

    func testStoringBeyondCapacityDropsThePagesFarthestAway() {
        var pages = RevisionPages<Int>(capacity: 3)
        for page in [1, 2, 3, 10] { pages.store(page, page: page, revision: 1) }
        XCTAssertEqual(Set(pages.values(revision: 1).keys), [2, 3, 10])
        pages.store(4, page: 4, revision: 1)
        XCTAssertEqual(Set(pages.values(revision: 1).keys), [2, 3, 4])
    }
}
