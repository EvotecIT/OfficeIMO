import Foundation
import XCTest
@testable import OfficeIMOStudioState

/// Comment threads as the engine reports them, filtered and stepped through as the Comments inspector does.
final class CommentThreadsTests: XCTestCase {
    private static let json = """
    { "abiVersion": 3, "revision": 7, "canReply": true, "threads": [
      { "id": "a", "objectNumber": 12, "page": 1, "subtype": "Text", "bounds": { "x": 40, "y": 40, "width": 20, "height": 20 },
        "state": "none", "resolved": false, "orphaned": false, "entries": [
          { "objectNumber": 12, "author": "Alex Morgan", "contents": "Check the totals", "modified": null, "depth": 0 },
          { "objectNumber": 15, "author": "Zoë Lee", "contents": "Fixed", "modified": null, "depth": 1 } ] },
      { "id": "b", "objectNumber": 20, "page": 2, "subtype": "Highlight", "bounds": null,
        "state": "completed", "resolved": true, "orphaned": false, "entries": [
          { "objectNumber": 20, "author": "Sam", "contents": null, "modified": null, "depth": 0 } ] },
      { "id": "c", "objectNumber": 31, "page": 3, "subtype": "Text", "bounds": null,
        "state": "none", "resolved": false, "orphaned": true, "entries": [
          { "objectNumber": 31, "author": null, "contents": "Orphan", "modified": null, "depth": 1 } ] } ] }
    """

    private func threads() throws -> [CommentThread] {
        try JSONDecoder().decode(CommentList.self, from: Data(Self.json.utf8)).threads
    }

    func testThreadsDecodeWithTheirRepliesInOrder() throws {
        let list = try JSONDecoder().decode(CommentList.self, from: Data(Self.json.utf8))
        XCTAssertEqual(list.revision, 7)
        XCTAssertTrue(list.canReply)
        let first = try XCTUnwrap(list.threads.first)
        XCTAssertEqual(first.comment?.contents, "Check the totals")
        XCTAssertEqual(first.replies.map(\.author), ["Zoë Lee"])
        XCTAssertEqual(first.bounds?.width, 20)
        XCTAssertNil(list.threads[1].bounds)
        XCTAssertTrue(list.threads[2].orphaned)
    }

    func testFiltersKeepOpenOrResolvedThreadsAndMatchAnyAuthorIgnoringCaseAndAccents() throws {
        let threads = try threads()
        XCTAssertEqual(CommentFilter.all.apply(to: threads).map(\.id), ["a", "b", "c"])
        XCTAssertEqual(CommentFilter.open.apply(to: threads).map(\.id), ["a", "c"])
        XCTAssertEqual(CommentFilter.resolved.apply(to: threads).map(\.id), ["b"])
        // A reply's author finds the thread; an entry without an author never matches a name.
        XCTAssertEqual(CommentFilter.all.apply(to: threads, author: "zoe").map(\.id), ["a"])
        XCTAssertEqual(CommentFilter.all.apply(to: threads, author: "  ").map(\.id), ["a", "b", "c"])
        XCTAssertEqual(CommentFilter.resolved.apply(to: threads, author: "alex").map(\.id), [])
    }

    func testNextOpenThreadWrapsAroundAndSkipsResolvedOnes() throws {
        let threads = try threads()
        XCTAssertEqual(CommentNavigation.nextOpen(after: nil, in: threads)?.id, "a")
        XCTAssertEqual(CommentNavigation.nextOpen(after: "a", in: threads)?.id, "c")
        XCTAssertEqual(CommentNavigation.nextOpen(after: "c", in: threads)?.id, "a")
        XCTAssertEqual(CommentNavigation.nextOpen(after: "missing", in: threads)?.id, "a")
        XCTAssertNil(CommentNavigation.nextOpen(after: nil, in: threads.filter(\.resolved)))
        XCTAssertNil(CommentNavigation.nextOpen(after: nil, in: []))
    }

    func testDraftsStayWithTheirThreadUntilTheReplyIsSent() {
        var drafts = CommentDrafts()
        drafts["a"] = "Looks good"
        drafts["b"] = "Not yet"
        XCTAssertEqual(drafts["a"], "Looks good")
        drafts.sent("Looks good", to: "a")
        XCTAssertEqual(drafts["a"], "")
        // A draft edited while its reply was being added is kept.
        drafts.sent("Not", to: "b")
        XCTAssertEqual(drafts["b"], "Not yet")
        drafts["b"] = ""
        XCTAssertEqual(drafts, CommentDrafts())
    }

    func testTheChosenAuthorNameWinsOverTheSystemName() {
        XCTAssertEqual(CommentAuthor.resolve(stored: nil, system: "Alex Morgan"), "Alex Morgan")
        XCTAssertEqual(CommentAuthor.resolve(stored: "  ", system: " Alex Morgan "), "Alex Morgan")
        XCTAssertEqual(CommentAuthor.resolve(stored: " Sam ", system: "Alex Morgan"), "Sam")
        XCTAssertEqual(CommentAuthor.resolve(stored: nil, system: ""), "")
    }
}
