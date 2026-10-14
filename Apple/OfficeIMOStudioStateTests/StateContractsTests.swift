import Foundation
import XCTest
@testable import OfficeIMOStudioState

// Test instances have no mutable stored state; each async test owns its actor-isolated fixtures.
final class StateContractsTests: XCTestCase, @unchecked Sendable {
    func testPageImagesRequireContentCopyPermission() {
        for format in [ExportFormat.word, .excel, .powerpoint, .html, .markdown, .json, .pageImages, .images] {
            XCTAssertFalse(format.isAllowed(canExtractText: true, canExtractContent: false))
            XCTAssertTrue(format.isAllowed(canExtractText: false, canExtractContent: true))
            XCTAssertFalse(format.isAllowed(canExtractText: false, canExtractContent: false))
        }
    }

    func testAccessibilityTextRemainsAvailableWithoutImageCopy() {
        XCTAssertTrue(ExportFormat.text.isAllowed(canExtractText: true, canExtractContent: false))
        XCTAssertFalse(ExportFormat.text.isAllowed(canExtractText: false, canExtractContent: false))
    }

    func testEquivalentWhitespaceKeepsRedactionBinding() {
        let binding = RedactionSearchBinding(find: RedactionFind(text: "  confidential  "), revision: 7)
        XCTAssertEqual(binding.find.text, "confidential")
        XCTAssertTrue(binding.matches(RedactionFind(text: "confidential"), revision: 7))
        XCTAssertTrue(binding.matches(RedactionFind(text: " confidential "), revision: 7))
    }

    func testChangedQueryOptionsOrDocumentInvalidateRedactionBinding() {
        let binding = RedactionSearchBinding(find: RedactionFind(text: "secret"), revision: 7)
        XCTAssertFalse(binding.matches(RedactionFind(text: "other"), revision: 7))
        XCTAssertFalse(binding.matches(RedactionFind(text: "secret", regex: true), revision: 7))
        XCTAssertFalse(binding.matches(RedactionFind(text: "secret", matchCase: true), revision: 7))
        XCTAssertFalse(binding.matches(RedactionFind(text: "secret"), revision: 8))
    }

    @MainActor
    func testRevisionChangeDuringReadRetriesCurrentRevision() async throws {
        let revision = RevisionFixture()
        let counter = Counter()
        let read = RevisionRead<Int>()
        let value = try await read.read(source: revision, revision: { revision.value }) {
            let call = await counter.next()
            if call == 1 { await revision.advance() }
            return call
        }
        XCTAssertEqual(value, 2)
        let calls = await counter.value
        XCTAssertEqual(calls, 2)
    }

    @MainActor
    func testStaleFailureDoesNotHideCurrentRevision() async throws {
        let revision = RevisionFixture()
        let counter = Counter()
        let read = RevisionRead<Int>()
        let value = try await read.read(source: revision, revision: { revision.value }) {
            let call = await counter.next()
            if call == 1 {
                await revision.advance()
                throw ReadFailure.failed
            }
            return call
        }
        XCTAssertEqual(value, 2)
    }

    @MainActor
    func testSuccessfulRevisionIsReadOnceAndNewSourceIsIndependent() async throws {
        let revision = RevisionFixture()
        let other = RevisionFixture()
        let counter = Counter()
        let read = RevisionRead<Int>()
        let first = try await read.read(source: revision, revision: { revision.value }) { await counter.next() }
        let cached = try await read.read(source: revision, revision: { revision.value }) { await counter.next() }
        let replaced = try await read.read(source: other, revision: { other.value }) { await counter.next() }
        XCTAssertEqual(first, 1)
        XCTAssertEqual(cached, 1)
        XCTAssertEqual(replaced, 2)
    }

    @MainActor
    func testFailedRevisionCanBeRetried() async throws {
        let revision = RevisionFixture()
        let read = RevisionRead<Int>()
        do {
            _ = try await read.read(source: revision, revision: { revision.value }) { throw ReadFailure.failed }
            XCTFail("The failed read must throw")
        } catch ReadFailure.failed { }
        let value = try await read.read(source: revision, revision: { revision.value }) { 42 }
        XCTAssertEqual(value, 42)
    }

    @MainActor
    func testNewRevisionStartsBeforeOldReadFinishesAndOldCompletionPreservesIt() async throws {
        let firstStarted = expectation(description: "Old revision started")
        let secondStarted = expectation(description: "Current revision started")
        let retried = expectation(description: "Old waiter sees current revision")
        let revision = RevisionFixture()
        let probe = ReadProbe { number in
            if number == 1 { firstStarted.fulfill() }
            if number == 2 { secondStarted.fulfill() }
        }
        let read = RevisionRead<Int>()
        let first = Task {
            try await read.read(source: revision, revision: {
                if revision.value == 1, !revision.notified {
                    revision.notified = true
                    retried.fulfill()
                }
                return revision.value
            }) { try await probe.read() }
        }
        await fulfillment(of: [firstStarted], timeout: 10)
        revision.advance()
        let second = Task {
            try await read.read(source: revision, revision: { revision.value }) { try await probe.read() }
        }
        await fulfillment(of: [secondStarted], timeout: 10)
        await probe.complete(1, value: 100)
        await fulfillment(of: [retried], timeout: 10)
        await probe.completeAll(value: 200)
        let current = try await second.value
        let recovered = try await first.value
        let calls = await probe.calls
        XCTAssertEqual(current, 200)
        XCTAssertEqual(recovered, 200)
        XCTAssertEqual(calls, 2)
    }
}

@MainActor
private final class RevisionFixture {
    var value: Int64 = 0
    var notified = false
    func advance() { value += 1 }
}

private actor Counter {
    var value = 0
    func next() -> Int { value += 1; return value }
}

private enum ReadFailure: Error { case failed }

private actor ReadProbe {
    let started: @Sendable (Int) -> Void
    var calls = 0
    private var pending: [Int: CheckedContinuation<Int, Error>] = [:]
    private var completedValue: Int?

    init(started: @escaping @Sendable (Int) -> Void) { self.started = started }

    func read() async throws -> Int {
        calls += 1
        let number = calls
        if let completedValue { return completedValue }
        return try await withCheckedThrowingContinuation { continuation in
            pending[number] = continuation
            started(number)
        }
    }

    func complete(_ number: Int, value: Int) { pending.removeValue(forKey: number)?.resume(returning: value) }

    func completeAll(value: Int) {
        completedValue = value
        let waiting = pending.values
        pending.removeAll()
        for continuation in waiting { continuation.resume(returning: value) }
    }
}
