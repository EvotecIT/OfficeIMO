import XCTest
import PDFKit
import UniformTypeIdentifiers
@testable import StudioApple

final class DocumentTests: XCTestCase {
    private func sample() throws -> Data {
        try Data(contentsOf: XCTUnwrap(Bundle.main.url(forResource: "Welcome", withExtension: "pdf")))
    }

    func testOfficeIMOAnnotationIsReadableByPDFKit() throws {
        let original = try sample()
        let text = "Review — café · Zażółć · 東京"
        let edited = try StudioEngine.addNote(to: original, page: 1, text: text)
        let pdf = try XCTUnwrap(PDFDocument(data: edited))
        XCTAssertEqual(try StudioEngine.inspect(edited).pages, pdf.pageCount)
        XCTAssertTrue(try XCTUnwrap(pdf.page(at: 0)).annotations.contains { $0.contents == text })
        XCTAssertEqual(PDFDocument(data: original)?.page(at: 0)?.annotations.count, 0)
    }

    func testAnnotationStaysInsideRotatedCropBoxOnTheSelectedPage() throws {
        let source = try XCTUnwrap(PDFDocument(data: sample()))
        let second = try XCTUnwrap(source.page(at: 1))
        let crop = CGRect(x: 80, y: 100, width: 400, height: 500)
        second.setBounds(crop, for: .cropBox)
        second.rotation = 90
        let edited = try StudioEngine.addNote(to: XCTUnwrap(source.dataRepresentation()), page: 2, text: "Second page")
        let reopened = try XCTUnwrap(PDFDocument(data: edited))
        let note = try XCTUnwrap(reopened.page(at: 1)?.annotations.first { $0.type == "Text" })
        XCTAssertTrue(crop.contains(note.bounds))
        XCTAssertEqual(reopened.page(at: 0)?.annotations.count, 0)
    }

    func testInvalidInputsReturnErrorsAcrossNativeBoundary() throws {
        XCTAssertThrowsError(try StudioEngine.inspect(Data("not a PDF".utf8)))
        let original = try sample()
        XCTAssertThrowsError(try StudioEngine.addNote(to: original, page: 0, text: "Review"))
        XCTAssertThrowsError(try StudioEngine.addNote(to: original, page: 3, text: "Review"))
        XCTAssertThrowsError(try StudioEngine.addNote(to: original, page: 1, text: "  "))
        XCTAssertThrowsError(try StudioEngine.addNote(to: original, page: 1, text: String(repeating: "a", count: 4001)))
    }

    @MainActor func testUndoHistoryReleasesItsManager() throws {
        let original = try sample()
        let document = try StudioDocument(data: original)
        let edited = try StudioEngine.addNote(to: original, page: 1, text: "Review")
        weak var releasedManager: UndoManager?
        try autoreleasepool {
            let undo = UndoManager()
            releasedManager = undo
            undo.groupsByEvent = false
            undo.beginUndoGrouping()
            try document.replace(with: edited, undoManager: undo)
            undo.endUndoGrouping()
        }
        XCTAssertNil(releasedManager, "Closing the workspace must release its undo history")
    }

    @MainActor func testUndoRedoAndSnapshotPreserveExactDocumentBytes() throws {
        let original = try sample()
        let document = try StudioDocument(data: original)
        let edited = try StudioEngine.addNote(to: original, page: 1, text: "Saved — review")
        let undo = UndoManager()
        undo.groupsByEvent = false
        undo.beginUndoGrouping()
        try document.replace(with: edited, undoManager: undo)
        undo.endUndoGrouping()
        XCTAssertEqual(try document.snapshot(contentType: .pdf), edited)
        XCTAssertFalse(document.pdf.allowsCommenting)
        XCTAssertFalse(document.pdf.allowsFormFieldEntry)
        undo.undo()
        XCTAssertEqual(document.data, original)
        undo.redo()
        XCTAssertEqual(document.data, edited)
        let snapshot = try document.snapshot(contentType: .pdf)
        let reopened = try StudioDocument(data: snapshot)
        XCTAssertEqual(reopened.data, edited)
        XCTAssertThrowsError(try document.replace(with: Data("invalid".utf8), undoManager: undo))
        XCTAssertEqual(document.data, edited)
    }
}
