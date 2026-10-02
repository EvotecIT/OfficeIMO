import XCTest

final class DocumentJourneyTests: XCTestCase {
    func testCreateAnnotateAndReopen() throws {
        continueAfterFailure = false
        let app = XCUIApplication()
        app.launchArguments = ["-AppleLanguages", "(en)", "-AppleLocale", "en_US"]
        app.launch()
        capture("document-browser", app)
        let create = app.buttons["Create Document"]
        XCTAssertTrue(create.waitForExistence(timeout: 20), app.debugDescription)
        let recents = app.buttons["Recents"].firstMatch
        XCTAssertTrue(recents.waitForExistence(timeout: 10))
        recents.tap()
        create.tap()
        let add = app.buttons["addNote"]
        XCTAssertTrue(add.waitForExistence(timeout: 30), app.debugDescription)
        capture("workspace-portrait", app)
        app.buttons["Next page"].tap()
        XCTAssertTrue(app.staticTexts["2 of 2"].waitForExistence(timeout: 5))
        add.tap()
        let note = app.textViews["noteText"]
        XCTAssertTrue(note.waitForExistence(timeout: 5), app.debugDescription)
        note.tap()
        note.typeText("Reviewed on Apple - saved by OfficeIMO")
        capture("note-entry", app)
        app.buttons["confirmNote"].tap()
        XCTAssertTrue(add.waitForExistence(timeout: 15))
        let save = app.buttons["saveCopy"]
        XCTAssertTrue(save.waitForExistence(timeout: 5))
        XCTAssertTrue(NSPredicate(format: "enabled == true").evaluate(with: save))
        save.tap()
        let saveAction = app.buttons["Save"]
        XCTAssertTrue(saveAction.waitForExistence(timeout: 10), app.debugDescription)
        capture("save-copy", app)
        let filename = "OfficeIMO review " + UUID().uuidString.prefix(8)
        let name = app.textFields["DOCPicker.filenameTextField"]
        XCTAssertTrue(name.exists)
        name.tap()
        let oldName = name.value as? String ?? ""
        name.typeText(String(repeating: XCUIKeyboardKey.delete.rawValue, count: oldName.count) + filename)
        saveAction.tap()
        XCTAssertTrue(saveAction.waitForNonExistence(timeout: 15), app.debugDescription)
        app.buttons["BackButton"].tap()
        XCTAssertTrue(create.waitForExistence(timeout: 10), app.debugDescription)
        app.terminate()
        app.launch()
        let browse = app.buttons["Browse"].firstMatch
        XCTAssertTrue(browse.waitForExistence(timeout: 20), app.debugDescription)
        browse.tap()
        let saved = app.cells[filename + ", pdf"]
        if !saved.waitForExistence(timeout: 3) {
            let local = app.cells.matching(NSPredicate(format: "label BEGINSWITH %@", "On My")).firstMatch
            XCTAssertTrue(local.waitForExistence(timeout: 5), app.debugDescription)
            local.tap()
            let folder = app.cells.matching(NSPredicate(format: "label BEGINSWITH %@", "OfficeIMO Studio Native")).firstMatch
            XCTAssertTrue(folder.waitForExistence(timeout: 5), app.debugDescription)
            folder.tap()
        }
        XCTAssertTrue(saved.waitForExistence(timeout: 10), app.debugDescription)
        saved.tap()
        XCTAssertTrue(app.otherElements["documentCanvas"].waitForExistence(timeout: 15), app.debugDescription)
        if !add.exists {
            app.buttons["OverflowBarButtonItem"].tap()
            let overflowAdd = app.buttons["Add Note"]
            XCTAssertTrue(overflowAdd.waitForExistence(timeout: 5), app.debugDescription)
            overflowAdd.tap()
            XCTAssertTrue(note.waitForExistence(timeout: 5))
            app.buttons["Cancel"].tap()
        }
        app.buttons["Next page"].tap()
        XCTAssertTrue(app.staticTexts["2 of 2"].waitForExistence(timeout: 5))
        if !app.buttons["Page 2"].exists { app.buttons["showPages"].tap() }
        XCTAssertTrue(app.staticTexts["Reviewed on Apple - saved by OfficeIMO"].waitForExistence(timeout: 5), app.debugDescription)
        capture("reopened-note", app)
        if app.buttons["Done"].exists { app.buttons["Done"].tap() }
        XCUIDevice.shared.orientation = .landscapeLeft
        let landscape = XCTNSPredicateExpectation(predicate: NSPredicate { _, _ in
            app.frame.width > app.frame.height
        }, object: nil)
        XCTAssertEqual(XCTWaiter.wait(for: [landscape], timeout: 10), .completed)
        app.buttons["Previous page"].tap()
        XCTAssertTrue(app.staticTexts["1 of 2"].waitForExistence(timeout: 5))
        app.buttons["Next page"].tap()
        XCTAssertTrue(app.staticTexts["2 of 2"].waitForExistence(timeout: 5))
        XCTAssertTrue(add.isHittable)
        capture("workspace-landscape", app)
        XCUIDevice.shared.orientation = .portrait
    }

    private func capture(_ name: String, _ app: XCUIApplication) {
        let attachment = XCTAttachment(screenshot: XCUIScreen.main.screenshot())
        attachment.name = name
        attachment.lifetime = .keepAlways
        add(attachment)
    }
}
