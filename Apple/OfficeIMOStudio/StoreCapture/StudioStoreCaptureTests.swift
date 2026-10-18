import XCTest
import ImageIO
#if os(macOS)
import AppKit
import ApplicationServices
#endif

/// Drives the public document UI against the repository's public dashboard.
/// This file belongs only to the capture test bundles, never the shipping app.
@MainActor
final class StudioStoreCaptureTests: XCTestCase {
    func testCapturePublicDashboard() throws {
        continueAfterFailure = false
        let environment = ProcessInfo.processInfo.environment
        let path = try XCTUnwrap(environment["STUDIO_CAPTURE_DOCUMENT_PATH"])
        let role = try XCTUnwrap(environment["STUDIO_CAPTURE_ROLE"])
        #if os(macOS)
        let appURL = URL(fileURLWithPath: try XCTUnwrap(environment["STUDIO_CAPTURE_APP_PATH"]))
        let app = XCUIApplication(url: appURL)
        app.launchEnvironment["TMPDIR"] = try XCTUnwrap(environment["STUDIO_CAPTURE_NATIVE_TMPDIR"])
        #else
        XCUIDevice.shared.orientation = .portrait
        let app = XCUIApplication(bundleIdentifier: "com.evotec.officeimo.studio")
        #endif
        app.launchArguments = ["-AppleLanguages", "(en)", "-AppleLocale", "en_US"]
        #if os(macOS)
        // Skip earlier document restoration for this capture launch without
        // deleting documents from an interactive session.
        app.launchArguments += ["-ApplePersistenceIgnoreState", "YES"]
        #endif
        app.launch()
        defer { app.terminate() }

        #if os(macOS)
        // Target this exact built app through the normal document-open route.
        // The Open panel can restore an unrelated removable-volume location;
        // bundle-only URL dispatch can instead choose another installed build.
        let running = try XCTUnwrap(NSRunningApplication.runningApplications(withBundleIdentifier: "com.evotec.officeimo.studio")
            .first { $0.bundleURL?.standardizedFileURL == appURL.standardizedFileURL })
        let expectedPID = running.processIdentifier
        let opened = expectation(description: "The exact capture app receives the public PDF")
        NSWorkspace.shared.open([URL(fileURLWithPath: path)], withApplicationAt: appURL,
                                configuration: NSWorkspace.OpenConfiguration()) { application, error in
            XCTAssertNil(error)
            XCTAssertEqual(application?.bundleURL?.standardizedFileURL, appURL.standardizedFileURL)
            XCTAssertEqual(application?.processIdentifier, expectedPID)
            opened.fulfill()
        }
        wait(for: [opened], timeout: 15)
        let window = app.windows[URL(fileURLWithPath: path).lastPathComponent]
        XCTAssertTrue(window.waitForExistence(timeout: 30), app.debugDescription)
        try resizePublicWindow(app: app, appURL: appURL, title: URL(fileURLWithPath: path).lastPathComponent)
        #else
        if role == "iphone" {
            // The compact launch scene ignores a URL sent while it is running.
            // Open the copied input through its normal native document browser.
            let browse = app.buttons["Browse"]
            XCTAssertTrue(browse.waitForExistence(timeout: 30), app.debugDescription)
            browse.tap()
            let input = URL(fileURLWithPath: path)
            let folder = app.cells.containing(.staticText, identifier: input.deletingLastPathComponent().lastPathComponent).firstMatch
            if !folder.waitForExistence(timeout: 3), app.buttons["OfficeIMO Studio"].exists {
                // A preceding capture may have left the browser in its now
                // removed input folder. Return through the native breadcrumb.
                app.buttons["OfficeIMO Studio"].tap()
            }
            if !folder.waitForExistence(timeout: 3) {
                let local = app.cells.containing(.staticText, identifier: "On My iPhone").firstMatch
                if local.exists {
                    local.tap()
                } else {
                    let back = app.buttons["On My iPhone"]
                    XCTAssertTrue(back.waitForExistence(timeout: 10), app.debugDescription)
                    back.tap()
                }
                let studio = app.cells.containing(.staticText, identifier: "OfficeIMO Studio").firstMatch
                XCTAssertTrue(studio.waitForExistence(timeout: 10), app.debugDescription)
                studio.tap()
            }
            XCTAssertTrue(folder.waitForExistence(timeout: 10), app.debugDescription)
            folder.tap()
            let document = app.cells.containing(.staticText, identifier: input.deletingPathExtension().lastPathComponent).firstMatch
            XCTAssertTrue(document.waitForExistence(timeout: 10), app.debugDescription)
            document.tap()
        } else {
            app.open(URL(fileURLWithPath: path))
        }
        #endif

        let page = app.descendants(matching: .any).matching(identifier: "Page 1 of 1").firstMatch
        XCTAssertTrue(page.waitForExistence(timeout: 30), app.debugDescription)
        try capture(app, page: page, name: "studio-dashboard-" + role)

        // Compact iPhone chrome has its own layout. Its full-page capture is
        // sufficient; search captures use the visible regular-width search UI.
        if role != "iphone" {
            let search = app.searchFields.firstMatch
            XCTAssertTrue(search.waitForExistence(timeout: 30), app.debugDescription)
            search.clickOrTap()
            search.typeText("Quarterly\n")
            XCTAssertTrue(app.staticTexts["1 of 1"].waitForExistence(timeout: 30), app.debugDescription)
            try capture(app, page: page, name: "studio-search-" + role)
        }
    }

    private func snapshot(_ app: XCUIApplication) -> XCUIScreenshot {
        #if os(macOS)
        let path = ProcessInfo.processInfo.environment["STUDIO_CAPTURE_DOCUMENT_PATH"]!
        return app.windows[URL(fileURLWithPath: path).lastPathComponent].screenshot()
        #else
        return app.screenshot()
        #endif
    }

    private func capture(_ app: XCUIApplication, page: XCUIElement, name: String) throws {
        var previousPixels: Data?
        var stableSince: TimeInterval?
        var accepted: XCUIScreenshot?
        var lastSampled: XCUIScreenshot?
        var lastError = "The public dashboard has not rendered."
        let deadline = ProcessInfo.processInfo.systemUptime + 45
        while ProcessInfo.processInfo.systemUptime < deadline {
            // Sample serially: an expectation callback performing OCR can still
            // be running when its waiter times out and the app is terminated.
            Thread.sleep(forTimeInterval: 0.25)
            let screenshot = snapshot(app)
            lastSampled = screenshot
            guard app.progressIndicators.count == 0, app.alerts.count == 0,
                  app.sheets.count == 0, app.dialogs.count == 0 else {
                stableSince = nil
                lastError = "A progress indicator or modal dialog is visible."
                continue
            }
            do {
                guard let source = CGImageSourceCreateWithData(screenshot.pngRepresentation as CFData, nil),
                      let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
                    throw StudioCapturePageReadiness.Failure.unreadablePixels
                }
                #if os(macOS)
                let path = ProcessInfo.processInfo.environment["STUDIO_CAPTURE_DOCUMENT_PATH"]!
                let viewport = app.windows[URL(fileURLWithPath: path).lastPathComponent].frame
                #else
                let viewport = app.frame
                #endif
                let visible = page.frame.intersection(viewport)
                guard viewport.width > 0, viewport.height > 0, !visible.isNull, !visible.isEmpty else {
                    throw StudioCapturePageReadiness.Failure.unreadablePixels
                }
                let scaleX = CGFloat(image.width) / viewport.width
                let scaleY = CGFloat(image.height) / viewport.height
                let contentBounds = CGRect(x: (visible.minX - viewport.minX) * scaleX,
                                           y: (visible.minY - viewport.minY) * scaleY,
                                           width: visible.width * scaleX, height: visible.height * scaleY)
                let pixels = try StudioCapturePageReadiness.pagePixels(from: image, contentBounds: contentBounds)
                lastError = "The public PDF is recognized; waiting for stable pixels in \(contentBounds)."
                let now = ProcessInfo.processInfo.systemUptime
                if pixels != previousPixels || stableSince == nil {
                    previousPixels = pixels
                    stableSince = now
                    continue
                }
                guard now - stableSince! >= 2 else { continue }
                accepted = screenshot
                break
            } catch {
                stableSince = nil
                lastError = String(describing: error)
                continue
            }
        }
        if accepted == nil, let lastSampled {
            // Keep the actual last sample without issuing another UI request
            // after the capture deadline.
            let rejected = XCTAttachment(screenshot: lastSampled)
            rejected.name = name + "-rejected"
            rejected.lifetime = .keepAlways
            add(rejected)
        }
        let screenshot = try XCTUnwrap(accepted, lastError)
        let attachment = XCTAttachment(screenshot: screenshot)
        attachment.name = name
        attachment.lifetime = .keepAlways
        add(attachment)
    }

    #if os(macOS)
    /// Accepts the required window size or uses existing Accessibility trust to
    /// resize only this app's public document window. It never requests permissions.
    private func resizePublicWindow(app: XCUIApplication, appURL: URL, title: String) throws {
        let desired = CGSize(width: 1280, height: 800)
        let documentWindow = app.windows[title]
        if documentWindow.frame.size == desired { return }
        guard AXIsProcessTrusted() else {
            // An operator may prepare this window with an already authorized
            // resize tool. The same dimensions remain mandatory for the capture.
            let sized = XCTNSPredicateExpectation(predicate: NSPredicate { _, _ in
                documentWindow.frame.size == desired
            }, object: nil)
            let status = XCTWaiter.wait(for: [sized], timeout: 30)
            XCTAssertEqual(status, .completed, "Prepare the 1280x800 document window or use a runner with existing Accessibility access.")
            guard status == .completed else { throw CaptureError.windowUnavailable }
            return
        }
        let matches = NSRunningApplication.runningApplications(withBundleIdentifier: "com.evotec.officeimo.studio")
            .filter { $0.bundleURL?.standardizedFileURL == appURL.standardizedFileURL }
        let running = try XCTUnwrap(matches.count == 1 ? matches.first : nil)
        let application = AXUIElementCreateApplication(running.processIdentifier)
        var raw: CFTypeRef?
        guard AXUIElementCopyAttributeValue(application, kAXWindowsAttribute as CFString, &raw) == .success,
              let windows = raw as? [AXUIElement] else { throw CaptureError.windowUnavailable }
        let documents = windows.filter { window in
            var value: CFTypeRef?
            return AXUIElementCopyAttributeValue(window, kAXTitleAttribute as CFString, &value) == .success &&
                (value as? String) == title
        }
        let window = try XCTUnwrap(documents.count == 1 ? documents.first : nil)
        var requested = desired
        let value = try XCTUnwrap(AXValueCreate(.cgSize, &requested))
        guard AXUIElementSetAttributeValue(window, kAXSizeAttribute as CFString, value) == .success,
              AXUIElementCopyAttributeValue(window, kAXSizeAttribute as CFString, &raw) == .success,
              let actual = raw, CFGetTypeID(actual) == AXValueGetTypeID() else { throw CaptureError.windowUnavailable }
        var size = CGSize.zero
        guard AXValueGetValue(actual as! AXValue, .cgSize, &size), size == desired else {
            throw CaptureError.windowUnavailable
        }
    }
    private enum CaptureError: Error { case windowUnavailable }
    #endif
}

private extension XCUIElement {
    func clickOrTap() {
        #if os(macOS)
        click()
        #else
        tap()
        #endif
    }
}
