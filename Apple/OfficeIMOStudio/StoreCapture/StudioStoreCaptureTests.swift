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
        // Use the app's File > Open panel. Opening a URL through Launch Services
        // can select another installed product with the same bundle identifier.
        app.typeKey("o", modifierFlags: .command)
        app.typeKey("g", modifierFlags: [.command, .shift])
        app.typeText(path + "\n")
        app.typeKey(.return, modifierFlags: [])
        let window = app.windows[URL(fileURLWithPath: path).lastPathComponent]
        XCTAssertTrue(window.waitForExistence(timeout: 30), app.debugDescription)
        try resizePublicWindow(appURL: appURL, title: URL(fileURLWithPath: path).lastPathComponent)
        #else
        app.open(URL(fileURLWithPath: path))
        #endif

        let page = app.descendants(matching: .any).matching(identifier: "Page 1 of 1").firstMatch
        XCTAssertTrue(page.waitForExistence(timeout: 30), app.debugDescription)
        try capture(app, name: "studio-dashboard-" + role)

        // Compact iPhone chrome has its own layout. Its full-page capture is
        // sufficient; search captures use the visible regular-width search UI.
        if role != "iphone" {
            let search = app.searchFields.firstMatch
            XCTAssertTrue(search.waitForExistence(timeout: 30), app.debugDescription)
            search.clickOrTap()
            search.typeText("Quarterly\n")
            XCTAssertTrue(app.staticTexts["1 of 1"].waitForExistence(timeout: 30), app.debugDescription)
            try capture(app, name: "studio-search-" + role)
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

    private func capture(_ app: XCUIApplication, name: String) throws {
        var previousPixels: Data?
        var stableSince: TimeInterval?
        var accepted: XCUIScreenshot?
        var lastError = "The public dashboard has not rendered."
        let rendered = XCTNSPredicateExpectation(predicate: NSPredicate { _, _ in
            guard app.progressIndicators.count == 0, app.alerts.count == 0 else {
                stableSince = nil
                return false
            }
            let screenshot = self.snapshot(app)
            do {
                guard let source = CGImageSourceCreateWithData(screenshot.pngRepresentation as CFData, nil),
                      let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
                    throw StudioCapturePageReadiness.Failure.unreadablePixels
                }
                let pixels = try StudioCapturePageReadiness.pagePixels(from: image)
                let now = ProcessInfo.processInfo.systemUptime
                if pixels != previousPixels || stableSince == nil {
                    previousPixels = pixels
                    stableSince = now
                    return false
                }
                guard now - stableSince! >= 2 else { return false }
                accepted = screenshot
                return true
            } catch {
                stableSince = nil
                lastError = String(describing: error)
                return false
            }
        }, object: nil)
        let status = XCTWaiter.wait(for: [rendered], timeout: 45)
        if status != .completed {
            let rejected = XCTAttachment(screenshot: snapshot(app))
            rejected.name = name + "-rejected"
            rejected.lifetime = .keepAlways
            add(rejected)
        }
        XCTAssertEqual(status, .completed, lastError)
        let screenshot = try XCTUnwrap(accepted, lastError)
        let attachment = XCTAttachment(screenshot: screenshot)
        attachment.name = name
        attachment.lifetime = .keepAlways
        add(attachment)
    }

    #if os(macOS)
    /// Uses existing Accessibility trust and touches only the exact capture app's
    /// public document window. It never requests or changes system permissions.
    private func resizePublicWindow(appURL: URL, title: String) throws {
        XCTAssertTrue(AXIsProcessTrusted(), "The capture runner needs existing Accessibility access.")
        guard AXIsProcessTrusted() else { throw CaptureError.windowUnavailable }
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
        var desired = CGSize(width: 1280, height: 800)
        let value = try XCTUnwrap(AXValueCreate(.cgSize, &desired))
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
