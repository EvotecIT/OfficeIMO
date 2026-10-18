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
        app.launchArguments = ["-AppleLanguages", "(en)", "-AppleLocale", "en_US"]
        app.launch()
        #endif
        defer { app.terminate() }

        #if os(macOS)
        // The host launches this exact app with its PDF and language arguments.
        // NSWorkspace ignores those arguments inside the sandboxed test runner.
        let matches = NSRunningApplication.runningApplications(withBundleIdentifier: "com.evotec.officeimo.studio")
            .filter { $0.bundleURL?.standardizedFileURL == appURL.standardizedFileURL }
        let running = try XCTUnwrap(matches.count == 1 ? matches.first : nil)
        app.activate()
        XCTAssertEqual(NSWorkspace.shared.frontmostApplication?.processIdentifier, running.processIdentifier)
        XCTAssertEqual(NSWorkspace.shared.frontmostApplication?.bundleURL?.standardizedFileURL, appURL.standardizedFileURL)
        let window = app.windows.matching(identifier: URL(fileURLWithPath: path).lastPathComponent).firstMatch
        XCTAssertTrue(window.waitForExistence(timeout: 30), "The public dashboard window did not appear.")
        try resizePublicWindow(window: window, appURL: appURL, title: URL(fileURLWithPath: path).lastPathComponent)
        let viewport: XCUIElement = window
        let expectedPID: Int32? = running.processIdentifier
        #else
        if role == "iphone" {
            openPhoneDocument(app, input: URL(fileURLWithPath: path))
        } else {
            app.open(URL(fileURLWithPath: path))
        }
        let viewport: XCUIElement = app
        let expectedPID: Int32? = nil
        #endif

        let page = viewport.descendants(matching: .any).matching(identifier: "Page 1 of 1").firstMatch
        XCTAssertTrue(page.waitForExistence(timeout: 30), viewport.debugDescription)
        try capture(viewport, page: page, expectedPID: expectedPID, name: "studio-dashboard-" + role)

        // Compact iPhone chrome has its own layout. Its full-page capture is
        // sufficient; search captures use the visible regular-width search UI.
        if role != "iphone" {
            let search = viewport.searchFields.firstMatch
            XCTAssertTrue(search.waitForExistence(timeout: 30), viewport.debugDescription)
            search.clickOrTap()
            search.typeText("Quarterly\n")
            XCTAssertTrue(viewport.staticTexts["1 of 1"].waitForExistence(timeout: 30), viewport.debugDescription)
            try capture(viewport, page: page, expectedPID: expectedPID, name: "studio-search-" + role)
        }
    }

    #if os(iOS)
    /// Uses Files' own hierarchy, including a browser restored inside an earlier
    /// input folder. Scope Browse to its tab bar so a breadcrumb cannot match it.
    private func openPhoneDocument(_ app: XCUIApplication, input: URL) {
        let browse = app.tabBars["DOC.browsingModeTabBar"].buttons["Browse"]
        XCTAssertTrue(browse.waitForExistence(timeout: 30), app.debugDescription)
        browse.tap()
        let folder = app.cells.containing(.staticText, identifier: input.deletingLastPathComponent().lastPathComponent).firstMatch
        let studio = app.cells.containing(.staticText, identifier: "OfficeIMO Studio").firstMatch
        let local = app.cells.containing(.staticText, identifier: "On My iPhone").firstMatch
        for _ in 0..<4 {
            if folder.waitForExistence(timeout: 2) { break }
            if studio.exists { studio.tap() }
            else if local.exists { local.tap() }
            else {
                let back = app.navigationBars["FullDocumentManagerViewControllerNavigationBar"].buttons["BackButton"]
                XCTAssertTrue(back.exists, app.debugDescription)
                back.tap()
            }
        }
        XCTAssertTrue(folder.waitForExistence(timeout: 10), app.debugDescription)
        folder.tap()
        let document = app.cells.containing(.staticText, identifier: input.deletingPathExtension().lastPathComponent).firstMatch
        XCTAssertTrue(document.waitForExistence(timeout: 10), app.debugDescription)
        document.tap()
    }
    #endif

    private func capture(_ viewport: XCUIElement, page: XCUIElement, expectedPID: Int32?, name: String) throws {
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
            let screenshot = viewport.screenshot()
            lastSampled = screenshot
            #if os(macOS)
            // Keep queries inside the captured window. A full application tree
            // also populates Open Recent and touches unrelated document paths.
            guard NSWorkspace.shared.frontmostApplication?.processIdentifier == expectedPID else {
                stableSince = nil
                lastError = "The capture app is not in the foreground."
                continue
            }
            let windows = CGWindowListCopyWindowInfo([.optionOnScreenOnly, .excludeDesktopElements], kCGNullWindowID) as? [[String: Any]] ?? []
            let captureFrame = viewport.frame
            guard let processIdentifier = expectedPID,
                  StudioCapturePageReadiness.hasNoPartialWindowOverlap(processIdentifier: processIdentifier,
                      frame: captureFrame, windows: windows) else {
                stableSince = nil
                // Retain only geometry and process identity, never other window
                // titles, so a rejected native run can explain its classification.
                let keys = [kCGWindowOwnerPID, kCGWindowBounds, kCGWindowAlpha, kCGWindowLayer]
                let metadata = windows.map { window in
                    Dictionary(uniqueKeysWithValues: keys.map { ($0 as String, window[$0 as String] ?? NSNull()) })
                }
                lastError = "Another window crosses the capture, or the target window could not be verified. PID=\(expectedPID ?? -1), frame=\(captureFrame), window metadata=\(metadata)"
                continue
            }
            #endif
            guard viewport.progressIndicators.count == 0, viewport.alerts.count == 0,
                  viewport.sheets.count == 0, viewport.dialogs.count == 0 else {
                stableSince = nil
                lastError = "A progress indicator or modal dialog is visible."
                continue
            }
            do {
                guard let source = CGImageSourceCreateWithData(screenshot.pngRepresentation as CFData, nil),
                      let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
                    throw StudioCapturePageReadiness.Failure.unreadablePixels
                }
                let viewportFrame = viewport.frame
                let visible = page.frame.intersection(viewportFrame)
                guard viewportFrame.width > 0, viewportFrame.height > 0, !visible.isNull, !visible.isEmpty else {
                    throw StudioCapturePageReadiness.Failure.unreadablePixels
                }
                let scaleX = CGFloat(image.width) / viewportFrame.width
                let scaleY = CGFloat(image.height) / viewportFrame.height
                let contentBounds = CGRect(x: (visible.minX - viewportFrame.minX) * scaleX,
                                           y: (visible.minY - viewportFrame.minY) * scaleY,
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
    private func resizePublicWindow(window documentWindow: XCUIElement, appURL: URL, title: String) throws {
        let desired = CGSize(width: 1280, height: 800)
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
