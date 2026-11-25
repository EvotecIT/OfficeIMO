import XCTest
import ImageIO
#if os(macOS)
import AppKit
import ApplicationServices
#endif

/// Drives the public document UI against the repository's public dashboard.
/// This file belongs only to the capture test bundles, never the shipping app.
/// The scenes are driven through stable accessibility identifiers, so the same test
/// captures every interface language the host selects.
@MainActor
final class StudioStoreCaptureTests: XCTestCase {
    private struct OpenedDocument {
        let app: XCUIApplication
        let viewport: XCUIElement
        let expectedPID: Int32?
    }

    private let environment = ProcessInfo.processInfo.environment
    private var role: String { environment["STUDIO_CAPTURE_ROLE"] ?? "" }
    /// iPhone and the iPhone Duo use the compact layout; iPad and Mac show the regular toolbar.
    private var isCompact: Bool { role == "iphone" || role == "duo" }

    func testCapturePublicDashboard() throws {
        continueAfterFailure = false
        let opened = try openDocument(environmentKey: "STUDIO_CAPTURE_DOCUMENT_PATH")
        let app = opened.app
        defer { app.terminate() }
        let viewport = opened.viewport

        let page = viewport.descendants(matching: .any).matching(identifier: StudioAccessibilityID.pages).firstMatch
        XCTAssertTrue(page.waitForExistence(timeout: 30), viewport.debugDescription)
        try capture(viewport, region: page, phrases: StudioCapturePageReadiness.dashboardPhrases,
                    expectedPID: opened.expectedPID, name: "studio-dashboard-" + role)
        #if os(iOS)
        assertSingleDocumentBar(app)
        #endif

        // Compact iPhone chrome has its own layout. Its full-page capture is
        // sufficient; search captures use the visible regular-width search UI.
        if !isCompact {
            #if os(iOS)
            let open = app.buttons[StudioAccessibilityID.searchButton]
            XCTAssertTrue(open.waitForExistence(timeout: 30), viewport.debugDescription)
            open.tap()
            let search = viewport.textFields[StudioAccessibilityID.searchField]
            #else
            let search = viewport.searchFields.firstMatch
            #endif
            XCTAssertTrue(search.waitForExistence(timeout: 30), viewport.debugDescription)
            search.clickOrTap()
            search.typeText("Quarterly\n")
            let position = viewport.descendants(matching: .any).matching(identifier: StudioAccessibilityID.searchPosition).firstMatch
            XCTAssertTrue(position.waitForExistence(timeout: 30), viewport.debugDescription)
            try capture(viewport, region: page, phrases: StudioCapturePageReadiness.dashboardPhrases,
                        expectedPID: opened.expectedPID, name: "studio-search-" + role)
            closeSearch(search, in: app)
        }

        if isCompact {
            // Pages uses the task picker, which is hidden during Markup on iPhone.
            selectTask(at: 2, in: app)
            XCTAssertTrue(page.waitForNonExistence(timeout: 10), viewport.debugDescription)
            try capture(viewport, region: nil, phrases: [], expectedPID: opened.expectedPID, name: "studio-pages-" + role)
            selectTask(at: 0, in: app)
            XCTAssertTrue(page.waitForExistence(timeout: 10), viewport.debugDescription)
        }

        // Markup: the drawing tools over the same page.
        selectTask(at: 1, in: app)
        #if os(iOS)
        let done = app.buttons[StudioAccessibilityID.markupDone]
        if isCompact {
            // Confirm the drawing tools even when Done moves into the system overflow.
            let palette = app.descendants(matching: .any).matching(identifier: "Drawing-Palette").firstMatch
            XCTAssertTrue(palette.waitForExistence(timeout: 10) && palette.frame.intersects(viewport.frame), app.debugDescription)
        } else {
            XCTAssertTrue(done.waitForExistence(timeout: 10), app.debugDescription)
        }
        #endif
        XCTAssertTrue(page.waitForExistence(timeout: 10), viewport.debugDescription)
        try capture(viewport, region: page, phrases: StudioCapturePageReadiness.dashboardPhrases,
                    expectedPID: opened.expectedPID, name: "studio-markup-" + role)
        if isCompact { return }
        #if os(iOS)
        done.tap()
        #else
        selectTask(at: 0, in: app)
        #endif
        XCTAssertTrue(page.waitForExistence(timeout: 10), viewport.debugDescription)

        // The inspector beside the page, with the engine's document facts.
        let inspector = app.buttons[StudioAccessibilityID.inspectorToggle]
        XCTAssertTrue(inspector.waitForExistence(timeout: 10), app.debugDescription)
        inspector.clickOrTap()
        let facts = app.descendants(matching: .any).matching(identifier: StudioAccessibilityID.inspectorFileSize).firstMatch
        XCTAssertTrue(facts.waitForExistence(timeout: 10), app.debugDescription)
        try capture(viewport, region: page, phrases: StudioCapturePageReadiness.dashboardPhrases,
                    expectedPID: opened.expectedPID, name: "studio-inspector-" + role)
        inspector.clickOrTap()
        XCTAssertTrue(facts.waitForNonExistence(timeout: 10), app.debugDescription)
        #if os(iOS)
        if role == "ipad" { try verifyRegularDocumentNavigation(app, page: page) }
        #endif
    }

    /// A public fillable PDF form with its fields ready to fill.
    func testCaptureFormDocument() throws {
        continueAfterFailure = false
        let opened = try openDocument(environmentKey: "STUDIO_CAPTURE_FORM_PATH")
        defer { opened.app.terminate() }
        let page = opened.viewport.descendants(matching: .any).matching(identifier: StudioAccessibilityID.pages).firstMatch
        XCTAssertTrue(page.waitForExistence(timeout: 30), opened.viewport.debugDescription)
        try capture(opened.viewport, region: page, phrases: StudioCapturePageReadiness.formPhrases,
                    expectedPID: opened.expectedPID, name: "studio-form-" + role)
    }

    /// A public Excel workbook in the read-only Office viewer.
    func testCaptureOfficeViewer() throws {
        continueAfterFailure = false
        let opened = try openDocument(environmentKey: "STUDIO_CAPTURE_OFFICE_PATH")
        defer { opened.app.terminate() }
        #if os(macOS)
        // The native PDF viewer exposes its scroll view without the SwiftUI wrapper identifier.
        let pages = opened.viewport.scrollViews.firstMatch
        #else
        let pages = opened.viewport.descendants(matching: .any).matching(identifier: StudioAccessibilityID.viewerPages).firstMatch
        #endif
        XCTAssertTrue(pages.waitForExistence(timeout: 60), opened.viewport.debugDescription)
        try capture(opened.viewport, region: pages, phrases: StudioCapturePageReadiness.officePhrases,
                    expectedPID: opened.expectedPID, name: "studio-office-" + role)
    }

    /// Opens one public input in the interface language the host selected.
    private func openDocument(environmentKey: String) throws -> OpenedDocument {
        let path = try XCTUnwrap(environment[environmentKey])
        let title = URL(fileURLWithPath: path).lastPathComponent
        #if os(macOS)
        let appURL = URL(fileURLWithPath: try XCTUnwrap(environment["STUDIO_CAPTURE_APP_PATH"]))
        let app = XCUIApplication(url: appURL)
        // The host launches this exact app with the document and its language arguments.
        // NSWorkspace ignores those arguments inside the sandboxed test runner.
        let matches = NSRunningApplication.runningApplications(withBundleIdentifier: "com.evotec.officeimo.studio")
            .filter { $0.bundleURL?.standardizedFileURL == appURL.standardizedFileURL }
        let running = try XCTUnwrap(matches.count == 1 ? matches.first : nil)
        app.activate()
        XCTAssertEqual(NSWorkspace.shared.frontmostApplication?.processIdentifier, running.processIdentifier)
        XCTAssertEqual(NSWorkspace.shared.frontmostApplication?.bundleURL?.standardizedFileURL, appURL.standardizedFileURL)
        let window = app.windows.matching(identifier: title).firstMatch
        XCTAssertTrue(window.waitForExistence(timeout: 30), "The \(title) window did not appear.")
        try resizePublicWindow(window: window, appURL: appURL, title: title)
        return OpenedDocument(app: app, viewport: window, expectedPID: running.processIdentifier)
        #else
        XCUIDevice.shared.orientation = .portrait
        let app = XCUIApplication(bundleIdentifier: "com.evotec.officeimo.studio")
        let language = environment["STUDIO_CAPTURE_LANGUAGE"] ?? "en"
        let region = environment["STUDIO_CAPTURE_REGION"] ?? "en_US"
        app.launchArguments = ["-AppleLanguages", "(\(language))", "-AppleLocale", region, "-OfficeIMOSkipWelcome"]
        app.launch()
        // The iPhone Duo is an iPhone in both poses, with the iPhone's document browser.
        if isCompact {
            openPhoneDocument(app, input: URL(fileURLWithPath: path))
        } else if environmentKey == "STUDIO_CAPTURE_OFFICE_PATH" {
            // On iPadOS 27, the direct XCTest URL handoff can crash SwiftUI while
            // opening the viewer scene. Selecting the workbook in Browse opens it.
            openPadWorkbook(app, input: URL(fileURLWithPath: path))
        } else {
            app.open(URL(fileURLWithPath: path))
        }
        return OpenedDocument(app: app, viewport: app, expectedPID: nil)
        #endif
    }

    /// Ends the search, so the next scene shows the page without the find bar.
    private func closeSearch(_ search: XCUIElement, in app: XCUIApplication) {
        #if os(iOS)
        let close = app.buttons[StudioAccessibilityID.searchClose]
        XCTAssertTrue(close.waitForExistence(timeout: 5), app.debugDescription)
        close.tap()
        #else
        search.typeKey(.escape, modifierFlags: [])
        #endif
        let position = app.descendants(matching: .any).matching(identifier: StudioAccessibilityID.searchPosition).firstMatch
        XCTAssertTrue(position.waitForNonExistence(timeout: 10), app.debugDescription)
    }

    /// Read, Markup and Pages are the task picker's segments, in that order in every language.
    private func selectTask(at index: Int, in app: XCUIApplication) {
        #if os(macOS)
        // AppKit can expose a segmented picker as a radio group.
        let tasks = app.descendants(matching: .any).matching(identifier: StudioAccessibilityID.taskPicker).firstMatch
        #else
        let tasks = app.segmentedControls[StudioAccessibilityID.taskPicker]
        #endif
        XCTAssertTrue(tasks.waitForExistence(timeout: 10), app.debugDescription)
        #if os(macOS)
        let segments = tasks.radioButtons.count > index ? tasks.radioButtons : tasks.buttons
        #else
        let segments = tasks.buttons
        #endif
        let choice = segments.element(boundBy: index)
        XCTAssertTrue(choice.exists, app.debugDescription)
        choice.clickOrTap()
    }

    #if os(iOS)
    /// The document host supplies one bar in both compact and regular layouts.
    private func assertSingleDocumentBar(_ app: XCUIApplication) {
        let bars = app.navigationBars.allElementsBoundByIndex.filter { $0.isHittable }
        XCTAssertEqual(bars.count, 1, app.debugDescription)
    }

    /// Keep sidebar, inspector and search usable after removing nested navigation.
    private func verifyRegularDocumentNavigation(_ app: XCUIApplication, page: XCUIElement) throws {
        let sidebar = app.segmentedControls[StudioAccessibilityID.sidebarMode]
        XCTAssertTrue(sidebar.waitForExistence(timeout: 5), app.debugDescription)
        let toggle = app.buttons[StudioAccessibilityID.sidebarToggle]
        toggle.tap()
        XCTAssertTrue(sidebar.waitForNonExistence(timeout: 5), app.debugDescription)
        assertSingleDocumentBar(app)
        toggle.tap()
        XCTAssertTrue(sidebar.waitForExistence(timeout: 5), app.debugDescription)

        let inspector = app.buttons[StudioAccessibilityID.inspectorToggle]
        let facts = app.descendants(matching: .any).matching(identifier: StudioAccessibilityID.inspectorFileSize).firstMatch
        inspector.tap()
        XCTAssertTrue(facts.waitForExistence(timeout: 5), app.debugDescription)
        inspector.tap()
        XCTAssertTrue(facts.waitForNonExistence(timeout: 5), app.debugDescription)

        XCUIDevice.shared.orientation = .landscapeLeft
        defer { XCUIDevice.shared.orientation = .portrait }
        XCTAssertTrue(page.waitForExistence(timeout: 10), app.debugDescription)
        assertSingleDocumentBar(app)
        // The application's union can include the document browser's old
        // portrait bounds after rotation. Capture the visible document window.
        try capture(app.windows.firstMatch, region: page, phrases: StudioCapturePageReadiness.dashboardPhrases,
                    expectedPID: nil, name: "studio-navigation-ipad-landscape")

        selectTask(at: 2, in: app)
        XCTAssertTrue(sidebar.waitForNonExistence(timeout: 5), app.debugDescription)
        selectTask(at: 0, in: app)
        XCTAssertTrue(page.waitForExistence(timeout: 10), app.debugDescription)
        XCTAssertTrue(sidebar.waitForExistence(timeout: 5), app.debugDescription)

        selectTask(at: 1, in: app)
        let done = app.buttons[StudioAccessibilityID.markupDone]
        XCTAssertTrue(done.waitForExistence(timeout: 5), app.debugDescription)
        assertSingleDocumentBar(app)
        done.tap()
        XCTAssertTrue(app.segmentedControls[StudioAccessibilityID.taskPicker].waitForExistence(timeout: 5), app.debugDescription)
        XCTAssertTrue(page.waitForExistence(timeout: 10), app.debugDescription)
    }

    /// Selects the staged workbook through the regular-width document browser.
    private func openPadWorkbook(_ app: XCUIApplication, input: URL) {
        func visibleCell(named title: String) -> XCUIElement? {
            let name = NSPredicate(format: "identifier == %@ OR label == %@", title, title)
            return app.cells.matching(name).allElementsBoundByIndex.first { $0.isHittable }
                ?? app.cells.containing(.staticText, identifier: title).allElementsBoundByIndex.first { $0.isHittable }
        }

        let folder = input.deletingLastPathComponent().lastPathComponent
        let sidebar = app.buttons["ToggleSideBar"].firstMatch
        let local = NSPredicate(format: "label CONTAINS[c] %@", "iPad")
        var toggledSidebar = false
        for _ in 0..<6 {
            if let staged = visibleCell(named: folder) {
                staged.tap()
                let title = input.deletingPathExtension().lastPathComponent
                let ready = XCTNSPredicateExpectation(
                    predicate: NSPredicate { _, _ in visibleCell(named: title) != nil }, object: nil)
                XCTAssertEqual(XCTWaiter().wait(for: [ready], timeout: 10), .completed, app.debugDescription)
                guard let workbook = visibleCell(named: title) else {
                    XCTFail(app.debugDescription)
                    return
                }
                workbook.tap()
                return
            }
            if let studio = visibleCell(named: "OfficeIMO Studio") {
                studio.tap()
            } else if let device = app.cells.matching(local).allElementsBoundByIndex.first(where: { $0.isHittable }) {
                device.tap()
            } else if !toggledSidebar && sidebar.waitForExistence(timeout: 15) && sidebar.isHittable {
                sidebar.tap()
                toggledSidebar = true
            } else {
                let navigation = app.navigationBars["FullDocumentManagerViewControllerNavigationBar"]
                let back = ["DOC.navBarButton.backInHistory", "BackButton"]
                    .map { navigation.buttons[$0].firstMatch }
                    .first { $0.exists && $0.isEnabled && $0.isHittable }
                guard let back else {
                    XCTFail(app.debugDescription)
                    return
                }
                back.tap()
            }
        }
        XCTFail("The iPad browser did not show the staged workbook folder. \(app.debugDescription)")
    }

    /// Uses Files' own hierarchy, including a browser restored inside an earlier
    /// input folder. Scope Browse to its tab bar so a breadcrumb cannot match it.
    /// iOS 27 replaces that tab bar with a rail, whose Browse button shows a folder.
    private func openPhoneDocument(_ app: XCUIApplication, input: URL) {
        // Find the Browse tab or rail button by structure and the local device row by its icon.
        let tabs = app.tabBars["DOC.browsingModeTabBar"].buttons
        let rail = app.buttons.containing(.image, identifier: "folder.fill").firstMatch
        let ready = XCTNSPredicateExpectation(
            predicate: NSPredicate { _, _ in tabs.count > 0 || rail.exists }, object: nil)
        XCTAssertEqual(XCTWaiter().wait(for: [ready], timeout: 30), .completed, app.debugDescription)
        let browse = tabs.count > 0 ? tabs.element(boundBy: tabs.count - 1) : rail
        browse.tap()
        let folder = app.cells.containing(.staticText, identifier: input.deletingLastPathComponent().lastPathComponent).firstMatch
        let studio = app.cells.containing(.staticText, identifier: "OfficeIMO Studio").firstMatch
        let local = app.cells.containing(.image, identifier: "iphone").firstMatch
        let namedLocal = app.cells.containing(NSPredicate(format: "elementType == %lu AND label CONTAINS 'iPhone'",
                                                          XCUIElement.ElementType.staticText.rawValue)).firstMatch
        for _ in 0..<4 {
            if folder.waitForExistence(timeout: 2) { break }
            if studio.exists { studio.tap() }
            else if local.exists { local.tap() }
            else if namedLocal.exists { namedLocal.tap() }
            else {
                let back = app.navigationBars["FullDocumentManagerViewControllerNavigationBar"].buttons["BackButton"]
                XCTAssertTrue(back.exists, app.debugDescription)
                back.tap()
            }
        }
        XCTAssertTrue(folder.waitForExistence(timeout: 10), app.debugDescription)
        folder.tap()
        let title = input.deletingPathExtension().lastPathComponent
        let document = app.cells.containing(.staticText, identifier: title).firstMatch
        XCTAssertTrue(document.waitForExistence(timeout: 10), app.debugDescription)
        document.tap()
        if input.pathExtension == "pdf" {
            let pages = app.descendants(matching: .any).matching(identifier: StudioAccessibilityID.pages).firstMatch
            // Files can leave the selected PDF visible in Browse after the first tap.
            if !pages.waitForExistence(timeout: 5), document.exists, document.isHittable {
                document.staticTexts[title].tap()
            }
        }
    }
    #endif

    /// Waits until the scene has stable pixels, then attaches it. With a region, that area must
    /// show the given phrases of the public input; without phrases, only stability is required.
    private func capture(_ viewport: XCUIElement, region: XCUIElement?, phrases: [String], expectedPID: Int32?,
                         name: String) throws {
        var previousPixels: Data?
        var stableSince: TimeInterval?
        var accepted: XCUIScreenshot?
        var lastSampled: XCUIScreenshot?
        var lastError = "The scene has not rendered."
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
                let visible = (region?.frame ?? viewportFrame).intersection(viewportFrame)
                guard viewportFrame.width > 0, viewportFrame.height > 0, !visible.isNull, !visible.isEmpty else {
                    throw StudioCapturePageReadiness.Failure.unreadablePixels
                }
                let scaleX = CGFloat(image.width) / viewportFrame.width
                let scaleY = CGFloat(image.height) / viewportFrame.height
                let contentBounds = CGRect(x: (visible.minX - viewportFrame.minX) * scaleX,
                                           y: (visible.minY - viewportFrame.minY) * scaleY,
                                           width: visible.width * scaleX, height: visible.height * scaleY)
                let pixels = try StudioCapturePageReadiness.pagePixels(from: image, contentBounds: contentBounds)
                lastError = "Waiting for stable pixels in \(contentBounds)."
                let now = ProcessInfo.processInfo.systemUptime
                if pixels != previousPixels || stableSince == nil {
                    previousPixels = pixels
                    stableSince = now
                    continue
                }
                guard now - stableSince! >= 2 else { continue }
                // Recognize the frame we will attach. Cold Vision startup can take longer than the sampling window.
                try StudioCapturePageReadiness.requirePhrases(in: image, contentBounds: contentBounds, phrases: phrases)
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
            // XCTest can drag the window corner without writing through Accessibility.
            let frame = documentWindow.frame
            let origin = documentWindow.coordinate(withNormalizedOffset: .zero)
            let corner = origin.withOffset(CGVector(dx: frame.width - 2, dy: frame.height - 2))
            let destination = origin.withOffset(CGVector(dx: desired.width - 2, dy: desired.height - 2))
            let cornerPoint = corner.screenPoint
            let destinationPoint = destination.screenPoint
            let screenKey = NSDeviceDescriptionKey("NSScreenNumber")
            let onScreen = NSScreen.screens.contains { screen in
                guard let number = screen.deviceDescription[screenKey] as? NSNumber else { return false }
                let bounds = CGDisplayBounds(number.uint32Value)
                return bounds.contains(cornerPoint) && bounds.contains(destinationPoint)
            }
            let windows = CGWindowListCopyWindowInfo([.optionOnScreenOnly, .excludeDesktopElements], kCGNullWindowID) as? [[String: Any]] ?? []
            if onScreen, documentWindow.isHittable,
               let front = NSWorkspace.shared.frontmostApplication,
               front.bundleURL?.standardizedFileURL == appURL.standardizedFileURL,
               StudioCapturePageReadiness.hasNoPartialWindowOverlap(processIdentifier: front.processIdentifier,
                                                                   frame: frame, windows: windows) {
                corner.click(forDuration: 0.15, thenDragTo: destination)
            }
            // A runner can also prepare the window. Keep the exact size requirement.
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
