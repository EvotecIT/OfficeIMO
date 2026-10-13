import CoreGraphics
import Foundation
import Observation

/// The engine journey for unattended runs: initialize, open a bundled PDF, show page 1, rotate, undo, redo and save a
/// copy.
@MainActor @Observable
final class JourneyModel {
    struct Step: Identifiable, Codable, Sendable {
        var id = UUID()
        let name: String
        let milliseconds: Double
        let passed: Bool
        let detail: String
    }

    struct LaunchTiming: Codable, Sendable {
        /// From process start to the app's initializer.
        var processToApp: Double?
        /// The first call into the library, which starts its runtime.
        var engineFirstCall: Double?
        var engineInitialize: Double?
    }

    private(set) var steps: [Step] = []
    private(set) var page: CGImage?
    private(set) var facts: DocumentFacts?
    private(set) var isBusy = false
    private(set) var isReady = false
    private(set) var launch = LaunchTiming()
    private(set) var engineVersion = ""

    private var document: EngineDocument?
    private let appStarted = ContinuousClock.now

    let folders = AppFolders()

    init() {
        if let start = ProcessMetrics.processStart() { launch.processToApp = Date().timeIntervalSince(start) * 1000 }
    }

    /// Starts the engine: the first call starts its runtime, then the app names its data folder.
    func start() async {
        guard !isReady else { return }
        let (abi, firstCall) = await measure { await Engine.abiVersion() }
        launch.engineFirstCall = firstCall
        record("engine ABI \(abi)", firstCall, abi == Engine.expectedAbiVersion, "expected \(Engine.expectedAbiVersion)")
        engineVersion = await Engine.libraryVersion()
        do {
            let (_, elapsed) = try await measure { try await Engine.initialize(dataDirectory: folders.engineData) }
            launch.engineInitialize = elapsed
            record("initialize", elapsed, true, folders.engineData.lastPathComponent)
            isReady = true
        } catch {
            record("initialize", 0, false, error.localizedDescription)
        }
    }

    /// Runs every step and checks the document facts after each edit, like the engine's journey script.
    func runJourney() async {
        await start()
        guard isReady, !isBusy else { return }
        isBusy = true
        defer { isBusy = false }
        do {
            let source = try folders.copyJourneySource()
            let firstPage = ContinuousClock.now
            let (opened, openTime) = try await measure { try await EngineDocument.open(source) }
            document = opened
            try await refreshFacts()
            record("open \(source.lastPathComponent)", openTime, facts?.pages.isEmpty == false, "\(facts?.pages.count ?? 0) page(s)")
            try await show(page: 1, step: "render page 1")
            record("first page", milliseconds(since: firstPage), page != nil, "open and render")
            try await thumbnail(page: 1)

            try await edit("rotate page 1") { try await $0.rotatePages([1], degrees: 90) } expect: {
                $0.pages[0].rotation == 90 && $0.isDirty && $0.canUndo
            }
            try await edit("undo") { try await $0.undo() } expect: { $0.pages[0].rotation == 0 && $0.canRedo }
            try await edit("redo") { try await $0.redo() } expect: { $0.pages[0].rotation == 90 }

            let copy = folders.journey.appending(path: "rotated.pdf")
            try? FileManager.default.removeItem(at: copy)
            let (_, saveTime) = try await measure { try await opened.save(to: copy) }
            record("save copy", saveTime, FileManager.default.fileExists(atPath: copy.path), copy.lastPathComponent)

            let reopened = try await EngineDocument.open(copy)
            let saved = try await reopened.facts()
            record("reopen copy", 0, saved.pages[0].rotation == 90 && !saved.isDirty, "rotation \(saved.pages[0].rotation)")
            let unchanged = try Data(contentsOf: source) == Data(contentsOf: folders.bundledJourneyPdf)
            record("source unchanged", 0, unchanged, source.lastPathComponent)
            try await runPagesAndActions()
        } catch {
            record("journey", 0, false, error.localizedDescription)
        }
        try? folders.write(JourneyReport(steps: steps, launch: launch, engineVersion: engineVersion), to: "journey.json")
    }

    /// Inserting another PDF's pages, then the Shortcuts actions on copies of the bundled PDF, as Shortcuts runs them.
    private func runPagesAndActions() async throws {
        guard let document, let count = facts?.pages.count else { return }
        let first = try copyOfBundledPdf(named: "first.pdf")
        let second = try copyOfBundledPdf(named: "second.pdf")
        try await edit("import pages") { _ = try await $0.importPages([ImportSource(url: first)], before: 1) } expect: {
            $0.pages.count == count * 2 && $0.pages[count].rotation == 90
        }
        try await edit("undo import") { try await $0.undo() } expect: { $0.pages.count == count }

        let combined = folders.journey.appending(path: "combined.pdf")
        try? FileManager.default.removeItem(at: combined)
        let (_, combineTime) = try await measure {
            try await DocumentActions.combine([first, second], names: ["first.pdf", "second.pdf"], into: combined)
        }
        let combinedPages = try await EngineDocument.open(combined).facts().pages.count
        record("combine PDFs", combineTime, combinedPages == count * 2, "\(combinedPages) page(s)")

        let extracted = folders.journey.appending(path: "extracted.pdf")
        try? FileManager.default.removeItem(at: extracted)
        let (_, extractTime) = try await measure { try await DocumentActions.extract("1", from: first, named: "first.pdf", into: extracted) }
        let extractedPages = try await EngineDocument.open(extracted).facts().pages.count
        record("extract pages", extractTime, extractedPages == 1, "\(extractedPages) page(s)")

        let (text, textTime) = try await measure { try await DocumentActions.text(of: first, named: "first.pdf") }
        record("get text", textTime, true, "\(text.count) character(s)")

        let searchable = folders.journey.appending(path: "searchable.pdf")
        try? FileManager.default.removeItem(at: searchable)
        let (_, searchableTime) = try await measure { try await DocumentActions.makeSearchable(second, named: "second.pdf", into: searchable) }
        let searchableText = try await DocumentActions.text(of: searchable, named: "searchable.pdf")
        record("make searchable", searchableTime, FileManager.default.fileExists(atPath: searchable.path), "\(searchableText.count) character(s) of text")

        let notes = folders.journey.appending(path: "notes.md")
        try "# Journey notes\n\nConverted on this device.\n".write(to: notes, atomically: true, encoding: .utf8)
        let convertedPDF = folders.journey.appending(path: "notes.pdf")
        try? FileManager.default.removeItem(at: convertedPDF)
        let (_, convertTime) = try await measure { try await DocumentActions.convert(notes, named: "notes.md", into: convertedPDF) }
        let convertedText = try await DocumentActions.text(of: convertedPDF, named: "notes.pdf")
        record("convert to PDF", convertTime, convertedText.contains("Journey notes"), "\(convertedText.count) character(s) of text")

        let image = try await document.renderPage(1, scale: 1)
        let (scan, scanTime) = try await measure { try await DocumentActions.pdf(fromImages: [image.png]) }
        let scanPages = try await EngineDocument.open(scan).facts().pages.count
        record("PDF from images", scanTime, scanPages == 1, "\(scanPages) page(s), \(scan.lastPathComponent)")
    }

    private func copyOfBundledPdf(named name: String) throws -> URL {
        let copy = folders.journey.appending(path: name)
        try? FileManager.default.removeItem(at: copy)
        try FileManager.default.copyItem(at: folders.bundledJourneyPdf, to: copy)
        return copy
    }

    // MARK: Steps

    private func edit(_ name: String, _ change: (EngineDocument) async throws -> Void, expect: (DocumentFacts) -> Bool) async throws {
        guard let document else { return }
        let (_, elapsed) = try await measure { try await change(document) }
        try await refreshFacts()
        record(name, elapsed, facts.map(expect) ?? false, "rotation \(facts?.pages[0].rotation ?? -1)")
        try await show(page: 1, step: "render page 1")
    }

    private func show(page number: Int, step: String) async throws {
        guard let document, let facts else { return }
        let scale = facts.renderScale(page: number, fillingPixelWidth: DisplayTarget.current.pixelWidth)
        let size = facts.pageImageSize(page: number, scale: scale)
        let (image, elapsed) = try await measure {
            try await document.renderRegion(number, scale: scale, x: 0, y: 0, width: size.width, height: size.height)
        }
        page = image
        record(step, elapsed, true, String(format: "%d x %d px at %.2f px/pt", size.width, size.height, scale))
    }

    /// A page thumbnail as PNG, the format for thumbnails, sharing and export.
    private func thumbnail(page number: Int) async throws {
        guard let document else { return }
        let (rendered, renderTime) = try await measure { try await document.renderPage(number, scale: 0.25) }
        let (image, decodeTime) = try await measure { try await Task.detached { try PageImage.decode(rendered.png) }.value }
        record("thumbnail page \(number)", renderTime + decodeTime, image.width == rendered.pixelWidth,
               String(format: "%d x %d px PNG, render %.0f ms, decode %.0f ms", rendered.pixelWidth, rendered.pixelHeight, renderTime, decodeTime))
    }

    private func refreshFacts() async throws {
        facts = try await document?.facts()
    }

    private func record(_ name: String, _ milliseconds: Double, _ passed: Bool, _ detail: String) {
        steps.append(Step(name: name, milliseconds: milliseconds, passed: passed, detail: detail))
    }
}

struct JourneyReport: Codable {
    let steps: [JourneyModel.Step]
    let launch: JourneyModel.LaunchTiming
    let engineVersion: String
    let passed: Bool

    init(steps: [JourneyModel.Step], launch: JourneyModel.LaunchTiming, engineVersion: String) {
        self.steps = steps
        self.launch = launch
        self.engineVersion = engineVersion
        passed = steps.allSatisfy(\.passed)
    }
}

/// The app's folders inside its container.
struct AppFolders: Sendable {
    let engineData: URL
    let journey: URL
    let measurements: URL
    let results: URL
    let bundledJourneyPdf: URL

    /// OFFICEIMO_STUDIO_FOLDER replaces the container folders for unattended measurement runs of an unsandboxed
    /// macOS build, whose container the measuring shell cannot reach.
    init() {
        let manager = FileManager.default
        let override = ProcessInfo.processInfo.environment["OFFICEIMO_STUDIO_FOLDER"].map { URL(fileURLWithPath: $0, isDirectory: true) }
        let support = override ?? manager.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
        let documents = override ?? manager.urls(for: .documentDirectory, in: .userDomainMask)[0]
        engineData = support.appending(path: "Engine")
        journey = documents.appending(path: "Journey")
        measurements = documents.appending(path: "Measurements")
        results = documents.appending(path: "Results")
        bundledJourneyPdf = Bundle.main.url(forResource: "openpreserve-high-resolution-image", withExtension: "pdf")!
        for folder in [journey, measurements, results] {
            try? manager.createDirectory(at: folder, withIntermediateDirectories: true)
        }
    }

    /// Copies the bundled PDF into the container: the bundle is read-only, and the journey must not change its source.
    func copyJourneySource() throws -> URL {
        let source = journey.appending(path: "source.pdf")
        try? FileManager.default.removeItem(at: source)
        try FileManager.default.copyItem(at: bundledJourneyPdf, to: source)
        return source
    }

    /// Writes a result to Documents/Results and to standard output, so an unattended run can collect it without
    /// reading the app's container.
    func write<T: Encodable>(_ value: T, to name: String) throws {
        let encoder = JSONEncoder()
        encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
        let data = try encoder.encode(value)
        try data.write(to: results.appending(path: name), options: .atomic)
        print("OFFICEIMO-RESULT \(name)\n\(String(decoding: data, as: UTF8.self))\nOFFICEIMO-RESULT-END")
    }
}
