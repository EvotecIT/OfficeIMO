import SwiftUI

@main
enum StudioMain {
    static func main() {
        #if DEBUG && os(iOS)
        if DevelopmentHarness.file != nil {
            DevelopmentHarnessApp.main()
            return
        }
        #endif
        OfficeIMOStudioApp.main()
    }
}

struct OfficeIMOStudioApp: App {
    #if os(macOS)
    @NSApplicationDelegateAdaptor(StudioAppDelegate.self) private var appDelegate
    #endif

    init() {
        StudioDocument.removeStaleWorkingCopies()
        // Unattended runs start with the app rather than a window, which may never appear when launched from a shell.
        if Autorun.isRequested {
            let model = JourneyModel()
            Task { await Autorun.run(model) }
        }
    }

    var body: some Scene {
        DocumentGroup(newDocument: { StudioDocument() }) { configuration in
            #if os(macOS)
            if let source = configuration.document.officeSource {
                OfficeDocumentView(source: source)
            } else {
                DocumentWindow(document: configuration.document, fileURL: configuration.fileURL)
            }
            #else
            DocumentWindow(document: configuration.document, fileURL: configuration.fileURL)
            #endif
        }
        #if os(macOS)
        .defaultSize(width: 1180, height: 860)
        #endif
        .commands {
            StudioCommands()
            #if os(macOS)
            // File > Import from iPhone or iPad: Continuity Camera's Scan Documents and Take Photo.
            ImportFromDevicesCommands()
            #endif
        }

        #if os(macOS)
        // A WindowGroup, as a Window scene beside a DocumentGroup crashes in SwiftUI at launch.
        WindowGroup("OfficeIMO Studio", id: StartWindow.id) { StartWindow() }
            .windowResizability(.contentSize)
            .defaultLaunchBehavior(.presented)
            .restorationBehavior(.disabled)
        #endif

        #if os(iOS)
        DocumentGroupLaunchScene("OfficeIMO Studio", backgroundStyle: Color.accentColor.gradient) {
            // The scene folds buttons after the first two into a More menu, so the other ways to start share one.
            Button("Convert to PDF") { LaunchDocuments.open { try await ConversionSources.newDocumentFromFile() } }
            Button("Create PDF…") { LaunchDocuments.open { try await ScanSources.newDocumentFromChoice() } }
        }
        #endif
    }
}

/// The menu bar on the Mac and the iPad, with Preview's shortcuts where Preview has the same command.
struct StudioCommands: Commands {
    @FocusedValue(\.documentSession) private var session
    #if os(macOS)
    @Environment(\.newDocument) private var newDocument
    #endif

    var body: some Commands {
        #if os(macOS)
        CommandGroup(after: .newItem) {
            Button("New PDF from Images…") { ScanSources.newDocumentFromImageFiles(newDocument) }
            Button("Convert to PDF…") { ConversionSources.newDocumentFromFile(newDocument) }
            Button("Combine Files…") { ConversionSources.newDocumentFromFiles(newDocument) }
            Button("Try a Sample PDF") { SampleDocument.open(newDocument) }
        }
        #endif

        CommandGroup(replacing: .printItem) {
            Button("Print…") {
                #if os(macOS)
                session?.print(from: NSApp.keyWindow)
                #else
                session?.print()
                #endif
            }
            .keyboardShortcut("p")
            .disabled(session?.canPrint != true)
        }

        CommandGroup(after: .importExport) {
            Section {
                if let session { SecurityCommands(session: session) }
            }
        }

        CommandGroup(before: .toolbar) {
            Group {
                ForEach(Array(StudioTask.allCases.enumerated()), id: \.element) { index, task in
                    Toggle(isOn: Binding { session?.task == task } set: { if $0 { session?.task = task } }) {
                        Text(task.title)
                    }
                    .keyboardShortcut(KeyEquivalent(Character("\(index + 1)")), modifiers: .command)
                }
                Divider()
                Picker("Reading Layout", selection: Binding { session?.readingLayout ?? .continuous } set: { session?.readingLayout = $0 }) {
                    ForEach(ReadingLayout.allCases) { Text($0.title).tag($0) }
                }
                .pickerStyle(.inline)
                Button("Rotate View Left") { session?.rotateView(clockwise: false) }
                Button("Rotate View Right") { session?.rotateView(clockwise: true) }
                Divider()
                Button("Actual Size") { session?.zoom(.actualSize) }
                    .keyboardShortcut("0")
                Button("Zoom to Fit") { session?.zoom(.fitWidth) }
                    .keyboardShortcut("9")
                Button("Zoom In") { session?.zoom(.zoomIn) }
                    .keyboardShortcut("+")
                Button("Zoom Out") { session?.zoom(.zoomOut) }
                    .keyboardShortcut("-")
                Divider()
                Toggle(isOn: Binding { session?.sidebarMode == .thumbnails } set: { if $0 { session?.sidebarMode = .thumbnails; session?.isPageListShown = true } }) {
                    Text("Thumbnails")
                }
                .keyboardShortcut("2", modifiers: [.command, .option])
                Toggle(isOn: Binding { session?.sidebarMode == .outline } set: { if $0 { session?.sidebarMode = .outline; session?.isPageListShown = true } }) {
                    Text("Table of Contents")
                }
                .keyboardShortcut("3", modifiers: [.command, .option])
                Divider()
                Button(session?.isInspectorShown == true ? "Hide Inspector" : "Show Inspector") { session?.isInspectorShown.toggle() }
                    .keyboardShortcut("i", modifiers: [.command, .option])
            }
            .disabled(session?.isReady != true)
            Divider()
        }

        CommandMenu("Go") {
            Button("Previous Page") { session?.showPreviousPage() }
                .keyboardShortcut(.upArrow, modifiers: [.command, .option])
                .disabled(session.map { $0.currentPage <= 1 } ?? true)
            Button("Next Page") { session?.showNextPage() }
                .keyboardShortcut(.downArrow, modifiers: [.command, .option])
                .disabled(session.map { $0.currentPage >= $0.pageCount } ?? true)
        }

        CommandGroup(after: .textEditing) {
            Section {
                Button("Find…") {
                    session?.task = .read
                    session?.isSearchFocused = true
                }
                .keyboardShortcut("f")
                Button("Find Next") { session?.showHit(step: 1) }
                    .keyboardShortcut("g")
                    .disabled(session?.searchHits.isEmpty != false)
                Button("Find Previous") { session?.showHit(step: -1) }
                    .keyboardShortcut("g", modifiers: [.command, .shift])
                    .disabled(session?.searchHits.isEmpty != false)
            }
            .disabled(session?.isReady != true || session?.facts?.canExtractText == false)
        }

        CommandMenu("Tools") {
            Group {
                Button("Rotate Left") { rotate(-90) }
                    .keyboardShortcut("l")
                Button("Rotate Right") { rotate(90) }
                    .keyboardShortcut("r")
            }
            .disabled(session?.canRotatePages != true || (session?.task == .pages && session?.selection.isEmpty == true))
            Divider()
            Group {
                Button("Duplicate Pages") { session.map { $0.duplicate($0.targetPages) } }
                    .keyboardShortcut("d")
                Button("Delete Pages") { session.map { $0.delete($0.targetPages) } }
                    .disabled((session?.targetPages.count ?? 0) >= (session?.pageCount ?? 0))
            }
            .disabled(session?.canEditPages != true || (session?.task == .pages && session?.selection.isEmpty == true))
            Button("Insert Blank Page") { session?.insertBlankPage() }
                .disabled(session?.canInsertBlankPages != true || (session?.task == .pages && session?.selection.isEmpty == true))
            Button("Insert Pages from File…") { session?.isChoosingPagesToImport = true }
                .disabled(session?.canImportPages != true)
            Divider()
            Button("Extract Pages…") { session.map { $0.extract($0.targetPages) } }
                .disabled(session?.canExtractPages != true || (session?.task == .pages && session?.selection.isEmpty == true))
            Divider()
            if let session { StructureCommands(session: session) }
        }
    }

    private func rotate(_ degrees: Int) {
        guard let session, session.isReady else { return }
        session.rotate(session.targetPages, degrees: degrees)
    }
}

/// Launch arguments for unattended runs:
/// -OfficeIMOJourney runs the journey and writes Results/journey.json;
/// -OfficeIMOMeasure <file> measures Documents/Measurements/<file> and writes Results/measure-<file>.json;
/// -OfficeIMOExit quits afterwards.
@MainActor
enum Autorun {
    static var isRequested: Bool {
        ProcessInfo.processInfo.arguments.contains { ["-OfficeIMOJourney", "-OfficeIMOMeasure", "-OfficeIMOExit"].contains($0) }
    }

    static func run(_ model: JourneyModel) async {
        let arguments = ProcessInfo.processInfo.arguments
        #if os(iOS)
        // The measurements size pages for the screen, which is known once a scene connects.
        for _ in 0..<50 where UIApplication.shared.connectedScenes.isEmpty {
            try? await Task.sleep(for: .milliseconds(100))
        }
        #endif
        await model.start()
        if arguments.contains("-OfficeIMOJourney") {
            await model.runJourney()
        }
        if let index = arguments.firstIndex(of: "-OfficeIMOMeasure"), index + 1 < arguments.count {
            let name = arguments[index + 1]
            do {
                let report = try await MeasurementRunner.run(model.folders.measurements.appending(path: name),
                                                             display: .current, launch: model.launch, engineVersion: model.engineVersion)
                try model.folders.write(report, to: "measure-\(name).json")
            } catch {
                try? model.folders.write(["file": name, "error": error.localizedDescription], to: "measure-\(name).json")
            }
        }
        if arguments.contains("-OfficeIMOExit") {
            exit(0)
        }
    }
}
