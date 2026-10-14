import AppIntents
import Foundation
import UniformTypeIdentifiers

/// Shortcuts actions over the same engine as the document windows. Each runs without opening the app, works on
/// copies of the files it is given, and returns a new file or text.
struct CombinePDFsIntent: AppIntent {
    static let title: LocalizedStringResource = "Combine into PDF"
    static let description = IntentDescription(
        "Puts PDFs, images and Word, Excel, PowerPoint and other documents together into one PDF, in the order given.")

    @Parameter(title: "Files", supportedContentTypes: PageSources.contentTypes)
    var files: [IntentFile]

    @Parameter(title: "Name", default: "Combined")
    var name: String

    static var parameterSummary: some ParameterSummary {
        Summary("Combine \(\.$files)") { \.$name }
    }

    func perform() async throws -> some IntentResult & ReturnsValue<IntentFile> {
        let work = try await ActionFolder.prepare()
        // Each copy keeps its extension, so it is read as what it is: a PDF, an image or a document to convert.
        let inputs = try files.enumerated().map { index, file in
            let fileExtension = (file.filename as NSString).pathExtension
            return try work.copy(file, as: fileExtension.isEmpty ? "\(index + 1).pdf" : "\(index + 1).\(fileExtension)")
        }
        let output = try await DocumentActions.combine(inputs, names: files.map(\.filename), into: work.output(named: name))
        return .result(value: work.takeOutput(output))
    }
}

struct MakePDFFromImagesIntent: AppIntent {
    static let title: LocalizedStringResource = "Make PDF from Images"
    static let description = IntentDescription(
        "Makes a PDF with a page for each image, with its text recognized on this device, so the PDF can be searched and its text selected.")

    @Parameter(title: "Images", supportedContentTypes: [.image])
    var images: [IntentFile]

    static var parameterSummary: some ParameterSummary {
        Summary("Make a PDF from \(\.$images)")
    }

    func perform() async throws -> some IntentResult & ReturnsValue<IntentFile> {
        try await EngineRuntime.ready()
        let url = try await DocumentActions.pdf(fromImages: images.map(\.data))
        StudioTemporaryFiles.markReturnedOutput(url)
        return .result(value: IntentFile(fileURL: url, filename: url.lastPathComponent, type: .pdf))
    }
}

struct ExtractPagesIntent: AppIntent {
    static let title: LocalizedStringResource = "Extract PDF Pages"
    static let description = IntentDescription("Makes a new PDF from some of a PDF’s pages, such as “1-3, 5”, in the order given.")

    @Parameter(title: "PDF", supportedContentTypes: [.pdf])
    var file: IntentFile

    @Parameter(title: "Pages", description: "Page numbers and ranges, such as 1-3, 5.")
    var pages: String

    static var parameterSummary: some ParameterSummary {
        Summary("Extract pages \(\.$pages) from \(\.$file)")
    }

    func perform() async throws -> some IntentResult & ReturnsValue<IntentFile> {
        let work = try await ActionFolder.prepare()
        let base = (file.filename as NSString).deletingPathExtension
        let output = try await DocumentActions.extract(pages, from: try work.copy(file, as: "source.pdf"), named: file.filename,
                                                       into: work.output(named: String(localized: "\(base) – Pages \(pages)")))
        return .result(value: work.takeOutput(output))
    }
}

struct GetPDFTextIntent: AppIntent {
    static let title: LocalizedStringResource = "Get Text from PDF"
    static let description = IntentDescription("Returns a PDF’s text in reading order, a blank line between pages.")

    @Parameter(title: "PDF", supportedContentTypes: [.pdf])
    var file: IntentFile

    static var parameterSummary: some ParameterSummary {
        Summary("Get the text of \(\.$file)")
    }

    func perform() async throws -> some IntentResult & ReturnsValue<String> {
        let work = try await ActionFolder.prepare()
        return .result(value: try await DocumentActions.text(of: try work.copy(file, as: "source.pdf"), named: file.filename))
    }
}

struct MakePDFSearchableIntent: AppIntent {
    static let title: LocalizedStringResource = "Make PDF Searchable"
    static let description = IntentDescription(
        "Reads the text of a PDF’s scanned pages on this device and adds it, so the PDF can be searched and its text selected. Pages that already have text keep it.")

    @Parameter(title: "PDF", supportedContentTypes: [.pdf])
    var file: IntentFile

    static var parameterSummary: some ParameterSummary {
        Summary("Make \(\.$file) searchable")
    }

    func perform() async throws -> some IntentResult & ReturnsValue<IntentFile> {
        let work = try await ActionFolder.prepare()
        let base = (file.filename as NSString).deletingPathExtension
        let output = try await DocumentActions.makeSearchable(try work.copy(file, as: "source.pdf"), named: file.filename,
                                                              into: work.output(named: base))
        return .result(value: work.takeOutput(output))
    }
}

struct ConvertToPDFIntent: AppIntent {
    static let title: LocalizedStringResource = "Convert to PDF"
    static let description = IntentDescription(
        "Converts a Word, Excel, PowerPoint, RTF, text, HTML, Markdown, XPS, Visio, Draw or Publisher document to PDF on this device.")

    @Parameter(title: "Document")
    var file: IntentFile

    static var parameterSummary: some ParameterSummary {
        Summary("Convert \(\.$file) to PDF")
    }

    func perform() async throws -> some IntentResult & ReturnsValue<IntentFile> {
        let work = try await ActionFolder.prepare()
        let input = try work.copy(file, as: file.filename)
        let output = try await DocumentActions.convert(input, named: file.filename,
                                                       into: work.output(named: (file.filename as NSString).deletingPathExtension))
        return .result(value: work.takeOutput(output))
    }
}

/// The phrases Siri and Spotlight offer for the actions.
struct StudioShortcuts: AppShortcutsProvider {
    static var appShortcuts: [AppShortcut] {
        AppShortcut(intent: CombinePDFsIntent(), phrases: ["Combine PDFs with \(.applicationName)",
                                                           "Combine files into a PDF with \(.applicationName)"],
                    shortTitle: "Combine into PDF", systemImageName: "doc.on.doc")
        AppShortcut(intent: MakePDFFromImagesIntent(), phrases: ["Make a PDF from images with \(.applicationName)"],
                    shortTitle: "PDF from Images", systemImageName: "doc.viewfinder")
        AppShortcut(intent: ExtractPagesIntent(), phrases: ["Extract PDF pages with \(.applicationName)"],
                    shortTitle: "Extract Pages", systemImageName: "doc.badge.arrow.up")
        AppShortcut(intent: ConvertToPDFIntent(), phrases: ["Convert a document to PDF with \(.applicationName)"],
                    shortTitle: "Convert to PDF", systemImageName: "doc.richtext")
        AppShortcut(intent: MakePDFSearchableIntent(), phrases: ["Make a PDF searchable with \(.applicationName)"],
                    shortTitle: "Make Searchable", systemImageName: "text.viewfinder")
        AppShortcut(intent: GetPDFTextIntent(), phrases: ["Get text from a PDF with \(.applicationName)"],
                    shortTitle: "Get Text", systemImageName: "text.page")
    }
}

/// A failure the action reports to the person running it, in words they can act on.
enum IntentProblem: Error, CustomLocalizedStringResourceConvertible {
    case message(String)

    var localizedStringResource: LocalizedStringResource {
        switch self {
        case .message(let text): "\(text)"
        }
    }
}

/// Page numbers and ranges, such as "1-3, 5" or "4–2", in the order written.
enum PageRange {
    static func parse(_ text: String, pageCount: Int) -> [Int]? {
        var pages: [Int] = []
        guard pageCount > 0 else { return nil }
        for part in text.split(omittingEmptySubsequences: false, whereSeparator: { $0 == "," || $0 == ";" }) {
            let bounds = part.split(omittingEmptySubsequences: false, whereSeparator: { $0 == "-" || $0 == "–" }).map { $0.trimmingCharacters(in: .whitespaces) }
            guard (1...2).contains(bounds.count), let first = Int(bounds[0]), let last = Int(bounds[bounds.count - 1]),
                  (1...pageCount).contains(first), (1...pageCount).contains(last) else { return nil }
            pages += first <= last ? Array(first...last) : Array((last...first).reversed())
        }
        return pages.isEmpty ? nil : pages
    }
}

/// The work behind the actions, on files the caller owns, so the journey runs it as Shortcuts does.
enum DocumentActions {
    /// Puts the pages of files together, in order, into a new PDF at output: PDFs, images and documents the engine
    /// converts.
    static func combine(_ inputs: [URL], names: [String], into output: URL) async throws -> URL {
        guard inputs.count >= 2 else { throw IntentProblem.message(String(localized: "Choose at least two files to combine.")) }
        var pdfs: [URL] = []
        defer { for pdf in pdfs where !inputs.contains(pdf) { StudioTemporaryFiles.removeGeneratedPDF(pdf) } }
        for input in inputs { pdfs.append(try await PageSources.pdf(for: input)) }
        let document = try await open(pdfs[0], named: names[0])
        do {
            _ = try await document.importPages(pdfs.dropFirst().map { ImportSource(url: $0) }, before: try await document.facts().pages.count + 1)
        } catch let needed as ImportPasswordNeeded {
            throw IntentProblem.message(protected(names[needed.source + 1]))
        }
        try await document.save(to: output)
        return output
    }

    /// A searchable PDF with a page for each image, its text recognized on this device.
    static func pdf(fromImages images: [Data]) async throws -> URL {
        guard images.count <= ScanBuilder.maximumPages else { throw CocoaError(.fileReadTooLarge) }
        let pages = images.compactMap(ScanBuilder.image(from:))
        guard pages.count == images.count else { throw CocoaError(.fileReadCorruptFile) }
        guard !pages.isEmpty else { throw IntentProblem.message(String(localized: "None of the files is an image OfficeIMO Studio can read.")) }
        return try await ScanBuilder.makePDF(from: pages, named: ScanBuilder.newName)
    }

    /// A new PDF at output from pages such as "1-3, 5", in the order written.
    static func extract(_ pages: String, from input: URL, named name: String, into output: URL) async throws -> URL {
        let document = try await open(input, named: name)
        let count = try await document.facts().pages.count
        guard let numbers = PageRange.parse(pages, pageCount: count) else {
            throw IntentProblem.message(String(localized: "“\(pages)” isn’t a list of pages between 1 and \(count)."))
        }
        try await document.extractPages(numbers, to: output)
        return output
    }

    /// A copy at output with the text of its scanned pages recognized and added; unchanged when there was none to add.
    static func makeSearchable(_ input: URL, named name: String, into output: URL) async throws -> URL {
        let document = try await open(input, named: name)
        guard try await document.facts().canMakeSearchable != false else {
            throw IntentProblem.message(String(localized: "Text can’t be added to “\(name)”."))
        }
        _ = try await document.makeSearchable { _, _ in }
        try await document.save(to: output)
        return output
    }

    /// A PDF at output converted from a Word, Excel, PowerPoint or other document the engine converts.
    static func convert(_ input: URL, named name: String, into output: URL) async throws -> URL {
        do {
            _ = try await Engine.convertToPDF(input, at: output)
        } catch let error as EngineError where error.status == .unsupported {
            throw IntentProblem.message(String(localized: "“\(name)” isn’t a format OfficeIMO Studio converts to PDF."))
        }
        return output
    }

    /// The text of every page in reading order, a blank line between pages.
    static func text(of input: URL, named name: String) async throws -> String {
        let document = try await open(input, named: name)
        let facts = try await document.facts()
        guard facts.canExtractText else { throw IntentProblem.message(String(localized: "“\(name)” doesn’t allow its text to be copied.")) }
        var pages: [String] = []
        for page in facts.pages.map(\.number) {
            try Task.checkCancellation()
            let text = try await document.pageText(page: page)
            pages.append(text.count > 0 ? text.string(0...(text.count - 1)) : "")
        }
        return pages.joined(separator: "\n\n")
    }

    /// Opens a PDF with the engine; a protected one is reported by name rather than asked for.
    private static func open(_ url: URL, named name: String) async throws -> EngineDocument {
        do {
            return try await EngineDocument.open(url)
        } catch let error as EngineError where error.status == .passwordRequired || error.status == .invalidPassword {
            throw IntentProblem.message(protected(name))
        }
    }

    private static func protected(_ name: String) -> String {
        String(localized: "“\(name)” is protected by a password. Open it in OfficeIMO Studio instead.")
    }
}

/// A private folder for one action: copies of its inputs and its output, which the system takes over once returned.
final class ActionFolder {
    let folder: URL
    private var hasReturnedOutput = false

    deinit {
        try? FileManager.default.removeItem(at: hasReturnedOutput ? folder.appending(path: "Inputs") : folder)
    }

    func takeOutput(_ url: URL) -> IntentFile {
        hasReturnedOutput = true
        StudioTemporaryFiles.markReturnedOutput(url)
        return IntentFile(fileURL: url, filename: url.lastPathComponent, type: .pdf)
    }

    /// Starts the engine, as an action may run before any document window, and makes the folder.
    static func prepare() async throws -> ActionFolder {
        try await EngineRuntime.ready()
        StudioTemporaryFiles.removeStale()
        return try ActionFolder()
    }

    private init() throws {
        folder = StudioTemporaryFiles.root.appending(path: "Shortcuts/\(UUID().uuidString)", directoryHint: .isDirectory)
        try FileManager.default.createDirectory(at: folder.appending(path: "Inputs", directoryHint: .isDirectory), withIntermediateDirectories: true)
    }

    func copy(_ file: IntentFile, as name: String) throws -> URL {
        let safe = (name as NSString).lastPathComponent
        guard !safe.isEmpty, safe != ".", safe != ".." else { throw CocoaError(.fileWriteInvalidFileName) }
        let url = folder.appending(path: "Inputs/\(safe)")
        try file.data.write(to: url)
        return url
    }

    func output(named name: String) -> URL {
        let safe = name.replacingOccurrences(of: "/", with: "-").replacingOccurrences(of: ":", with: "-")
        return folder.appending(path: "\(safe.isEmpty ? "Document" : safe).pdf")
    }
}
