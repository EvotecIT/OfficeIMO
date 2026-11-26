import Foundation

/// The documents OfficeIMO Studio opens to view: Word, Excel, PowerPoint and the other formats the engine converts to
/// PDF in process, among them web pages, XPS, Visio, Publisher, OpenDocument drawings and EPUB books. Pages, Numbers and
/// Keynote files aren't listed: the engine converts them only through the separate iWork workflow.
/// They show as converted pages, read-only, and become a PDF only when the person chooses Edit as PDF or Save as PDF.
/// Both Info.plists declare these types with the Viewer role, so Files and Finder offer Studio for them.
enum ViewerFormats {
    struct Format: Sendable {
        let identifier: String
        let extensions: [String]
    }

    static let formats: [Format] = [
        Format(identifier: "org.openxmlformats.wordprocessingml.document", extensions: ["docx"]),
        Format(identifier: "org.openxmlformats.wordprocessingml.document.macroenabled", extensions: ["docm"]),
        Format(identifier: "org.openxmlformats.wordprocessingml.template", extensions: ["dotx"]),
        Format(identifier: "org.openxmlformats.wordprocessingml.template.macroenabled", extensions: ["dotm"]),
        Format(identifier: "com.microsoft.word.doc", extensions: ["doc"]),
        Format(identifier: "org.openxmlformats.spreadsheetml.sheet", extensions: ["xlsx"]),
        Format(identifier: "org.openxmlformats.spreadsheetml.sheet.macroenabled", extensions: ["xlsm"]),
        Format(identifier: "org.openxmlformats.spreadsheetml.template", extensions: ["xltx"]),
        Format(identifier: "org.openxmlformats.spreadsheetml.template.macroenabled", extensions: ["xltm"]),
        Format(identifier: "org.openxmlformats.presentationml.presentation", extensions: ["pptx"]),
        Format(identifier: "org.openxmlformats.presentationml.presentation.macroenabled", extensions: ["pptm"]),
        Format(identifier: "org.openxmlformats.presentationml.template", extensions: ["potx"]),
        Format(identifier: "org.openxmlformats.presentationml.template.macroenabled", extensions: ["potm"]),
        Format(identifier: "org.openxmlformats.presentationml.slideshow", extensions: ["ppsx"]),
        Format(identifier: "org.openxmlformats.presentationml.slideshow.macroenabled", extensions: ["ppsm"]),
        Format(identifier: "public.rtf", extensions: ["rtf"]),
        Format(identifier: "net.daringfireball.markdown", extensions: ["md", "markdown"]),
        Format(identifier: "org.oasis-open.opendocument.text", extensions: ["odt"]),
        Format(identifier: "org.oasis-open.opendocument.spreadsheet", extensions: ["ods"]),
        Format(identifier: "org.oasis-open.opendocument.presentation", extensions: ["odp"]),
        Format(identifier: "public.html", extensions: ["html", "htm"]),
        Format(identifier: "com.microsoft.xps", extensions: ["xps", "oxps"]),
        Format(identifier: "com.microsoft.visio.drawing", extensions: ["vsdx", "vdx", "vtx"]),
        Format(identifier: "com.evotec.officeimo.cbz", extensions: ["cbz"]),
        Format(identifier: "com.microsoft.publisher.pub", extensions: ["pub"]),
        Format(identifier: "com.evotec.officeimo.mobi", extensions: ["mobi", "prc", "pdb", "azw", "azw3"]),
        Format(identifier: "org.oasis-open.opendocument.graphics", extensions: ["odg"]),
        Format(identifier: "org.idpf.epub-container", extensions: ["epub"]),
    ]

    static var identifiers: [String] { formats.map(\.identifier) }

    /// Types that not every system release declares, so the Info.plists import them with these extensions.
    static let importedIdentifiers: Set<String> = ["net.daringfireball.markdown", "org.oasis-open.opendocument.text",
                                                   "org.oasis-open.opendocument.spreadsheet",
                                                   "org.oasis-open.opendocument.presentation", "com.microsoft.xps",
                                                   "com.evotec.officeimo.mobi",
                                                   "com.microsoft.visio.drawing", "com.microsoft.publisher.pub",
                                                   "com.evotec.officeimo.cbz",
                                                   "org.oasis-open.opendocument.graphics"]

    /// A larger document can take minutes to convert on a phone and exhaust its memory, so it is refused before
    /// conversion starts.
    static let maximumBytes: Int64 = 100 * 1024 * 1024

    /// Why a document of this size can't be shown, or nil when it can.
    static func sizeProblem(byteCount: Int64, name: String) -> String? {
        if byteCount == 0 { return String(localized: "“\(name)” is empty.") }
        guard byteCount > maximumBytes else { return nil }
        return String(localized: "“\(name)” is larger than 100 MB, too large for OfficeIMO Studio to show.")
    }
}
