// Opt-in macOS validation only. PDFKit is an independent oracle, not an OfficeIMO runtime dependency.
import AppKit
import PDFKit
import Foundation

enum VerificationError: Error { case failed(String) }

func verify(_ condition: Bool, _ message: String) throws {
    if !condition { throw VerificationError.failed(message) }
}

func selectionBounds(_ text: String, page: PDFPage) -> CGRect? {
    guard let content = page.string else { return nil }
    let range = (content as NSString).range(of: text)
    guard range.location != NSNotFound else { return nil }
    return page.selection(for: range)?.bounds(for: page)
}

func rectangle(_ value: CGRect) -> [String: Double] {
    ["x": value.minX, "y": value.minY, "width": value.width, "height": value.height]
}

func png(_ page: PDFPage, destination: URL) throws {
    let image = page.thumbnail(of: NSSize(width: 1000, height: 1000), for: .mediaBox)
    guard let tiff = image.tiffRepresentation,
          let bitmap = NSBitmapImageRep(data: tiff),
          let data = bitmap.representation(using: .png, properties: [:]) else {
        throw VerificationError.failed("PDFKit could not render a page.")
    }
    try data.write(to: destination)
}

do {
    let args = Array(CommandLine.arguments.dropFirst())
    try verify(args.count >= 5, "Usage: swift Verify-PdfRedactionPdfKit.swift source.pdf redacted.pdf output-directory removed-regex retained-text [retained-text ...]")
    guard let source = PDFDocument(url: URL(fileURLWithPath: args[0])),
          let output = PDFDocument(url: URL(fileURLWithPath: args[1])) else {
        throw VerificationError.failed("PDFKit could not open both PDFs.")
    }
    try verify(source.pageCount > 0 && source.pageCount == output.pageCount, "Page count changed.")
    let expression = try NSRegularExpression(pattern: args[3])
    let beforeText = source.string ?? ""
    let afterText = output.string ?? ""
    try verify(expression.numberOfMatches(in: beforeText, range: NSRange(location: 0, length: (beforeText as NSString).length)) > 0,
               "Removal criterion did not match the source in PDFKit.")
    try verify(expression.numberOfMatches(in: afterText, range: NSRange(location: 0, length: (afterText as NSString).length)) == 0,
               "Selected text is still extractable in PDFKit.")
    let directory = URL(fileURLWithPath: args[2], isDirectory: true)
    try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
    var pages: [[String: Any]] = []
    var retainedCounts = Dictionary(uniqueKeysWithValues: args.dropFirst(4).map { ($0, 0) })
    for index in 0..<source.pageCount {
        guard let before = source.page(at: index), let after = output.page(at: index) else {
            throw VerificationError.failed("A page is missing.")
        }
        try verify(before.rotation == after.rotation, "Page rotation changed.")
        try verify(before.bounds(for: .mediaBox) == after.bounds(for: .mediaBox), "MediaBox changed.")
        var retained: [[String: Any]] = []
        for marker in args.dropFirst(4) {
            guard let original = selectionBounds(marker, page: before) else { continue }
            guard let saved = selectionBounds(marker, page: after) else {
                throw VerificationError.failed("Neighboring text was removed: \(marker)")
            }
            let differences = [abs(original.minX - saved.minX), abs(original.minY - saved.minY),
                               abs(original.width - saved.width), abs(original.height - saved.height)]
            try verify(differences.allSatisfy { $0 < 0.05 }, "Neighboring text moved: \(marker)")
            retainedCounts[marker, default: 0] += 1
            retained.append(["text": marker, "source": rectangle(original), "output": rectangle(saved)])
        }
        try png(before, destination: directory.appendingPathComponent("source-\(index + 1).png"))
        try png(after, destination: directory.appendingPathComponent("redacted-\(index + 1).png"))
        pages.append(["page": index + 1, "rotation": after.rotation, "mediaBox": rectangle(after.bounds(for: .mediaBox)),
                      "textCharacters": (after.string ?? "").count, "retained": retained])
    }
    try verify(retainedCounts.values.allSatisfy { $0 > 0 }, "A required retained marker was absent from the source.")
    let report: [String: Any] = ["reader": "Apple PDFKit", "os": ProcessInfo.processInfo.operatingSystemVersionString,
                               "pageCount": output.pageCount, "removedMatches": 0,
                               "retainedMarkerPageCounts": retainedCounts, "pages": pages]
    let data = try JSONSerialization.data(withJSONObject: report, options: [.prettyPrinted, .sortedKeys])
    try data.write(to: directory.appendingPathComponent("pdfkit-evidence.json"))
    print("PDFKit verified \(output.pageCount) pages, removed text and unchanged neighboring selection bounds.")
} catch {
    fputs("PDFKit verification failed: \(error)\n", stderr)
    exit(1)
}
