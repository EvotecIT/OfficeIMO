// Opt-in macOS validation only. PDFKit is an independent oracle, not an OfficeIMO runtime dependency.
import AppKit
import PDFKit
import Foundation

enum VerificationError: Error { case failed(String) }

func verify(_ condition: Bool, _ message: String) throws {
    if !condition { throw VerificationError.failed(message) }
}

func selectionBounds(_ text: String, page: PDFPage) throws -> [CGRect] {
    guard let content = page.string else { return [] }
    let value = content as NSString
    var location = 0
    var bounds: [CGRect] = []
    while location < value.length {
        let range = value.range(of: text, options: [], range: NSRange(location: location, length: value.length - location))
        if range.location == NSNotFound { break }
        guard let selection = page.selection(for: range) else {
            throw VerificationError.failed("PDFKit could not locate retained text: \(text)")
        }
        bounds.append(selection.bounds(for: page))
        location = NSMaxRange(range)
    }
    return bounds
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
    let markers = Array(Set(args.dropFirst(4))).sorted()
    try verify(markers.allSatisfy { !$0.isEmpty }, "Retained markers must not be empty.")
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
    var retainedCounts = Dictionary(uniqueKeysWithValues: markers.map { ($0, 0) })
    var retainedOccurrenceCounts = retainedCounts
    for index in 0..<source.pageCount {
        guard let before = source.page(at: index), let after = output.page(at: index) else {
            throw VerificationError.failed("A page is missing.")
        }
        try verify(before.rotation == after.rotation, "Page rotation changed.")
        try verify(before.bounds(for: .mediaBox) == after.bounds(for: .mediaBox), "MediaBox changed.")
        var retained: [[String: Any]] = []
        for marker in markers {
            let originals = try selectionBounds(marker, page: before)
            let saved = try selectionBounds(marker, page: after)
            try verify(originals.count == saved.count, "Neighboring text occurrence count changed on page \(index + 1): \(marker)")
            if originals.isEmpty { continue }
            for (occurrence, pair) in zip(originals, saved).enumerated() {
                let differences = [abs(pair.0.minX - pair.1.minX), abs(pair.0.minY - pair.1.minY),
                                   abs(pair.0.width - pair.1.width), abs(pair.0.height - pair.1.height)]
                try verify(differences.allSatisfy { $0 < 0.05 }, "Neighboring text moved on page \(index + 1), occurrence \(occurrence + 1): \(marker)")
                retained.append(["text": marker, "occurrence": occurrence + 1,
                                 "source": rectangle(pair.0), "output": rectangle(pair.1)])
            }
            retainedCounts[marker, default: 0] += 1
            retainedOccurrenceCounts[marker, default: 0] += originals.count
        }
        try png(before, destination: directory.appendingPathComponent("source-\(index + 1).png"))
        try png(after, destination: directory.appendingPathComponent("redacted-\(index + 1).png"))
        pages.append(["page": index + 1, "rotation": after.rotation, "mediaBox": rectangle(after.bounds(for: .mediaBox)),
                      "textCharacters": (after.string ?? "").count, "retained": retained])
    }
    try verify(retainedCounts.values.allSatisfy { $0 > 0 }, "A required retained marker was absent from the source.")
    let report: [String: Any] = ["reader": "Apple PDFKit", "os": ProcessInfo.processInfo.operatingSystemVersionString,
                               "pageCount": output.pageCount, "removedMatches": 0,
                               "retainedMarkerPageCounts": retainedCounts,
                               "retainedMarkerOccurrenceCounts": retainedOccurrenceCounts, "pages": pages]
    let data = try JSONSerialization.data(withJSONObject: report, options: [.prettyPrinted, .sortedKeys])
    try data.write(to: directory.appendingPathComponent("pdfkit-evidence.json"))
    print("PDFKit verified \(output.pageCount) pages, removed text and unchanged neighboring selection bounds.")
} catch {
    fputs("PDFKit verification failed: \(error)\n", stderr)
    exit(1)
}
