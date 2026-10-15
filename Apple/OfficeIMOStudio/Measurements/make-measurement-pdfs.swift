// Builds the two PDFs the engine feasibility measurements use, from fixtures in this repository:
//   long.pdf     600 pages cycled from text, table, form and presentation PDFs produced by different applications
//   scanned.pdf  40 pages of 300 dpi colour scans stored as JPEG, the way scanners and phone scan apps write them
//
// Usage, from the repository root: swift Apple/OfficeIMOStudio/Measurements/make-measurement-pdfs.swift <output-folder>
import CoreGraphics
import Foundation
import ImageIO
import PDFKit
import UniformTypeIdentifiers

let arguments = CommandLine.arguments
guard arguments.count == 2 else {
    FileHandle.standardError.write(Data("Usage: make-measurement-pdfs.swift <output-folder>\n".utf8))
    exit(2)
}
let output = URL(fileURLWithPath: arguments[1], isDirectory: true)
try FileManager.default.createDirectory(at: output, withIntermediateDirectories: true)
let root = URL(fileURLWithPath: FileManager.default.currentDirectoryPath)

let textSources = [
    "OfficeIMO.Pdf.Tests/Pdf/ReferenceBaselines/microsoft-word-windows-word-business-delivery-summary.pdf",
    "Website/static/downloads/showcase/features/pdf-table-styles/pdf.tablestylegallery.pdf",
    "OfficeIMO.TestAssets/OcrQuality/fw9.pdf",
    "Website/static/downloads/showcase/features/word-comparison/comparison-report.pdf",
    "OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-v15.4.pdf",
    "OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-web-capture.pdf",
]
let scanSources = [
    "OfficeIMO.TestAssets/OcrQuality/8087.tif",
    "OfficeIMO.TestAssets/MultilingualLayout/latin-0-scan.png",
    "OfficeIMO.TestAssets/MultilingualLayout/rtl-0-scan.png",
    "OfficeIMO.TestAssets/OcrQuality/8071.tif",
]

func makeLong(pageCount: Int) throws {
    let sources = textSources.map { path -> PDFDocument in
        guard let document = PDFDocument(url: root.appending(path: path)) else { fatalError("Cannot read \(path)") }
        return document
    }
    let long = PDFDocument()
    var index = 0
    while long.pageCount < pageCount {
        let source = sources[index % sources.count]
        for number in 0..<source.pageCount where long.pageCount < pageCount {
            long.insert(source.page(at: number)!.copy() as! PDFPage, at: long.pageCount)
        }
        index += 1
    }
    guard long.write(to: output.appending(path: "long.pdf")) else { fatalError("Cannot write long.pdf") }
}

func makeScanned(pageCount: Int) throws {
    let url = output.appending(path: "scanned.pdf")
    guard let context = CGContext(url as CFURL, mediaBox: nil, nil) else { fatalError("Cannot create scanned.pdf") }
    let scans = scanSources.map { path -> CGImage in
        guard let source = CGImageSourceCreateWithURL(root.appending(path: path) as CFURL, nil),
              let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else { fatalError("Cannot read \(path)") }
        return image
    }
    for number in 0..<pageCount {
        let scan = scans[number % scans.count]
        // Every page gets its own image, like a real scan: the source is stamped with a page mark, encoded once as
        // JPEG, and the PDF context keeps JPEG data as a DCT stream.
        let bitmap = CGContext(data: nil, width: scan.width, height: scan.height, bitsPerComponent: 8, bytesPerRow: 0,
                               space: CGColorSpace(name: CGColorSpace.sRGB)!, bitmapInfo: CGImageAlphaInfo.noneSkipLast.rawValue)!
        bitmap.draw(scan, in: CGRect(x: 0, y: 0, width: scan.width, height: scan.height))
        bitmap.setFillColor(gray: 0, alpha: 1)
        for bit in 0..<8 where (number >> bit) & 1 == 1 {
            bitmap.fill(CGRect(x: 40 + bit * 30, y: 40, width: 20, height: 20))
        }
        let jpeg = NSMutableData()
        let destination = CGImageDestinationCreateWithData(jpeg, UTType.jpeg.identifier as CFString, 1, nil)!
        CGImageDestinationAddImage(destination, bitmap.makeImage()!, [kCGImageDestinationLossyCompressionQuality: 0.75] as CFDictionary)
        CGImageDestinationFinalize(destination)
        let image = CGImage(jpegDataProviderSource: CGDataProvider(data: jpeg)!, decode: nil, shouldInterpolate: true, intent: .defaultIntent)!
        // 300 dots per inch.
        var box = CGRect(x: 0, y: 0, width: Double(scan.width) * 72 / 300, height: Double(scan.height) * 72 / 300)
        context.beginPage(mediaBox: &box)
        context.draw(image, in: box)
        context.endPage()
    }
    context.closePDF()
}

try makeLong(pageCount: 600)
try makeScanned(pageCount: 40)
for name in ["long.pdf", "scanned.pdf"] {
    let url = output.appending(path: name)
    let size = (try FileManager.default.attributesOfItem(atPath: url.path)[.size] as? NSNumber)?.intValue ?? 0
    print("\(name): \(PDFDocument(url: url)?.pageCount ?? 0) pages, \(size / 1024) KB")
}
