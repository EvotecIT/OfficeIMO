import CoreGraphics
import Foundation
import ImageIO
import UniformTypeIdentifiers
import Vision

/// A scan for officeimo_create_pdf_from_scan_json: page images and the words recognized on each.
struct ScanRequest: Encodable, Sendable {
    struct Word: Encodable, Sendable {
        let text: String
        let x: Double
        let y: Double
        let width: Double
        let height: Double
        let confidence: Double
    }

    struct Page: Encodable, Sendable {
        let image: String
        let words: [Word]
    }

    let pages: [Page]
}

struct ScanSummary: Decodable, Sendable {
    let pages: Int
    let searchablePages: Int
    let words: Int
}

/// Makes a searchable PDF from page images: Apple's Vision framework recognizes the text on the device, and the
/// engine composes the PDF with the words as invisible text over each image, so the scan can be searched and copied.
enum ScanBuilder {
    /// The longest side a page image keeps, which holds 300 dpi for a letter page and keeps photos to a sensible size.
    static let maximumPixels = 3300

    static func makePDF(from images: [CGImage], named name: String) async throws -> URL {
        guard !images.isEmpty else { throw CocoaError(.userCancelled) }
        let folder = FileManager.default.temporaryDirectory.appending(path: "Scans/\(UUID().uuidString)", directoryHint: .isDirectory)
        try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
        var pages: [ScanRequest.Page] = []
        for (index, image) in images.enumerated() {
            try Task.checkCancellation()
            let page = scaled(image)
            let url = folder.appending(path: "page-\(index + 1).jpg")
            try writeJPEG(page, to: url)
            pages.append(ScanRequest.Page(image: url.path, words: await recognizeWords(in: page)))
        }
        let output = folder.appending(path: "\(name).pdf")
        _ = try await Engine.createPDFFromScan(ScanRequest(pages: pages), at: output)
        for page in pages { try? FileManager.default.removeItem(atPath: page.image) }
        return output
    }

    /// The words on a page, each with its box from the image's top-left corner in fractions of its size.
    static func recognizeWords(in image: CGImage) async -> [ScanRequest.Word] {
        var request = RecognizeTextRequest()
        request.recognitionLevel = .accurate
        request.usesLanguageCorrection = true
        guard let observations = try? await request.perform(on: image) else { return [] }
        var words: [ScanRequest.Word] = []
        for observation in observations {
            guard let candidate = observation.topCandidates(1).first else { continue }
            let text = candidate.string
            var index = text.startIndex
            // Words are runs between spaces, so punctuation stays with its word as it reads.
            while index < text.endIndex {
                guard let start = text[index...].firstIndex(where: { !$0.isWhitespace }) else { break }
                let end = text[start...].firstIndex(where: \.isWhitespace) ?? text.endIndex
                if let box = candidate.boundingBox(for: start..<end)?.boundingBox.cgRect, box.width > 0, box.height > 0 {
                    // Vision measures from the bottom left; the engine from the top left.
                    let x = min(max(box.minX, 0), 1)
                    let y = min(max(1 - box.maxY, 0), 1)
                    words.append(ScanRequest.Word(text: String(text[start..<end]), x: x, y: y,
                                                  width: min(box.width, 1 - x), height: min(box.height, 1 - y),
                                                  confidence: Double(candidate.confidence)))
                }
                index = end
            }
        }
        return words
    }

    private static func scaled(_ image: CGImage) -> CGImage {
        let longest = max(image.width, image.height)
        guard longest > maximumPixels else { return image }
        let factor = Double(maximumPixels) / Double(longest)
        let width = Int(Double(image.width) * factor), height = Int(Double(image.height) * factor)
        guard let context = CGContext(data: nil, width: width, height: height, bitsPerComponent: 8, bytesPerRow: 0,
                                      space: CGColorSpace(name: CGColorSpace.sRGB)!,
                                      bitmapInfo: CGImageAlphaInfo.noneSkipLast.rawValue) else { return image }
        context.interpolationQuality = .high
        context.draw(image, in: CGRect(x: 0, y: 0, width: width, height: height))
        return context.makeImage() ?? image
    }

    private static func writeJPEG(_ image: CGImage, to url: URL) throws {
        guard let destination = CGImageDestinationCreateWithURL(url as CFURL, UTType.jpeg.identifier as CFString, 1, nil) else {
            throw CocoaError(.fileWriteUnknown)
        }
        CGImageDestinationAddImage(destination, image, [kCGImageDestinationLossyCompressionQuality: 0.82] as CFDictionary)
        guard CGImageDestinationFinalize(destination) else { throw CocoaError(.fileWriteUnknown) }
    }

    /// A document name for a new scan, as the system names new documents.
    static var newName: String {
        String(localized: "Scan \(Date.now.formatted(date: .abbreviated, time: .shortened))").replacingOccurrences(of: "/", with: "-")
            .replacingOccurrences(of: ":", with: ".")
    }
}
