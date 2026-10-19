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
    /// Matches the shared scan composer's admission limit, before image preparation starts.
    static let maximumPages = 500

    @MainActor
    static func makePDF(pageCount: Int, named name: String,
                        loadPage: @MainActor (Int) async throws -> CGImage) async throws -> URL {
        guard pageCount > 0 else { throw CocoaError(.userCancelled) }
        guard pageCount <= maximumPages else { throw CocoaError(.fileReadTooLarge) }
        let name = try StudioTemporaryFiles.outputName(name)
        let folder = StudioTemporaryFiles.root.appending(path: "Scans/\(UUID().uuidString)", directoryHint: .isDirectory)
        try FileManager.default.createDirectory(at: folder, withIntermediateDirectories: true)
        var completed = false
        defer { if !completed { try? FileManager.default.removeItem(at: folder) } }
        var pages: [ScanRequest.Page] = []
        for index in 0..<pageCount {
            try Task.checkCancellation()
            // This frame returns only disk paths and words; no source image survives into the next load.
            pages.append(try await preparePage(index, in: folder, loadPage: loadPage))
        }
        try Task.checkCancellation()
        let output = folder.appending(path: "\(name).pdf")
        _ = try await Engine.createPDFFromScan(ScanRequest(pages: pages), at: output)
        try Task.checkCancellation()
        for page in pages { try? FileManager.default.removeItem(atPath: page.image) }
        completed = true
        return output
    }

    @MainActor
    private static func preparePage(_ index: Int, in folder: URL,
                                    loadPage: @MainActor (Int) async throws -> CGImage) async throws -> ScanRequest.Page {
        var image: CGImage? = try await loadPage(index)
        try Task.checkCancellation()
        let url = folder.appending(path: "page-\(index + 1).jpg")
        try autoreleasepool { try writeJPEG(scaled(image!), to: url) }
        image = nil
        // Vision consumes the staged page rather than retaining the caller's full-resolution photo across the await.
        let request = textRequest()
        return ScanRequest.Page(image: url.path, words: try words(from: await request.perform(on: url)))
    }

    /// Providers own their temporary URL only during the callback. Decode a bounded, upright image before returning.
    @MainActor
    static func image(from provider: NSItemProvider) async throws -> CGImage {
        try await withCheckedThrowingContinuation { continuation in
            _ = provider.loadFileRepresentation(forTypeIdentifier: UTType.image.identifier) { url, error in
                if let error { continuation.resume(throwing: error) }
                else if let url, let image = autoreleasepool(invoking: { image(at: url) }) {
                    continuation.resume(returning: image)
                } else { continuation.resume(throwing: CocoaError(.fileReadCorruptFile)) }
            }
        }
    }

    /// The words on a page, each with its box from the image's top-left corner in fractions of its size.
    static func recognizeWords(in image: CGImage) async throws -> [ScanRequest.Word] {
        try Task.checkCancellation()
        let request = textRequest()
        let observations = try await request.perform(on: image)
        return try words(from: observations)
    }

    private static func textRequest() -> RecognizeTextRequest {
        var request = RecognizeTextRequest()
        request.recognitionLevel = .accurate
        request.usesLanguageCorrection = true
        return request
    }

    private static func words(from observations: [RecognizedTextObservation]) throws -> [ScanRequest.Word] {
        try Task.checkCancellation()
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

    private static func scaled(_ image: CGImage) throws -> CGImage {
        let longest = max(image.width, image.height)
        guard longest > maximumPixels else { return image }
        let factor = Double(maximumPixels) / Double(longest)
        let width = Int(Double(image.width) * factor), height = Int(Double(image.height) * factor)
        guard let context = CGContext(data: nil, width: width, height: height, bitsPerComponent: 8, bytesPerRow: 0,
                                      space: CGColorSpace(name: CGColorSpace.sRGB)!,
                                      bitmapInfo: CGImageAlphaInfo.noneSkipLast.rawValue) else { throw CocoaError(.fileReadTooLarge) }
        context.interpolationQuality = .high
        context.draw(image, in: CGRect(x: 0, y: 0, width: width, height: height))
        guard let result = context.makeImage() else { throw CocoaError(.fileReadTooLarge) }
        return result
    }

    private static func writeJPEG(_ image: CGImage, to url: URL) throws {
        guard let destination = CGImageDestinationCreateWithURL(url as CFURL, UTType.jpeg.identifier as CFString, 1, nil) else {
            throw CocoaError(.fileWriteUnknown)
        }
        CGImageDestinationAddImage(destination, image, [kCGImageDestinationLossyCompressionQuality: 0.82,
                                                      kCGImagePropertyDPIWidth: 300, kCGImagePropertyDPIHeight: 300] as CFDictionary)
        guard CGImageDestinationFinalize(destination) else { throw CocoaError(.fileWriteUnknown) }
    }

    /// A document name for a new scan, as the system names new documents.
    /// An image file, upright and at most the longest side a page keeps; nil when it is not an image.
    static func image(at url: URL) -> CGImage? {
        guard let source = CGImageSourceCreateWithURL(url as CFURL, nil) else { return nil }
        return image(from: source)
    }

    static func image(from data: Data) -> CGImage? {
        guard let source = CGImageSourceCreateWithData(data as CFData, nil) else { return nil }
        return image(from: source)
    }

    private static func image(from source: CGImageSource) -> CGImage? {
        // The thumbnail path applies the photo's orientation, so the page is upright.
        let options: [CFString: Any] = [kCGImageSourceCreateThumbnailFromImageAlways: true, kCGImageSourceCreateThumbnailWithTransform: true,
                                        kCGImageSourceThumbnailMaxPixelSize: maximumPixels, kCGImageSourceShouldCacheImmediately: true]
        return CGImageSourceCreateThumbnailAtIndex(source, 0, options as CFDictionary)
    }

    static var newName: String {
        String(localized: "Scan \(Date.now.formatted(date: .abbreviated, time: .shortened))").replacingOccurrences(of: "/", with: "-")
            .replacingOccurrences(of: ":", with: ".")
    }
}
