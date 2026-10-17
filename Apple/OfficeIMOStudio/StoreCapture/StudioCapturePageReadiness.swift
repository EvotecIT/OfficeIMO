import CoreGraphics
import Foundation
import Vision

/// Capture-only recognition of the public fixture, followed by deterministic
/// page-region pixels. Chrome, the clock and the search caret are outside it.
enum StudioCapturePageReadiness {
    enum Failure: Error { case missingDashboard, unreadablePixels }

    static func pagePixels(from image: CGImage) throws -> Data {
        let request = VNRecognizeTextRequest()
        request.recognitionLevel = .accurate
        request.recognitionLanguages = ["en-US"]
        try VNImageRequestHandler(cgImage: image).perform([request])
        let lines = (request.results ?? []).compactMap { observation -> (String, CGRect)? in
            guard let text = observation.topCandidates(1).first?.string else { return nil }
            return (text.lowercased(), observation.boundingBox)
        }
        // The three phrases span the heading, chart and lower decisions table.
        // A title bar, thumbnail, blank page or different sample is insufficient.
        let phrases = ["quarterly operations dashboard", "tickets resolved per month", "review the 37 open items"]
        let regions = try phrases.map { phrase in
            guard let line = lines.first(where: { $0.0.contains(phrase) }) else { throw Failure.missingDashboard }
            return line.1
        }
        let normalized = regions.dropFirst().reduce(regions[0]) { $0.union($1) }
        let bounds = CGRect(x: 0, y: 0, width: image.width, height: image.height)
        let page = CGRect(x: normalized.minX * bounds.width, y: (1 - normalized.maxY) * bounds.height,
                          width: normalized.width * bounds.width, height: normalized.height * bounds.height)
            .integral.intersection(bounds)
        guard let crop = image.cropping(to: page) else { throw Failure.unreadablePixels }
        var pixels = [UInt8](repeating: 0, count: crop.width * crop.height * 4)
        let drawn = pixels.withUnsafeMutableBytes { bytes in
            guard let context = CGContext(data: bytes.baseAddress, width: crop.width, height: crop.height,
                                          bitsPerComponent: 8, bytesPerRow: crop.width * 4,
                                          space: CGColorSpaceCreateDeviceRGB(),
                                          bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else { return false }
            context.draw(crop, in: CGRect(x: 0, y: 0, width: crop.width, height: crop.height))
            return true
        }
        guard drawn else { throw Failure.unreadablePixels }
        return Data(pixels)
    }
}
