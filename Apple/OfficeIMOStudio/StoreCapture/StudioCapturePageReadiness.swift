import CoreGraphics
import Foundation
import Vision

/// Capture-only recognition of the public fixture and deterministic reading-area
/// pixels. The caller supplies the visible reading area from native accessibility.
enum StudioCapturePageReadiness {
    enum Failure: Error {
        case missingDashboard(phrase: String, recognized: [String])
        case unreadablePixels
    }

    #if os(macOS)
    /// Window screenshots can include system alerts drawn above the app. Check
    /// front-to-back window metadata without reading unrelated window titles.
    static func isWindowUnoccluded(processIdentifier: Int32, frame: CGRect,
                                   windows: [[String: Any]]) -> Bool {
        func bounds(_ window: [String: Any]) -> CGRect? {
            guard let dictionary = window[kCGWindowBounds as String] as? [String: Any] else { return nil }
            return CGRect(dictionaryRepresentation: dictionary as CFDictionary)
        }
        func visible(_ window: [String: Any]) -> Bool {
            (window[kCGWindowAlpha as String] as? NSNumber)?.doubleValue != 0
        }
        guard processIdentifier > 0, !frame.isEmpty,
              let target = windows.firstIndex(where: {
                  ($0[kCGWindowOwnerPID as String] as? NSNumber)?.int32Value == processIdentifier &&
                      bounds($0)?.integral == frame.integral && visible($0)
              }) else { return false }
        for window in windows.prefix(target) where visible(window) {
            guard let rectangle = bounds(window) else { return false }
            if rectangle.intersects(frame) { return false }
        }
        return true
    }
    #endif

    static func pagePixels(from source: CGImage, contentBounds: CGRect? = nil) throws -> Data {
        let bounds = CGRect(x: 0, y: 0, width: source.width, height: source.height)
        let visible = (contentBounds ?? bounds).integral.intersection(bounds)
        guard !visible.isNull, !visible.isEmpty,
              let image = source.cropping(to: visible) else { throw Failure.unreadablePixels }
        var lines = try recognize(image)
        // The three phrases span the heading, chart and lower decisions table.
        // A title bar, thumbnail, blank page or different sample is insufficient.
        let phrases = ["quarterly operations dashboard", "tickets resolved per month", "review the 37 open items"]
        func missingPhrase() -> String? {
            phrases.first { phrase in !lines.contains { $0.contains(phrase) } }
        }
        // Overlapping crops also isolate small table text in compact views.
        // Every phrase must still occur in this same unmodified screenshot.
        if missingPhrase() != nil {
            let width = (image.width * 2 + 2) / 3
            let height = (image.height * 2 + 4) / 5
            tiles: for y in [0, (image.height - height) / 2, image.height - height] {
                for x in [0, image.width - width] {
                    guard let tile = image.cropping(to: CGRect(x: x, y: y, width: width, height: height)) else {
                        throw Failure.unreadablePixels
                    }
                    lines += try recognize(tile)
                    if missingPhrase() == nil { break tiles }
                }
            }
        }
        if let phrase = missingPhrase() {
            throw Failure.missingDashboard(phrase: phrase, recognized: Array(lines.prefix(20)))
        }
        // Compare a fixed native reading area, rather than OCR bounding boxes
        // that may shift between requests even when the rendered pixels do not.
        var pixels = [UInt8](repeating: 0, count: image.width * image.height * 4)
        let drawn = pixels.withUnsafeMutableBytes { bytes in
            guard let context = CGContext(data: bytes.baseAddress, width: image.width, height: image.height,
                                          bitsPerComponent: 8, bytesPerRow: image.width * 4,
                                          space: CGColorSpaceCreateDeviceRGB(),
                                          bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else { return false }
            context.draw(image, in: CGRect(x: 0, y: 0, width: image.width, height: image.height))
            return true
        }
        guard drawn else { throw Failure.unreadablePixels }
        return Data(pixels)
    }

    private static func recognize(_ image: CGImage) throws -> [String] {
        // Cropped CGImages can retain a larger parent's stride and an offset.
        // Give Vision its own contiguous raster instead of that provider view.
        guard let context = CGContext(data: nil, width: image.width, height: image.height,
                                      bitsPerComponent: 8, bytesPerRow: image.width * 4,
                                      space: CGColorSpaceCreateDeviceRGB(),
                                      bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else {
            throw Failure.unreadablePixels
        }
        context.draw(image, in: CGRect(x: 0, y: 0, width: image.width, height: image.height))
        guard let raster = context.makeImage() else { throw Failure.unreadablePixels }
        let request = VNRecognizeTextRequest()
        request.recognitionLevel = .accurate
        request.recognitionLanguages = ["en-US"]
        try VNImageRequestHandler(cgImage: raster).perform([request])
        return (request.results ?? []).compactMap { $0.topCandidates(1).first?.string.lowercased() }
    }
}
