import CoreGraphics
import Foundation
import ImageIO

#if os(macOS)
import AppKit
#else
import UIKit
#endif

enum PageImage {
    /// Decodes the engine's PNG fully, so timing includes the decode and display does not decode again.
    static func decode(_ png: Data) throws -> CGImage {
        let options = [kCGImageSourceShouldCacheImmediately: true] as CFDictionary
        guard let source = CGImageSourceCreateWithData(png as CFData, nil),
              let image = CGImageSourceCreateImageAtIndex(source, 0, options) else {
            throw EngineError(status: nil, code: -1, message: "The page image could not be decoded.")
        }
        return image
    }
}

/// The display the measurements render for: pixels per point and the width a page fills, in points.
struct DisplayTarget: Codable, Sendable {
    let scale: Double
    let widthPoints: Double

    /// The pixel width of a page that fills the target width at the display's pixel density.
    var pixelWidth: Double { scale * widthPoints }

    @MainActor static var current: DisplayTarget {
        #if os(macOS)
        // A document window on a Mac rather than the whole display.
        DisplayTarget(scale: Double(NSScreen.main?.backingScaleFactor ?? 2), widthPoints: 900)
        #else
        let screen = UIScreen.main
        return DisplayTarget(scale: Double(screen.scale), widthPoints: Double(min(screen.bounds.width, screen.bounds.height)))
        #endif
    }
}
