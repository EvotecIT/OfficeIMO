#!/usr/bin/env swift

import AppKit
import Foundation

private func fail(_ message: String) -> Never {
    FileHandle.standardError.write(Data((message + "\n").utf8))
    exit(1)
}

guard CommandLine.arguments.count == 3 else {
    fail("Usage: screenshot-contact-sheet.swift <input-directory> <output.png>")
}

let fileManager = FileManager.default
let inputRoot = URL(fileURLWithPath: CommandLine.arguments[1], isDirectory: true).standardizedFileURL
let outputURL = URL(fileURLWithPath: CommandLine.arguments[2]).standardizedFileURL
var imageURLs: [URL] = []

guard let enumerator = fileManager.enumerator(
    at: inputRoot,
    includingPropertiesForKeys: [.isRegularFileKey, .isSymbolicLinkKey],
    options: [.skipsHiddenFiles]
) else {
    fail("Could not enumerate screenshot candidate directory: \(inputRoot.path)")
}

for case let url as URL in enumerator {
    let values = try? url.resourceValues(forKeys: [.isRegularFileKey, .isSymbolicLinkKey])
    guard values?.isSymbolicLink != true else {
        fail("Screenshot candidate directory must not contain symbolic links: \(url.path)")
    }
    guard url.pathExtension.lowercased() == "png" else { continue }
    guard values?.isRegularFile == true else {
        fail("Screenshot candidate must be an unlinked PNG file: \(url.path)")
    }
    imageURLs.append(url)
}

imageURLs.sort {
    let left = ($0.lastPathComponent, $0.deletingLastPathComponent().lastPathComponent)
    let right = ($1.lastPathComponent, $1.deletingLastPathComponent().lastPathComponent)
    return left < right
}
guard !imageURLs.isEmpty else {
    fail("No PNG screenshots were found under \(inputRoot.path)")
}

let images: [(url: URL, image: NSImage, width: CGFloat, height: CGFloat)] = imageURLs.map { url in
    guard let image = NSImage(contentsOf: url) else {
        fail("Could not read screenshot candidate: \(url.path)")
    }
    let pixelWidth = image.representations.map(\.pixelsWide).max() ?? Int(image.size.width)
    let pixelHeight = image.representations.map(\.pixelsHigh).max() ?? Int(image.size.height)
    guard pixelWidth > 0, pixelHeight > 0 else {
        fail("Screenshot candidate has invalid dimensions: \(url.path)")
    }
    return (url, image, CGFloat(pixelWidth), CGFloat(pixelHeight))
}

let columns = min(6, images.count)
let rows = Int(ceil(Double(images.count) / Double(columns)))
let margin = 28
let gap = 20
let cellWidth = 250
let imageHeight = 500
let labelHeight = 42
let cellHeight = imageHeight + labelHeight
let canvasWidth = margin * 2 + columns * cellWidth + (columns - 1) * gap
let canvasHeight = margin * 2 + rows * cellHeight + (rows - 1) * gap

guard let bitmap = NSBitmapImageRep(
    bitmapDataPlanes: nil,
    pixelsWide: canvasWidth,
    pixelsHigh: canvasHeight,
    bitsPerSample: 8,
    samplesPerPixel: 4,
    hasAlpha: true,
    isPlanar: false,
    colorSpaceName: .deviceRGB,
    bytesPerRow: 0,
    bitsPerPixel: 0
) else {
    fail("Could not allocate the screenshot contact sheet.")
}

guard let context = NSGraphicsContext(bitmapImageRep: bitmap) else {
    fail("Could not create the screenshot contact-sheet graphics context.")
}

NSGraphicsContext.saveGraphicsState()
NSGraphicsContext.current = context
NSColor(calibratedWhite: 0.055, alpha: 1).setFill()
NSRect(x: 0, y: 0, width: canvasWidth, height: canvasHeight).fill()

let labelAttributes: [NSAttributedString.Key: Any] = [
    .font: NSFont.systemFont(ofSize: 17, weight: .semibold),
    .foregroundColor: NSColor(calibratedWhite: 0.92, alpha: 1)
]

for (index, entry) in images.enumerated() {
    let column = index % columns
    let row = index / columns
    let cellX = margin + column * (cellWidth + gap)
    let cellY = canvasHeight - margin - (row + 1) * cellHeight - row * gap
    let scale = min(CGFloat(cellWidth) / entry.width, CGFloat(imageHeight) / entry.height)
    let drawWidth = entry.width * scale
    let drawHeight = entry.height * scale
    let imageRect = NSRect(
        x: CGFloat(cellX) + (CGFloat(cellWidth) - drawWidth) / 2,
        y: CGFloat(cellY + labelHeight) + (CGFloat(imageHeight) - drawHeight) / 2,
        width: drawWidth,
        height: drawHeight
    )
    entry.image.draw(in: imageRect)

    let appearance = entry.url.deletingLastPathComponent().lastPathComponent
    let name = entry.url.deletingPathExtension().lastPathComponent
    let label = "\(appearance) · \(name)" as NSString
    label.draw(
        in: NSRect(x: cellX, y: cellY + 9, width: cellWidth, height: labelHeight - 9),
        withAttributes: labelAttributes
    )
}

context.flushGraphics()
NSGraphicsContext.restoreGraphicsState()

guard let png = bitmap.representation(using: .png, properties: [:]) else {
    fail("Could not encode the screenshot contact sheet as PNG.")
}
try fileManager.createDirectory(at: outputURL.deletingLastPathComponent(), withIntermediateDirectories: true)
do {
    try png.write(to: outputURL, options: .atomic)
} catch {
    fail("Could not write screenshot contact sheet: \(error)")
}

print("Screenshot contact sheet written to \(outputURL.path)")
