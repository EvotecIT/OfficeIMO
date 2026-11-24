#!/usr/bin/env swift

import CoreGraphics
import Darwin
import Foundation
import ImageIO

guard CommandLine.arguments.count == 2 else {
    fputs("Usage: verify-screenshot-content.swift <screenshot.png>\n", stderr)
    exit(64)
}

let path = CommandLine.arguments[1]
let url = URL(fileURLWithPath: path) as CFURL
guard let source = CGImageSourceCreateWithURL(url, nil),
      let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
    fputs("Could not decode screenshot: \(path)\n", stderr)
    exit(65)
}

let sampleWidth = 64
let sampleHeight = 64
var pixels = [UInt8](repeating: 0, count: sampleWidth * sampleHeight * 4)
guard let context = CGContext(
    data: &pixels,
    width: sampleWidth,
    height: sampleHeight,
    bitsPerComponent: 8,
    bytesPerRow: sampleWidth * 4,
    space: CGColorSpaceCreateDeviceRGB(),
    bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
) else {
    fputs("Could not create screenshot verification context.\n", stderr)
    exit(70)
}

context.interpolationQuality = .low
context.draw(image, in: CGRect(x: 0, y: 0, width: sampleWidth, height: sampleHeight))

var visiblyRenderedPixels = 0
var minimumLuminance = Double.greatestFiniteMagnitude
var maximumLuminance = -Double.greatestFiniteMagnitude
var luminanceSum = 0.0
var luminanceSquaredSum = 0.0
for offset in stride(from: 0, to: pixels.count, by: 4) {
    let red = Double(pixels[offset])
    let green = Double(pixels[offset + 1])
    let blue = Double(pixels[offset + 2])
    let luminance = (0.2126 * red) + (0.7152 * green) + (0.0722 * blue)

    minimumLuminance = min(minimumLuminance, luminance)
    maximumLuminance = max(maximumLuminance, luminance)
    luminanceSum += luminance
    luminanceSquaredSum += luminance * luminance

    if max(red, green, blue) > 24 {
        visiblyRenderedPixels += 1
    }
}

let sampleCount = Double(sampleWidth * sampleHeight)
let renderedRatio = Double(visiblyRenderedPixels) / sampleCount
guard renderedRatio >= 0.02 else {
    fputs("Screenshot is effectively blank (rendered ratio \(renderedRatio)): \(path)\n", stderr)
    exit(1)
}

let meanLuminance = luminanceSum / sampleCount
let luminanceVariance = max(0, (luminanceSquaredSum / sampleCount) - (meanLuminance * meanLuminance))
let luminanceDeviation = sqrt(luminanceVariance)
let luminanceRange = maximumLuminance - minimumLuminance
guard luminanceRange >= 16, luminanceDeviation >= 4 else {
    fputs("Screenshot is effectively blank (luminance range \(luminanceRange), deviation \(luminanceDeviation)): \(path)\n", stderr)
    exit(1)
}
