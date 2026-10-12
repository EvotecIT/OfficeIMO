import CoreGraphics
import Foundation

/// Feasibility measurements for one PDF: opening, the first page, a repeated view of it, sampled pages at the display's
/// width, the PNG path for comparison, zoomed viewports rendered as tiles, and the process footprint. Run one document
/// per launch so the peak footprint belongs to it.
struct MeasurementReport: Codable, Sendable {
    struct Render: Codable, Sendable {
        let page: Int
        let scale: Double
        let pixelWidth: Int
        let pixelHeight: Int
        let milliseconds: Double
    }

    struct Png: Codable, Sendable {
        let renderMilliseconds: Double
        let decodeMilliseconds: Double
        let pngBytes: Int
    }

    /// A screen-sized view of page 1 at a zoom of the fit-to-width scale, from its top-left corner, as tiles.
    struct Viewport: Codable, Sendable {
        let zoom: Double
        let scale: Double
        let tiles: Int
        let pixelWidth: Int
        let pixelHeight: Int
        let milliseconds: Double
        let error: String?
    }

    struct Summary: Codable, Sendable {
        let count: Int
        let median: Double
        let p90: Double
        let max: Double
    }

    let file: String
    let fileSize: Int64
    let pages: Int
    let device: String
    let system: String
    let display: DisplayTarget
    let engineVersion: String
    let launch: JourneyModel.LaunchTiming
    let openMilliseconds: Double
    /// From the start of opening until page 1 is an image ready to show.
    let firstPageMilliseconds: Double
    let firstPage: Render
    /// Page 1 again, with its drawing and images kept by the engine.
    let repeatFirstPage: Render
    /// Page 1 through the PNG path at the same scale.
    let firstPagePng: Png
    let sampled: [Render]
    let sampledSummary: Summary
    let zoom: [Viewport]
    let footprintBeforeOpen: ProcessMetrics.Footprint
    let footprintAfter: ProcessMetrics.Footprint
}

enum MeasurementRunner {
    /// Pages rendered beyond page 1, spread evenly through the document.
    static let sampleCount = 24
    /// Zoom factors of the fit-to-width scale for the tiled viewports.
    static let zoomFactors: [Double] = [1.5, 2, 3, 4]
    static let tileSize = 1024

    @MainActor
    static func run(_ url: URL, display: DisplayTarget, launch: JourneyModel.LaunchTiming, engineVersion: String) async throws -> MeasurementReport {
        let before = ProcessMetrics.footprint()
        let start = ContinuousClock.now
        let (document, openTime) = try await measure { try await EngineDocument.open(url) }
        let facts = try await document.facts()
        let first = try await render(document, facts, page: 1, display: display)
        let firstPageTime = milliseconds(since: start)
        let repeated = try await render(document, facts, page: 1, display: display)

        let (png, pngRender) = try await measure { try await document.renderPage(1, scale: first.scale) }
        let (_, pngDecode) = try await measure { try await Task.detached { try PageImage.decode(png.png) }.value }

        var sampled: [MeasurementReport.Render] = []
        for page in samplePages(count: facts.pages.count) {
            sampled.append(try await render(document, facts, page: page, display: display))
        }

        var zoom: [MeasurementReport.Viewport] = []
        let fitScale = display.pixelWidth / facts.pages[0].displaySize.width
        for factor in zoomFactors {
            zoom.append(await viewport(document, facts, scale: fitScale * factor, zoom: factor, display: display))
        }

        let attributes = try? FileManager.default.attributesOfItem(atPath: url.path)
        let times = sampled.map(\.milliseconds).sorted()
        func percentile(_ fraction: Double) -> Double {
            times.isEmpty ? 0 : times[min(times.count - 1, Int((Double(times.count - 1) * fraction).rounded()))]
        }
        return MeasurementReport(
            file: url.lastPathComponent,
            fileSize: (attributes?[.size] as? NSNumber)?.int64Value ?? 0,
            pages: facts.pages.count,
            device: ProcessMetrics.deviceModel,
            system: ProcessInfo.processInfo.operatingSystemVersionString,
            display: display,
            engineVersion: engineVersion,
            launch: launch,
            openMilliseconds: openTime,
            firstPageMilliseconds: firstPageTime,
            firstPage: first,
            repeatFirstPage: repeated,
            firstPagePng: .init(renderMilliseconds: pngRender, decodeMilliseconds: pngDecode, pngBytes: png.png.count),
            sampled: sampled,
            sampledSummary: .init(count: times.count, median: percentile(0.5), p90: percentile(0.9), max: times.last ?? 0),
            zoom: zoom,
            footprintBeforeOpen: before,
            footprintAfter: ProcessMetrics.footprint())
    }

    /// A page filling the display's width, within the engine's budget, as one region.
    private static func render(_ document: EngineDocument, _ facts: DocumentFacts, page: Int, display: DisplayTarget) async throws -> MeasurementReport.Render {
        let scale = facts.renderScale(page: page, fillingPixelWidth: display.pixelWidth)
        let size = facts.pageImageSize(page: page, scale: scale)
        let (_, elapsed) = try await measure {
            try await document.renderRegion(page, scale: scale, x: 0, y: 0, width: size.width, height: size.height)
        }
        return .init(page: page, scale: scale, pixelWidth: size.width, pixelHeight: size.height, milliseconds: elapsed)
    }

    /// The display-sized area at the page's top-left corner, rendered as tiles of at most tileSize pixels a side.
    private static func viewport(_ document: EngineDocument, _ facts: DocumentFacts, scale: Double, zoom: Double,
                                 display: DisplayTarget) async -> MeasurementReport.Viewport {
        let page = facts.pageImageSize(page: 1, scale: scale)
        let width = min(page.width, Int(display.pixelWidth)), height = min(page.height, Int(display.pixelHeight))
        var tiles = 0
        let start = ContinuousClock.now
        do {
            for y in stride(from: 0, to: height, by: tileSize) {
                for x in stride(from: 0, to: width, by: tileSize) {
                    _ = try await document.renderRegion(1, scale: scale, x: x, y: y, width: min(tileSize, width - x),
                                                        height: min(tileSize, height - y))
                    tiles += 1
                }
            }
            return .init(zoom: zoom, scale: scale, tiles: tiles, pixelWidth: width, pixelHeight: height,
                         milliseconds: milliseconds(since: start), error: nil)
        } catch {
            return .init(zoom: zoom, scale: scale, tiles: tiles, pixelWidth: width, pixelHeight: height,
                         milliseconds: milliseconds(since: start), error: error.localizedDescription)
        }
    }

    /// Up to sampleCount pages after page 1, evenly spaced and including the last page.
    static func samplePages(count: Int) -> [Int] {
        guard count > 1 else { return [] }
        let wanted = min(sampleCount, count - 1)
        return Array(Set((1...wanted).map { 1 + Int((Double($0) * Double(count - 1) / Double(wanted)).rounded()) })).sorted()
    }
}
