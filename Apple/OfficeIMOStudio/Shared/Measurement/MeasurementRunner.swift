import CoreGraphics
import Foundation

/// Feasibility measurements for one PDF: opening, the first page, rendering sampled pages at the display's scale,
/// zooming page 1, and the process footprint. Run one document per launch so the peak footprint belongs to it.
struct MeasurementReport: Codable, Sendable {
    struct Render: Codable, Sendable {
        let page: Int
        let scale: Double
        let pixelWidth: Int
        let pixelHeight: Int
        let pngBytes: Int
        let renderMilliseconds: Double
        let decodeMilliseconds: Double
    }

    struct Zoom: Codable, Sendable {
        let zoom: Double
        let render: Render?
        let error: String?
    }

    struct Summary: Codable, Sendable {
        let count: Int
        let medianRender: Double
        let p90Render: Double
        let maxRender: Double
        let medianDecode: Double
        let medianPngBytes: Int
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
    /// From the start of opening until page 1 is decoded and ready to show.
    let firstPageMilliseconds: Double
    let firstPage: Render
    let sampled: [Render]
    let sampledSummary: Summary
    let zoom: [Zoom]
    let footprintBeforeOpen: ProcessMetrics.Footprint
    let footprintAfter: ProcessMetrics.Footprint
}

enum MeasurementRunner {
    /// Pages rendered beyond page 1, spread evenly through the document.
    static let sampleCount = 24
    /// Zoom factors applied to page 1 on top of the fit-to-width scale.
    static let zoomFactors: [Double] = [1, 1.5, 2, 3, 4]

    @MainActor
    static func run(_ url: URL, display: DisplayTarget, launch: JourneyModel.LaunchTiming, engineVersion: String) async throws -> MeasurementReport {
        let before = ProcessMetrics.footprint()
        let start = ContinuousClock.now
        let (document, openTime) = try await measure { try await EngineDocument.open(url) }
        let facts = try await document.facts()
        let first = try await render(document, facts, page: 1, display: display, zoom: 1)
        let firstPageTime = milliseconds(since: start)

        var sampled: [MeasurementReport.Render] = []
        for page in samplePages(count: facts.pages.count) {
            sampled.append(try await render(document, facts, page: page, display: display, zoom: 1))
        }

        var zoom: [MeasurementReport.Zoom] = []
        for factor in zoomFactors {
            do {
                zoom.append(.init(zoom: factor, render: try await render(document, facts, page: 1, display: display, zoom: factor), error: nil))
            } catch {
                zoom.append(.init(zoom: factor, render: nil, error: error.localizedDescription))
            }
        }

        let attributes = try? FileManager.default.attributesOfItem(atPath: url.path)
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
            sampled: sampled,
            sampledSummary: summarize(sampled),
            zoom: zoom,
            footprintBeforeOpen: before,
            footprintAfter: ProcessMetrics.footprint())
    }

    private static func render(_ document: EngineDocument, _ facts: DocumentFacts, page: Int, display: DisplayTarget, zoom: Double) async throws -> MeasurementReport.Render {
        // Pages fill the display width within the engine's budget; zoom goes past it, to show where tiles are needed.
        let scale = zoom == 1
            ? facts.renderScale(page: page, fillingPixelWidth: display.pixelWidth)
            : display.pixelWidth / facts.pages[page - 1].displaySize.width * zoom
        let (rendered, renderTime) = try await measure { try await document.renderPage(page, scale: scale) }
        let (_, decodeTime) = try await measure { try await Task.detached { try PageImage.decode(rendered.png) }.value }
        return .init(page: page, scale: scale, pixelWidth: rendered.pixelWidth, pixelHeight: rendered.pixelHeight,
                     pngBytes: rendered.png.count, renderMilliseconds: renderTime, decodeMilliseconds: decodeTime)
    }

    /// Up to sampleCount pages after page 1, evenly spaced and including the last page.
    static func samplePages(count: Int) -> [Int] {
        guard count > 1 else { return [] }
        let wanted = min(sampleCount, count - 1)
        return Array(Set((1...wanted).map { 1 + Int((Double($0) * Double(count - 1) / Double(wanted)).rounded()) })).sorted()
    }

    private static func summarize(_ renders: [MeasurementReport.Render]) -> MeasurementReport.Summary {
        func percentile(_ values: [Double], _ fraction: Double) -> Double {
            guard !values.isEmpty else { return 0 }
            let sorted = values.sorted()
            return sorted[min(sorted.count - 1, Int((Double(sorted.count - 1) * fraction).rounded()))]
        }
        let render = renders.map(\.renderMilliseconds)
        let bytes = renders.map { Double($0.pngBytes) }
        return .init(count: renders.count, medianRender: percentile(render, 0.5), p90Render: percentile(render, 0.9),
                     maxRender: render.max() ?? 0, medianDecode: percentile(renders.map(\.decodeMilliseconds), 0.5),
                     medianPngBytes: Int(percentile(bytes, 0.5)))
    }
}
