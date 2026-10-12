import CoreGraphics
import Foundation
import Synchronization

/// Printing draws the engine's own page images, as OfficeIMO Studio prints on every platform, so paper matches the
/// screen. The system's print panel or sheet chooses the printer, paper, pages and copies; each page renders when the
/// system draws it, at the resolution of what it draws into, a preview thumbnail or the printer.
final class PagePrinting: Sendable {
    let document: EngineDocument
    let facts: DocumentFacts
    let jobName: String
    private let failed = Mutex<Set<Int>>([])

    init(document: EngineDocument, facts: DocumentFacts, jobName: String) {
        self.document = document
        self.facts = facts
        self.jobName = jobName
    }

    var pageCount: Int { facts.pages.count }

    /// Whether the first page is wider than tall, so the paper starts in landscape as the document does.
    var startsLandscape: Bool {
        guard let size = facts.pages.first?.displaySize else { return false }
        return size.width > size.height
    }

    /// Pages that could not be rendered; they printed blank.
    var failedPages: [Int] { failed.withLock { $0.sorted() } }

    /// Where a page goes in the printable area: as large as fits, turned a quarter when that fits it larger, as
    /// Preview's Scale to Fit with automatic rotation does.
    static func placement(of page: CGSize, in area: CGRect) -> (rect: CGRect, rotated: Bool) {
        func fit(_ size: CGSize) -> CGFloat { min(area.width / size.width, area.height / size.height) }
        let turned = CGSize(width: page.height, height: page.width)
        let rotated = fit(turned) > fit(page) * 1.001
        let size = rotated ? turned : page
        let scale = fit(size)
        let fitted = CGSize(width: size.width * scale, height: size.height * scale)
        return (CGRect(x: area.midX - fitted.width / 2, y: area.midY - fitted.height / 2, width: fitted.width, height: fitted.height), rotated)
    }

    /// Draws a page (numbered from 1) into a context whose y axis runs down the paper.
    func draw(page number: Int, in area: CGRect, context: CGContext) {
        guard facts.pages.indices.contains(number - 1), area.width > 0, area.height > 0 else { return }
        let (rect, rotated) = Self.placement(of: facts.pages[number - 1].displaySize, in: area)
        let device = context.userSpaceToDeviceSpaceTransform
        let pixelsPerPoint = max(hypot(device.a, device.b), hypot(device.c, device.d), 1)
        let width = (rotated ? rect.height : rect.width)
        let height = (rotated ? rect.width : rect.height)
        let scale = facts.renderScale(page: number, fillingPixelWidth: width * pixelsPerPoint)
        let pixels = facts.pageImageSize(page: number, scale: scale)
        let image: CGImage
        do {
            image = try document.renderPageAndWait(number, scale: scale, width: pixels.width, height: pixels.height)
        } catch {
            _ = failed.withLock { $0.insert(number) }
            return
        }
        context.saveGState()
        context.translateBy(x: rect.midX, y: rect.midY)
        if rotated { context.rotate(by: .pi / 2) }
        // Images draw with their origin at the bottom, so the page is flipped back upright.
        context.scaleBy(x: 1, y: -1)
        context.interpolationQuality = .high
        context.draw(image, in: CGRect(x: -width / 2, y: -height / 2, width: width, height: height))
        context.restoreGState()
    }
}
