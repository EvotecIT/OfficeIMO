import CoreGraphics
import QuartzCore

/// The page images of one page view, as layers in PDF points. Each page has a preview of the whole page under tiles of
/// the visible part at the current zoom, rendered by the engine as exact pixel rectangles that join without seams.
/// Pages and tiles away from the visible area are released, and tiles of an earlier zoom stay until the new ones
/// cover the view, so zooming never shows a blank page.
@MainActor
final class PageTiles {
    /// The side of a tile in pixels.
    static let tileSize = 512
    /// The pixel area of a page preview.
    static let previewPixels = 1_200_000.0

    private struct TileKey: Hashable {
        let page: Int
        let scale: Double
        let column: Int
        let row: Int
    }

    /// The pages, positioned in PDF points from the top left. The host adds it to its zooming view.
    let layer = CALayer()
    private(set) var layout = PageLayout(pages: [])

    private var engine: EngineDocument?
    private var facts: DocumentFacts?
    private var pageLayers: [CALayer] = []
    private var clipLayers: [CALayer] = []
    private var previews: [Int: CALayer] = [:]
    private var previewTasks: [Int: Task<Void, Never>] = [:]
    private var tiles: [TileKey: CALayer] = [:]
    private var tileTasks: [TileKey: Task<Void, Never>] = [:]
    private var needed = Set<TileKey>()
    private var tileScale = 0.0
    /// Images of the earlier revision, shown over the new ones until those cover the view, so an edit never blanks it.
    private var staleLayers: [CALayer] = []
    private var lastVisible = CGRect.null
    private var shownRevision: Int64?
    /// Called once each revision's images cover the visible area.
    var revisionShown: ((Int64) -> Void)?
    private var highlights: [SearchResult.Hit] = []
    private var currentHighlight: Int?
    private var highlightLayers: [CALayer] = []

    init() {
        layer.anchorPoint = .zero
    }

    /// Shows a revision of a document. A new revision lays the pages out again and replaces every image.
    func show(_ engine: EngineDocument, facts: DocumentFacts) {
        guard engine !== self.engine || facts.revision != self.facts?.revision else { return }
        cancelAll()
        let newLayout = PageLayout(pages: facts.pages)
        let keepsLayout = engine === self.engine && newLayout == layout
        self.engine = engine
        self.facts = facts
        shownRevision = nil
        if keepsLayout {
            withoutAnimation {
                for layer in Array(previews.values) + Array(tiles.values) {
                    layer.zPosition = 3
                    staleLayers.append(layer)
                }
            }
            previews.removeAll()
            tiles.removeAll()
            needed.removeAll()
            return
        }
        layout = newLayout
        withoutAnimation {
            staleLayers.forEach { $0.removeFromSuperlayer() }
            staleLayers.removeAll()
            pageLayers.forEach { $0.removeFromSuperlayer() }
            pageLayers.removeAll()
            clipLayers.removeAll()
            previews.removeAll()
            tiles.removeAll()
            layer.bounds = CGRect(origin: .zero, size: layout.size)
            for frame in layout.frames {
                let page = CALayer()
                page.frame = frame
                page.backgroundColor = CGColor(gray: 1, alpha: 1)
                page.shadowOpacity = 0.2
                page.shadowRadius = 3
                page.shadowOffset = CGSize(width: 0, height: 1)
                page.shadowPath = CGPath(rect: page.bounds, transform: nil)
                let clip = CALayer()
                clip.frame = page.bounds
                clip.masksToBounds = true
                page.addSublayer(clip)
                layer.addSublayer(page)
                pageLayers.append(page)
                clipLayers.append(clip)
            }
            highlightLayers.removeAll()
            drawHighlights()
            selectionLayers.removeAll()
            drawSelection()
        }
    }

    private var selection: (page: Int, rects: [CGRect])?
    private var selectionLayers: [CALayer] = []
    /// The selection's colour: the platform's text selection colour.
    var selectionColor = CGColor(red: 0, green: 0.48, blue: 1, alpha: 0.3)

    /// Shows a text selection's lines on a page, in the displayed page's points.
    func select(page: Int, rects: [CGRect]) {
        guard selection?.page != page || selection?.rects != rects else { return }
        selection = rects.isEmpty ? nil : (page, rects)
        withoutAnimation { drawSelection() }
    }

    private func drawSelection() {
        selectionLayers.forEach { $0.removeFromSuperlayer() }
        selectionLayers.removeAll()
        guard let selection, clipLayers.indices.contains(selection.page - 1) else { return }
        for rect in selection.rects {
            let mark = CALayer()
            mark.frame = rect.insetBy(dx: -0.5, dy: -0.5)
            mark.backgroundColor = selectionColor
            mark.zPosition = 11
            clipLayers[selection.page - 1].addSublayer(mark)
            selectionLayers.append(mark)
        }
    }

    /// Marks search hits on the pages, the current one more strongly. Rectangles are in the displayed page's points.
    func highlight(_ hits: [SearchResult.Hit], current: Int?) {
        guard hits.map(\.occurrence) != highlights.map(\.occurrence) || current != currentHighlight else { return }
        highlights = hits
        currentHighlight = current
        withoutAnimation { drawHighlights() }
    }

    private func drawHighlights() {
        highlightLayers.forEach { $0.removeFromSuperlayer() }
        highlightLayers.removeAll()
        for hit in highlights where clipLayers.indices.contains(hit.page - 1) {
            let isCurrent = hit.occurrence == currentHighlight
            for line in hit.lineBounds {
                let mark = CALayer()
                mark.frame = line.cgRect.insetBy(dx: -1.5, dy: -1)
                mark.cornerRadius = 2
                // Multiplying keeps the text legible under the mark, as a highlighter does.
                mark.compositingFilter = "multiplyBlendMode"
                mark.backgroundColor = isCurrent ? CGColor(red: 1, green: 0.62, blue: 0, alpha: 0.75)
                                                 : CGColor(red: 1, green: 0.87, blue: 0.2, alpha: 0.55)
                mark.zPosition = 10
                clipLayers[hit.page - 1].addSublayer(mark)
                highlightLayers.append(mark)
            }
        }
    }

    /// Brings the images in line with the view: visible is the visible rectangle in PDF points and pixelsPerPoint the
    /// device pixels a point covers. While a pinch or magnify is in progress (settled is false), the tiles keep their
    /// zoom and the system scales them; the gesture's end renders the new zoom.
    func update(visible: CGRect, pixelsPerPoint: Double, settled: Bool) {
        guard let facts, !layout.frames.isEmpty, !visible.isEmpty else { return }
        lastVisible = visible
        if settled || tileScale == 0 { setTileScale(pixelsPerPoint) }

        // Previews for the visible pages and a screen above and below.
        let previewRange = layout.pages(in: visible.insetBy(dx: 0, dy: -visible.height)) ?? 1...1
        for (page, task) in previewTasks where !previewRange.contains(page) {
            task.cancel()
            previewTasks[page] = nil
        }
        let keep = max(1, previewRange.lowerBound - 2)...min(layout.frames.count, previewRange.upperBound + 2)
        withoutAnimation {
            for (page, preview) in previews where !keep.contains(page) {
                preview.removeFromSuperlayer()
                previews[page] = nil
            }
        }
        for page in previewRange where previews[page] == nil && previewTasks[page] == nil {
            renderPreview(page, facts: facts)
        }

        // Tiles for the visible area and a quarter of a screen around it, nearest the middle first.
        let near = visible.insetBy(dx: -visible.width / 4, dy: -visible.height / 4)
        needed = tileKeys(in: near, facts: facts)
        for (key, task) in tileTasks where !needed.contains(key) {
            task.cancel()
            tileTasks[key] = nil
        }
        removeTiles(outside: needed.allSatisfy { tiles[$0] != nil } ? nil : near)
        let middle = CGPoint(x: visible.midX, y: visible.midY)
        let missing = needed.filter { tiles[$0] == nil && tileTasks[$0] == nil }
            .sorted { distance(of: $0, to: middle) < distance(of: $1, to: middle) }
        for key in missing { renderTile(key) }
        checkCoverage()
    }

    /// Once the current revision's previews and tiles cover the visible area, the earlier revision's images go.
    private func checkCoverage() {
        guard let facts, shownRevision != facts.revision, !lastVisible.isNull,
              let visiblePages = layout.pages(in: lastVisible),
              visiblePages.allSatisfy({ previews[$0] != nil }),
              needed.allSatisfy({ tiles[$0] != nil }) else { return }
        shownRevision = facts.revision
        withoutAnimation {
            staleLayers.forEach { $0.removeFromSuperlayer() }
            staleLayers.removeAll()
        }
        revisionShown?(facts.revision)
    }

    func cancelAll() {
        previewTasks.values.forEach { $0.cancel() }
        tileTasks.values.forEach { $0.cancel() }
        previewTasks.removeAll()
        tileTasks.removeAll()
    }

    // MARK: Previews

    private func previewScale(_ page: Int) -> Double {
        let size = layout.frames[page - 1].size
        return min(2, (Self.previewPixels / (size.width * size.height)).squareRoot())
    }

    private func renderPreview(_ page: Int, facts: DocumentFacts) {
        guard let engine else { return }
        let scale = previewScale(page)
        let size = facts.pageImageSize(page: page, scale: scale)
        let revision = facts.revision
        previewTasks[page] = Task {
            let image = try? await engine.renderRegion(page, scale: scale, x: 0, y: 0, width: size.width, height: size.height)
            guard !Task.isCancelled, revision == self.facts?.revision else { return }
            previewTasks[page] = nil
            guard let image else { return }
            let preview = imageLayer(image, frame: CGRect(x: 0, y: 0, width: Double(size.width) / scale, height: Double(size.height) / scale))
            withoutAnimation { clipLayers[page - 1].insertSublayer(preview, at: 0) }
            previews[page] = preview
            checkCoverage()
        }
    }

    // MARK: Tiles

    private func setTileScale(_ scale: Double) {
        guard scale != tileScale else { return }
        tileScale = scale
        withoutAnimation {
            for (key, tile) in tiles { tile.zPosition = key.scale == scale ? 2 : 1 }
        }
    }

    private func tileKeys(in rect: CGRect, facts: DocumentFacts) -> Set<TileKey> {
        guard tileScale > 0, let range = layout.pages(in: rect) else { return [] }
        var keys = Set<TileKey>()
        let side = Double(Self.tileSize)
        for page in range where tileScale > previewScale(page) * 1.15 {
            let frame = layout.frames[page - 1]
            let local = rect.intersection(frame).offsetBy(dx: -frame.minX, dy: -frame.minY)
            guard !local.isNull, !local.isEmpty else { continue }
            let size = facts.pageImageSize(page: page, scale: tileScale)
            let columns = (size.width + Self.tileSize - 1) / Self.tileSize
            let rows = (size.height + Self.tileSize - 1) / Self.tileSize
            let firstColumn = max(0, Int(local.minX * tileScale / side))
            let lastColumn = min(columns - 1, Int(local.maxX * tileScale / side))
            let firstRow = max(0, Int(local.minY * tileScale / side))
            let lastRow = min(rows - 1, Int(local.maxY * tileScale / side))
            guard firstColumn <= lastColumn, firstRow <= lastRow else { continue }
            for row in firstRow...lastRow {
                for column in firstColumn...lastColumn {
                    keys.insert(TileKey(page: page, scale: tileScale, column: column, row: row))
                }
            }
        }
        return keys
    }

    /// A tile's pixel rectangle in the page image.
    private func pixelRect(_ key: TileKey) -> (x: Int, y: Int, width: Int, height: Int)? {
        guard let facts else { return nil }
        let size = facts.pageImageSize(page: key.page, scale: key.scale)
        let x = key.column * Self.tileSize
        let y = key.row * Self.tileSize
        return (x, y, min(Self.tileSize, size.width - x), min(Self.tileSize, size.height - y))
    }

    /// A tile's rectangle in PDF points, from the page's top left.
    private func pageRect(_ key: TileKey) -> CGRect {
        guard let pixels = pixelRect(key) else { return .null }
        return CGRect(x: Double(pixels.x) / key.scale, y: Double(pixels.y) / key.scale,
                      width: Double(pixels.width) / key.scale, height: Double(pixels.height) / key.scale)
    }

    private func distance(of key: TileKey, to point: CGPoint) -> Double {
        let rect = pageRect(key).offsetBy(dx: layout.frames[key.page - 1].minX, dy: layout.frames[key.page - 1].minY)
        return hypot(rect.midX - point.x, rect.midY - point.y)
    }

    private func renderTile(_ key: TileKey) {
        guard let engine, let pixels = pixelRect(key), pixels.width > 0, pixels.height > 0 else { return }
        let revision = facts?.revision
        tileTasks[key] = Task {
            let image = try? await engine.renderRegion(key.page, scale: key.scale, x: pixels.x, y: pixels.y,
                                                      width: pixels.width, height: pixels.height)
            guard !Task.isCancelled, revision == self.facts?.revision else { return }
            tileTasks[key] = nil
            guard let image else { return }
            let tile = imageLayer(image, frame: pageRect(key))
            tile.zPosition = key.scale == tileScale ? 2 : 1
            withoutAnimation { clipLayers[key.page - 1].addSublayer(tile) }
            tiles[key] = tile
            if needed.allSatisfy({ tiles[$0] != nil }) { removeTiles(outside: nil) }
            checkCoverage()
        }
    }

    /// Removes tiles no longer needed. Tiles of another zoom stay inside rect until the current zoom covers it; with no
    /// rect they all go.
    private func removeTiles(outside rect: CGRect?) {
        withoutAnimation {
            for (key, tile) in tiles where !needed.contains(key) {
                if key.scale != tileScale, let rect, tiles.count < 256 {
                    let frame = layout.frames[key.page - 1]
                    if pageRect(key).offsetBy(dx: frame.minX, dy: frame.minY).intersects(rect) { continue }
                }
                tile.removeFromSuperlayer()
                tiles[key] = nil
            }
        }
    }

    // MARK: Layers

    private func imageLayer(_ image: CGImage, frame: CGRect) -> CALayer {
        let layer = CALayer()
        layer.contents = image
        layer.contentsGravity = .resize
        layer.isOpaque = true
        layer.frame = frame
        return layer
    }

    private func withoutAnimation(_ work: () -> Void) {
        CATransaction.begin()
        CATransaction.setDisableActions(true)
        work()
        CATransaction.commit()
    }
}
