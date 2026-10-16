import CoreGraphics
import SwiftUI

/// How the Read and Markup tasks lay the pages out, as Preview and OfficeIMO Studio on Windows and Linux offer them.
enum ReadingLayout: String, CaseIterable, Identifiable {
    case singlePage, continuous, twoPages, grid

    var id: Self { self }

    var title: LocalizedStringKey {
        switch self {
        case .singlePage: "Single Page"
        case .continuous: "Continuous"
        case .twoPages: "Two Pages"
        case .grid: "Grid"
        }
    }

    var systemImage: String {
        switch self {
        case .singlePage: "doc"
        case .continuous: "rectangle.split.1x2"
        case .twoPages: "book"
        case .grid: "square.grid.2x2"
        }
    }

    /// The pages a grid shows side by side.
    static let gridColumns = 4
}

/// Pages in PDF points, in one of the reading layouts: one column, a single page, a cover followed by two-page
/// spreads, or a grid. The scroll view zooms the whole layout, so a page's frame never changes with the zoom and the
/// tiles only need to know how many pixels a point covers. In Single Page every other page has an empty frame.
///
/// The view can turn every page a quarter, a half or three quarters clockwise without changing the document. Frames are
/// then the turned pages; the views over them convert between the layout and each upright page with pagePoint,
/// layoutPoint, pageRect and layoutRect, which add only the frame's origin when the view isn't turned.
struct PageLayout: Equatable {
    static let spacing: CGFloat = 12
    static let margin: CGFloat = 12

    let frames: [CGRect]
    let size: CGSize
    /// Clockwise degrees: 0, 90, 180 or 270.
    let rotation: Int
    /// Each page's upright size in displayed points.
    private let pageSizes: [CGSize]

    init(pages: [DocumentFacts.Page], mode: ReadingLayout = .continuous, shownPage: Int = 1, rotation: Int = 0) {
        let rotation = ((rotation % 360) + 360) % 360 / 90 * 90
        self.rotation = rotation
        pageSizes = pages.map(\.displaySize)
        let sizes = pageSizes.map { rotation % 180 == 0 ? $0 : CGSize(width: $0.height, height: $0.width) }
        let widest = sizes.map(\.width).max() ?? 0
        let tallest = sizes.map(\.height).max() ?? 0
        var frames = Array(repeating: CGRect.zero, count: sizes.count)
        var size = CGSize.zero
        switch mode {
        case _ where sizes.isEmpty:
            break
        case .continuous:
            var y = Self.margin
            for (index, page) in sizes.enumerated() {
                frames[index] = CGRect(x: Self.margin + (widest - page.width) / 2, y: y, width: page.width, height: page.height)
                y += page.height + Self.spacing
            }
            size = CGSize(width: widest + 2 * Self.margin, height: y - Self.spacing + Self.margin)
        case .singlePage:
            let index = min(max(shownPage, 1), sizes.count) - 1
            let page = sizes[index]
            frames[index] = CGRect(x: Self.margin, y: Self.margin, width: page.width, height: page.height)
            size = CGSize(width: page.width + 2 * Self.margin, height: page.height + 2 * Self.margin)
        case .twoPages:
            // The cover stands alone on the right, then left and right pages meet in the middle, as in a book.
            var y = Self.margin
            var index = 0
            var row = 0
            while index < sizes.count {
                let members = row == 0 ? [index] : Array(index..<min(index + 2, sizes.count))
                let height = members.map { sizes[$0].height }.max() ?? 0
                for (slot, member) in members.enumerated() {
                    let page = sizes[member]
                    let isRight = row == 0 || slot == 1
                    let x = isRight ? Self.margin + widest + Self.spacing : Self.margin + widest - page.width
                    frames[member] = CGRect(x: x, y: y, width: page.width, height: page.height)
                }
                y += height + Self.spacing
                index += members.count
                row += 1
            }
            size = CGSize(width: 2 * widest + Self.spacing + 2 * Self.margin, height: y - Self.spacing + Self.margin)
        case .grid:
            let columns = min(ReadingLayout.gridColumns, sizes.count)
            for (index, page) in sizes.enumerated() {
                let column = CGFloat(index % columns), row = CGFloat(index / columns)
                let cell = CGPoint(x: Self.margin + column * (widest + Self.spacing), y: Self.margin + row * (tallest + Self.spacing))
                frames[index] = CGRect(x: cell.x + (widest - page.width) / 2, y: cell.y + (tallest - page.height) / 2,
                                       width: page.width, height: page.height)
            }
            let rows = CGFloat((sizes.count + columns - 1) / columns)
            size = CGSize(width: CGFloat(columns) * (widest + Self.spacing) - Self.spacing + 2 * Self.margin,
                          height: rows * (tallest + Self.spacing) - Self.spacing + 2 * Self.margin)
        }
        self.frames = frames
        self.size = size
    }

    /// The page, numbered from 1, under a point; elsewhere the nearest shown page.
    func page(at point: CGPoint) -> Int {
        guard !frames.isEmpty else { return 1 }
        if let index = frames.firstIndex(where: { $0.contains(point) }) { return index + 1 }
        var best = 1
        var distance = CGFloat.greatestFiniteMagnitude
        for (index, frame) in frames.enumerated() where !frame.isEmpty {
            let dx = max(frame.minX - point.x, 0, point.x - frame.maxX), dy = max(frame.minY - point.y, 0, point.y - frame.maxY)
            if dx * dx + dy * dy < distance {
                distance = dx * dx + dy * dy
                best = index + 1
            }
        }
        return best
    }

    /// The shown page, numbered from 1, that shows the most of itself in a rectangle.
    func mostVisiblePage(in rect: CGRect) -> Int {
        guard let range = pages(in: rect) else { return page(at: CGPoint(x: rect.midX, y: rect.midY)) }
        return range.filter { !frames[$0 - 1].isEmpty }.max {
            area(frames[$0 - 1].intersection(rect)) < area(frames[$1 - 1].intersection(rect))
        } ?? range.lowerBound
    }

    /// Pages, numbered from 1, from the first to the last shown page that a rectangle touches.
    func pages(in rect: CGRect) -> ClosedRange<Int>? {
        var first: Int?
        var last = 0
        for (index, frame) in frames.enumerated() where !frame.isEmpty && frame.intersects(rect) {
            if first == nil { first = index + 1 }
            last = index + 1
        }
        guard let first else { return nil }
        return first...last
    }

    private func area(_ rect: CGRect) -> CGFloat { rect.isNull ? 0 : rect.width * rect.height }

    // MARK: The layout and the upright page

    /// A page's upright size in displayed points.
    func pageSize(_ page: Int) -> CGSize {
        pageSizes.indices.contains(page - 1) ? pageSizes[page - 1] : .zero
    }

    /// A point of the layout as a point of the upright page, in its displayed points from its top left.
    func pagePoint(_ point: CGPoint, page: Int) -> CGPoint {
        guard frames.indices.contains(page - 1) else { return point }
        let frame = frames[page - 1], size = pageSizes[page - 1]
        let v = CGPoint(x: point.x - frame.minX, y: point.y - frame.minY)
        switch rotation {
        case 90: return CGPoint(x: v.y, y: size.height - v.x)
        case 180: return CGPoint(x: size.width - v.x, y: size.height - v.y)
        case 270: return CGPoint(x: size.width - v.y, y: v.x)
        default: return v
        }
    }

    /// A point of the upright page as a point of the layout.
    func layoutPoint(_ point: CGPoint, page: Int) -> CGPoint {
        guard frames.indices.contains(page - 1) else { return point }
        let frame = frames[page - 1], size = pageSizes[page - 1]
        let v: CGPoint
        switch rotation {
        case 90: v = CGPoint(x: size.height - point.y, y: point.x)
        case 180: v = CGPoint(x: size.width - point.x, y: size.height - point.y)
        case 270: v = CGPoint(x: point.y, y: size.width - point.x)
        default: v = point
        }
        return CGPoint(x: v.x + frame.minX, y: v.y + frame.minY)
    }

    /// A rectangle of the upright page as a rectangle of the layout.
    func layoutRect(_ rect: CGRect, page: Int) -> CGRect {
        Self.bounds(layoutPoint(CGPoint(x: rect.minX, y: rect.minY), page: page), layoutPoint(CGPoint(x: rect.maxX, y: rect.maxY), page: page))
    }

    /// A rectangle of the layout as a rectangle of the upright page.
    func pageRect(_ rect: CGRect, page: Int) -> CGRect {
        guard !rect.isNull else { return rect }
        return Self.bounds(pagePoint(CGPoint(x: rect.minX, y: rect.minY), page: page), pagePoint(CGPoint(x: rect.maxX, y: rect.maxY), page: page))
    }

    /// A vector of the layout, such as a drag, as a vector of the upright page.
    func pageVector(_ vector: CGVector) -> CGVector {
        switch rotation {
        case 90: return CGVector(dx: vector.dy, dy: -vector.dx)
        case 180: return CGVector(dx: -vector.dx, dy: -vector.dy)
        case 270: return CGVector(dx: -vector.dy, dy: vector.dx)
        default: return vector
        }
    }

    private static func bounds(_ a: CGPoint, _ b: CGPoint) -> CGRect {
        CGRect(x: min(a.x, b.x), y: min(a.y, b.y), width: abs(a.x - b.x), height: abs(a.y - b.y))
    }
}
