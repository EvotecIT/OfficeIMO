import CoreGraphics

/// Pages in one column, in PDF points. The scroll view zooms the whole column, so a page's frame never changes with
/// the zoom and the tiles only need to know how many pixels a point covers.
struct PageLayout: Equatable {
    static let spacing: CGFloat = 12
    static let margin: CGFloat = 12

    let frames: [CGRect]
    let size: CGSize

    init(pages: [DocumentFacts.Page]) {
        let width = pages.map(\.displaySize.width).max() ?? 0
        var y = Self.margin
        var frames: [CGRect] = []
        frames.reserveCapacity(pages.count)
        for page in pages {
            let size = page.displaySize
            frames.append(CGRect(x: Self.margin + (width - size.width) / 2, y: y, width: size.width, height: size.height))
            y += size.height + Self.spacing
        }
        self.frames = frames
        size = CGSize(width: width + 2 * Self.margin, height: frames.isEmpty ? 0 : y - Self.spacing + Self.margin)
    }

    /// The page, numbered from 1, under a point; between pages, the one below.
    func page(atY y: CGFloat) -> Int {
        guard !frames.isEmpty else { return 1 }
        var low = 0
        var high = frames.count - 1
        while low < high {
            let middle = (low + high) / 2
            if frames[middle].maxY + Self.spacing / 2 < y { low = middle + 1 } else { high = middle }
        }
        return low + 1
    }

    /// The page, numbered from 1, that shows the most of itself in a rectangle.
    func mostVisiblePage(in rect: CGRect) -> Int {
        guard let range = pages(in: rect) else { return page(atY: rect.midY) }
        return range.max { frames[$0 - 1].intersection(rect).height < frames[$1 - 1].intersection(rect).height } ?? range.lowerBound
    }

    /// Pages, numbered from 1, that a rectangle touches.
    func pages(in rect: CGRect) -> ClosedRange<Int>? {
        guard !frames.isEmpty, rect.maxY >= frames[0].minY, rect.minY <= frames[frames.count - 1].maxY else { return nil }
        let first = page(atY: rect.minY)
        let last = page(atY: rect.maxY)
        return first...max(first, last)
    }
}
