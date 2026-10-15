#if os(macOS)
import AppKit

/// The pages as one tall view, a sheet of paper per page, for the system print panel, its preview and the printer.
final class PagePrintView: NSView {
    private let printing: PagePrinting
    private var sheetSize = NSSize(width: 612, height: 792)

    init(_ printing: PagePrinting, printInfo: NSPrintInfo) {
        self.printing = printing
        super.init(frame: .zero)
        fitSheets(to: printInfo)
    }

    @available(*, unavailable)
    required init?(coder: NSCoder) { fatalError("init(coder:) is not supported") }

    override var isFlipped: Bool { true }

    override func knowsPageRange(_ range: NSRangePointer) -> Bool {
        // The panel's paper and orientation may have changed since the last pagination.
        if let info = NSPrintOperation.current?.printInfo { fitSheets(to: info) }
        range.pointee = NSRange(location: 1, length: printing.pageCount)
        return true
    }

    override func rectForPage(_ page: Int) -> NSRect {
        NSRect(origin: NSPoint(x: 0, y: CGFloat(page - 1) * sheetSize.height), size: sheetSize)
    }

    override func draw(_ dirtyRect: NSRect) {
        guard let context = NSGraphicsContext.current?.cgContext, sheetSize.height > 0 else { return }
        let page = NSPrintOperation.current?.currentPage ?? Int(dirtyRect.midY / sheetSize.height) + 1
        printing.draw(page: page, in: rectForPage(page), context: context)
    }

    /// Each sheet is the paper's printable area: the margins are the printer's unprintable edges.
    private func fitSheets(to info: NSPrintInfo) {
        let paper = info.paperSize
        let printable = info.imageablePageBounds
        info.leftMargin = printable.minX
        info.bottomMargin = printable.minY
        info.rightMargin = max(paper.width - printable.maxX, 0)
        info.topMargin = max(paper.height - printable.maxY, 0)
        info.isHorizontallyCentered = true
        info.isVerticallyCentered = true
        info.horizontalPagination = .clip
        info.verticalPagination = .clip
        sheetSize = printable.size
        setFrameSize(NSSize(width: sheetSize.width, height: sheetSize.height * CGFloat(max(printing.pageCount, 1))))
    }
}

extension DocumentSession {
    /// Shows the system print panel for the document, as a sheet on its window.
    func print(from window: NSWindow?) {
        guard let printing = makePrinting() else { return }
        let info = (NSPrintInfo.shared.copy() as? NSPrintInfo) ?? NSPrintInfo()
        info.orientation = printing.startsLandscape ? .landscape : .portrait
        let operation = NSPrintOperation(view: PagePrintView(printing, printInfo: info), printInfo: info)
        operation.jobTitle = printing.jobName
        operation.showsPrintPanel = true
        operation.showsProgressPanel = true
        operation.printPanel.options.formUnion([.showsPaperSize, .showsOrientation, .showsPreview, .showsPageRange, .showsCopies])
        guard let window else {
            let succeeded = operation.run()
            if succeeded { printFinished(printing, error: nil) }
            return
        }
        let completion = PrintCompletion { [weak self] succeeded in
            if succeeded { self?.printFinished(printing, error: nil) }
        }
        operation.runModal(for: window, delegate: completion,
                           didRun: #selector(PrintCompletion.printOperationDidRun(_:success:contextInfo:)), contextInfo: nil)
    }
}

/// Receives the end of a print sheet, which AppKit reports through a selector on the main thread; it keeps itself
/// alive until then.
@MainActor
private final class PrintCompletion: NSObject {
    private let completion: @MainActor (Bool) -> Void
    private var retained: PrintCompletion?

    init(_ completion: @escaping @MainActor (Bool) -> Void) {
        self.completion = completion
        super.init()
        retained = self
    }

    @objc func printOperationDidRun(_ operation: NSPrintOperation, success: Bool, contextInfo: UnsafeMutableRawPointer?) {
        retained = nil
        completion(success)
    }
}
#endif
