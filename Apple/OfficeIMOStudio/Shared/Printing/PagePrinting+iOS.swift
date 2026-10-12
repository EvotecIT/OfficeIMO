#if os(iOS)
import UIKit

/// The pages for the system print sheet, which previews them, chooses the printer, paper, range and copies, and
/// prints in the background.
final class PagePrintRenderer: UIPrintPageRenderer {
    private let printing: PagePrinting

    init(_ printing: PagePrinting) {
        self.printing = printing
        super.init()
    }

    override var numberOfPages: Int { printing.pageCount }

    override func drawPage(at pageIndex: Int, in printableRect: CGRect) {
        guard let context = UIGraphicsGetCurrentContext() else { return }
        printing.draw(page: pageIndex + 1, in: printableRect, context: context)
    }
}

extension DocumentSession {
    /// Shows the system print sheet for the document, once a sheet or popover on its way out, such as the share
    /// sheet Print came from, has gone; UIKit ignores a presentation that starts during another's dismissal.
    func print(attempts: Int = 20) {
        let window = UIApplication.shared.connectedScenes.compactMap { ($0 as? UIWindowScene)?.keyWindow }.first
        if attempts > 0, window?.rootViewController?.presentedViewController != nil {
            DispatchQueue.main.asyncAfter(deadline: .now() + 0.1) { [weak self] in self?.print(attempts: attempts - 1) }
            return
        }
        guard let printing = makePrinting() else { return }
        let info = UIPrintInfo.printInfo()
        info.jobName = printing.jobName
        info.outputType = .general
        info.orientation = printing.startsLandscape ? .landscape : .portrait
        let controller = UIPrintInteractionController.shared
        controller.printInfo = info
        controller.printPageRenderer = PagePrintRenderer(printing)
        controller.present(animated: true) { [weak self] _, completed, error in
            guard completed || error != nil else { return }
            self?.printFinished(printing, error: error)
        }
    }
}

/// Print in the share sheet, drawing the engine's pages; it replaces the system's Print, which would print the saved
/// file with another renderer.
final class PrintActivity: UIActivity {
    private let action: @MainActor () -> Void

    init(_ action: @escaping @MainActor () -> Void) {
        self.action = action
        super.init()
    }

    override var activityType: UIActivity.ActivityType? { UIActivity.ActivityType("com.evotec.officeimo.studio.print") }
    override var activityTitle: String? { String(localized: "Print") }
    override var activityImage: UIImage? { UIImage(systemName: "printer") }
    override class var activityCategory: UIActivity.Category { .action }

    override func canPerform(withActivityItems activityItems: [Any]) -> Bool { true }

    override func perform() {
        activityDidFinish(true)
        // The print sheet follows once the share sheet has gone.
        let action = action
        Task { @MainActor in action() }
    }
}
#endif
