import Foundation

extension DocumentSession {
    /// Whether the document can be printed now: the document permits it, and no drawing waits to be added, which
    /// would otherwise be left off the paper.
    var canPrint: Bool { isReady && !isWorking && !hasPendingMarkup && facts?.canPrint != false }

    /// The pages of the revision shown, for the system's print panel or sheet.
    func makePrinting() -> PagePrinting? {
        guard canPrint, let engine, let facts else { return nil }
        let name = (facts.fileName as NSString).deletingPathExtension
        return PagePrinting(document: engine, facts: facts, jobName: name.isEmpty ? String(localized: "Document") : name)
    }

    /// Reports a print job that failed or printed pages blank.
    func printFinished(_ printing: PagePrinting, error: (any Error)?) {
        if let error {
            errorMessage = error.localizedDescription
        } else if !printing.failedPages.isEmpty {
            let pages = printing.failedPages.map(String.init).joined(separator: ", ")
            errorMessage = String(localized: "These pages could not be drawn and printed blank: \(pages).")
        }
    }
}
