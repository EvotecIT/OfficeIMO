import Foundation

/// Page organization and import, with temporary inputs retained only through their password retry.
extension DocumentSession {
    func rotate(_ pages: [Int], degrees: Int) {
        guard !pages.isEmpty else { return }
        edit(degrees < 0 ? String(localized: "Rotate Left") : String(localized: "Rotate Right")) { try await $0.rotatePages(pages, degrees: degrees) }
    }

    /// Deletes pages; a document keeps at least one.
    func delete(_ pages: [Int]) {
        guard !pages.isEmpty, pages.count < pageCount else { return }
        let remaining = pageCount - pages.count
        edit(pages.count == 1 ? String(localized: "Delete Page") : String(localized: "Delete Pages")) { try await $0.deletePages(pages) } then: { session in
            session.selection = []
            session.currentPage = min(session.currentPage, remaining)
        }
    }

    /// Duplicates pages; each copy follows its page, and the copies become the selection.
    func duplicate(_ pages: [Int]) {
        let pages = pages.sorted()
        guard !pages.isEmpty else { return }
        // Each copy lands after its page, shifted by the copies inserted before it.
        let copies = Set(pages.enumerated().map { index, page in page + index + 1 })
        edit(pages.count == 1 ? String(localized: "Duplicate Page") : String(localized: "Duplicate Pages")) { try await $0.duplicatePages(pages) } then: { session in
            session.selection = copies
        }
    }

    /// Moves pages, in their current order, before a page, or to the end with the page count + 1.
    func move(_ pages: [Int], before target: Int) {
        let moving = pages.sorted()
        let count = pageCount
        guard !moving.isEmpty, (1...count + 1).contains(target) else { return }
        let set = Set(moving)
        var order = (1...count).filter { !set.contains($0) }
        let position = order.firstIndex { $0 >= target } ?? order.count
        order.insert(contentsOf: moving, at: position)
        guard order != Array(1...count) else { return }
        let moved = Set((position + 1)...(position + moving.count))
        let newOrder = order
        edit(moving.count == 1 ? String(localized: "Move Page") : String(localized: "Move Pages")) { try await $0.reorderPages(newOrder) } then: { session in
            session.selection = moved
        }
    }

    /// Inserts a blank page the size of the page it follows, after the last target page.
    func insertBlankPage() {
        guard let facts, let after = targetPages.last, facts.pages.indices.contains(after - 1) else { return }
        let size = facts.pages[after - 1].displaySize
        edit(String(localized: "Insert Blank Page")) {
            try await $0.insertBlankPage(before: after + 1, width: size.width, height: size.height)
        } then: { session in
            session.selection = [after + 1]
            session.show(page: after + 1)
        }
    }

    /// Inserts the pages of files the user chose after the target pages, as one undoable change: PDFs, images and
    /// documents the engine converts, which become PDFs first.
    func importPages(from urls: [URL]) {
        guard canImportPages, !urls.isEmpty, let after = targetPages.last else { return }
        isRunningOperation = true
        Task {
            var sources: [ImportSource] = []
            do {
                for url in urls { sources.append(ImportSource(url: try await PageSources.pdf(for: url))) }
            } catch {
                for source in sources where !urls.contains(source.url) { StudioTemporaryFiles.removeGeneratedPDF(source.url) }
                isRunningOperation = false
                errorMessage = error.localizedDescription
                return
            }
            isRunningOperation = false
            importPages(sources, before: after + 1,
                temporaryFiles: ImportTemporaryFiles(sources.map(\.url).filter { !urls.contains($0) }))
        }
    }

    /// Supplies the password a pending import asked for and tries again.
    func continueImport(password: String) {
        guard var pending = pendingImport else { return }
        pendingImport = nil
        pending.sources[pending.needed.source].password = password
        importPages(pending.sources, before: pending.before, temporaryFiles: pending.temporaryFiles)
    }

    private func importPages(_ sources: [ImportSource], before page: Int, temporaryFiles: ImportTemporaryFiles) {
        let pagesBefore = pageCount
        edit(String(localized: "Insert Pages")) {
            _ = try await $0.importPages(sources, before: page)
        } then: { session in
            temporaryFiles.remove()
            let inserted = session.pageCount - pagesBefore
            guard inserted > 0 else { return }
            session.selection = Set(page..<(page + inserted))
            session.show(page: page)
        } failed: {
            temporaryFiles.remove()
        } passwordNeeded: { [weak self] needed in
            self?.pendingImport = PendingImport(sources: sources, before: page, needed: needed, temporaryFiles: temporaryFiles)
        }
    }

}

struct PendingImport: Identifiable {
    let id = UUID()
    var sources: [ImportSource]
    let before: Int
    let needed: ImportPasswordNeeded
    let temporaryFiles: ImportTemporaryFiles
}

/// Converted inputs live through password retries and are removed on success, failure, cancellation or window close.
final class ImportTemporaryFiles: Sendable {
    private let urls: [URL]
    init(_ urls: [URL]) { self.urls = urls }
    deinit { remove() }
    func remove() { for url in urls { StudioTemporaryFiles.removeGeneratedPDF(url) } }
}
