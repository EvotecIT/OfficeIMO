import SwiftUI
import UniformTypeIdentifiers

/// The page commands for a set of pages, used by the toolbar for the selection. Each is one undoable change except
/// extraction, which writes a new PDF.
struct PageActionButtons: View {
    let session: DocumentSession
    let pages: [Int]

    var body: some View {
        Menu {
            Button("Rotate Left", systemImage: "rotate.left") { session.rotate(pages, degrees: -90) }
            Button("Rotate Right", systemImage: "rotate.right") { session.rotate(pages, degrees: 90) }
        } label: {
            Label("Rotate", systemImage: "rotate.right")
        } primaryAction: {
            session.rotate(pages, degrees: 90)
        }
        .disabled(pages.isEmpty || !session.canEditPages)
        .help("Rotate Right; hold for Rotate Left")

        Button("Duplicate", systemImage: "plus.square.on.square") { session.duplicate(pages) }
            .disabled(pages.isEmpty || !session.canEditPages)

        Button("Delete", systemImage: "trash", role: .destructive) { session.delete(pages) }
            .disabled(pages.isEmpty || pages.count >= session.pageCount || !session.canEditPages)

        Menu("More", systemImage: "ellipsis") {
            Button("Insert Blank Page", systemImage: "doc.badge.plus") { session.insertBlankPage() }
                .disabled(!session.canEditPages)
            Button("Extract Pages…", systemImage: "doc.on.doc") { session.extract(pages) }
                .disabled(pages.isEmpty || !session.canExtractPages)
        }
    }
}

/// The same commands for one page's context menu: the selection when the page is selected, otherwise that page.
struct PageContextMenu: View {
    let session: DocumentSession
    let page: Int

    private var pages: [Int] {
        session.selection.contains(page) ? session.selection.sorted() : [page]
    }

    var body: some View {
        Button("Rotate Left", systemImage: "rotate.left") { session.rotate(pages, degrees: -90) }
        Button("Rotate Right", systemImage: "rotate.right") { session.rotate(pages, degrees: 90) }
        Divider()
        Button("Duplicate", systemImage: "plus.square.on.square") { session.duplicate(pages) }
        Button("Insert Blank Page After", systemImage: "doc.badge.plus") {
            session.selection = [pages.last ?? page]
            session.insertBlankPage()
        }
        Button("Extract Pages…", systemImage: "doc.on.doc") { session.extract(pages) }
            .disabled(!session.canExtractPages)
        Divider()
        Button("Delete", systemImage: "trash", role: .destructive) { session.delete(pages) }
            .disabled(pages.count >= session.pageCount)
    }
}

/// Pages dragged within a window's page grid. The window's identity keeps a drop from another document out.
struct PageDrag: Codable, Transferable {
    let window: UUID
    let pages: [Int]

    static var transferRepresentation: some TransferRepresentation {
        CodableRepresentation(contentType: .officeIMOPages)
    }
}

extension UTType {
    /// Declared in Info.plist; only this app reads it.
    static let officeIMOPages = UTType(exportedAs: "com.evotec.officeimo.studio.pages")
}
