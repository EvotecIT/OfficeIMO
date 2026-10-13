import SwiftUI

/// The page production and output commands, for the Tools menu on the Mac and the iPad and the task menu on touch.
struct StructureCommands: View {
    let session: DocumentSession

    var body: some View {
        Menu("Page Layout", systemImage: "rectangle.and.pencil.and.ellipsis") {
            Button("Crop…", systemImage: "crop") { session.structureSheet = .crop }
                .disabled(!session.canCrop)
            Button("Page Size…", systemImage: "arrow.up.left.and.arrow.down.right") { session.structureSheet = .pageSize }
                .disabled(!session.canResize)
            Divider()
            Button("Header and Footer…", systemImage: "textformat.123") { session.structureSheet = .headerFooter }
            Button("Watermark…", systemImage: "drop") { session.structureSheet = .watermark }
            Button("Bates Numbering…", systemImage: "number") { session.structureSheet = .bates }
        }
        .disabled(!session.isReady || session.currentStructure == nil)
        .disabled(!session.canStampPages && !session.canCrop && !session.canResize)
        Button("Add Bookmark", systemImage: "bookmark") { session.addBookmark() }
            .disabled(!session.canEditOutline)
        Button("Flatten Form…", systemImage: "square.stack.3d.down.forward") { session.structureSheet = .flatten }
            .disabled(!session.canFlattenForms)
        Divider()
        Button("Split Document…", systemImage: "square.split.1x2") { session.structureSheet = .split }
            .disabled(!session.canSplit)
        Button("Export Text…", systemImage: "doc.plaintext") { session.structureSheet = .export }
            .disabled(!session.canExportText)
    }
}
