import SwiftUI

/// The tools for a document, by what the user wants to do, for the Tools menu on the Mac and the toolbar's Tools button.
struct StructureCommands: View {
    let session: DocumentSession

    var body: some View {
        Menu("Fill & Sign", systemImage: "signature") {
            SignatureMenu(session: session)
            Button("Sign with Certificate…", systemImage: "checkmark.seal.text.page") { session.securitySheet = .sign }
                .disabled(!session.canSignWithCertificate)
            Menu("Add Stamp", systemImage: "seal") {
                ForEach(StandardStamp.allCases) { stamp in
                    Button(stamp.title) { session.addStamp(stamp) }
                }
            }
            .disabled(!session.canEditAnnotations)
            Divider()
            Button("Verify Signatures…", systemImage: "list.bullet.rectangle") { session.structureSheet = .signatures }
                .disabled(!session.canCheck)
        }
        Menu("Protect", systemImage: "lock") {
            Button("Protect with Password…", systemImage: "lock") { session.securitySheet = .protect }
                .disabled(!session.canChangeSecurity)
            if session.isEncrypted {
                Button("Remove Password…", systemImage: "lock.open") { session.securitySheet = .removePassword }
                    .disabled(!session.canChangeSecurity)
            }
            Divider()
            Button("Redact…", systemImage: "rectangle.fill.badge.xmark") { session.structureSheet = .redact }
                .disabled(!session.canRedact)
            Button("Mark Areas to Redact", systemImage: "rectangle.dashed") { session.beginMarkingRedactions() }
                .disabled(!session.canRedact || session.redactionAreas != nil)
            Button("Flatten…", systemImage: "square.stack.3d.down.forward") { session.structureSheet = .flatten }
                .disabled(!session.canFlattenForms && !session.canFlattenMarkup)
        }
        Button("Convert to Word, Excel…", systemImage: "square.and.arrow.up.on.square") { session.structureSheet = .export }
            .disabled(!session.canExportDocument)
        Button("Compress PDF…", systemImage: "arrow.down.right.and.arrow.up.left") { session.structureSheet = .compress }
            .disabled(!session.canCheck)
        Menu("Organize Pages", systemImage: "square.grid.2x2") {
            Button("Split Document…", systemImage: "square.split.1x2") { session.structureSheet = .split }
                .disabled(!session.canSplit)
            Divider()
            Group {
                Button("Crop…", systemImage: "crop") { session.structureSheet = .crop }
                    .disabled(!session.canCrop)
                Button("Page Size…", systemImage: "arrow.up.left.and.arrow.down.right") { session.structureSheet = .pageSize }
                    .disabled(!session.canResize)
                Button("Header and Footer…", systemImage: "textformat.123") { session.structureSheet = .headerFooter }
                    .disabled(!session.canStampPages)
                Button("Watermark…", systemImage: "drop") { session.structureSheet = .watermark }
                    .disabled(!session.canStampPages)
                Button("Bates Numbering…", systemImage: "number") { session.structureSheet = .bates }
                    .disabled(!session.canStampPages)
            }
            .disabled(session.currentStructure == nil)
            Divider()
            Button("Add Bookmark", systemImage: "bookmark") { session.addBookmark() }
                .disabled(!session.canEditOutline)
            Button("Add Link…", systemImage: "link") { session.structureSheet = .link }
                .disabled(session.textSelection == nil || !session.canEditAnnotations)
        }
        Menu("Check", systemImage: "checkmark.seal") {
            Button("Make Searchable", systemImage: "text.viewfinder") { session.makeSearchable() }
                .disabled(!session.canMakeSearchable)
            Button("Check Standard…", systemImage: "checkmark.seal") { session.structureSheet = .standard }
                .disabled(!session.canCheck)
            Button("Document Health…", systemImage: "stethoscope") { session.structureSheet = .health }
                .disabled(!session.canCheck)
        }
    }
}
