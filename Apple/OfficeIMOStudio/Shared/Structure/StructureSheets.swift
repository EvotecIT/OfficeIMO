import SwiftUI

/// The sheet for a StructureSheet. Page changes apply as one undoable change; outputs are offered for saving once
/// the sheet closes.
struct StructureSheetView: View {
    let session: DocumentSession
    let sheet: StructureSheet

    var body: some View {
        switch sheet {
        case .crop: CropSheet(session: session)
        case .pageSize: PageSizeSheet(session: session)
        case .headerFooter: HeaderFooterSheet(session: session)
        case .watermark: WatermarkSheet(session: session)
        case .bates: BatesSheet(session: session)
        case .split: SplitSheet(session: session)
        case .export: ExportSheet(session: session)
        case .flatten: FlattenSheet(session: session)
        case .signatures: SignaturesSheet(session: session)
        case .standard: StandardSheet(session: session)
        case .health: HealthSheet(session: session)
        case .redact: RedactSheet(session: session)
        case .link: LinkSheet(session: session)
        }
    }
}

/// Which pages a page change applies to.
enum PageScope: String, CaseIterable, Identifiable {
    case current, selected, all

    var id: Self { self }

    @MainActor func pages(_ session: DocumentSession) -> [Int] {
        switch self {
        case .current: [session.currentPage]
        case .selected: session.selection.sorted()
        case .all: Array(1...max(session.pageCount, 1))
        }
    }

    @MainActor func title(_ session: DocumentSession) -> LocalizedStringKey {
        switch self {
        case .current: "Page \(session.currentPage)"
        case .selected: "Selected Pages (\(session.selection.count))"
        case .all: "All Pages (\(session.pageCount))"
        }
    }

    @MainActor static func available(_ session: DocumentSession) -> [PageScope] {
        session.task == .pages && !session.selection.isEmpty ? [.selected, .all] : [.current, .all]
    }

    @MainActor static func initial(_ session: DocumentSession) -> PageScope {
        session.task == .pages && !session.selection.isEmpty ? .selected : .current
    }
}

private struct PageScopePicker: View {
    let session: DocumentSession
    @Binding var scope: PageScope

    var body: some View {
        Picker("Pages", selection: $scope) {
            ForEach(PageScope.available(session)) { scope in
                Text(scope.title(session)).tag(scope)
            }
        }
    }
}

/// A length shown in millimetres, or inches where the US measurement system is used, and kept in points.
private struct LengthField: View {
    let title: LocalizedStringKey
    @Binding var points: Double

    private static let isMetric = Locale.current.measurementSystem != .us
    private static let pointsPerUnit = isMetric ? 72 / 25.4 : 72

    var body: some View {
        LabeledContent(title) {
            HStack(spacing: 4) {
                TextField(title, value: Binding { points / Self.pointsPerUnit } set: { points = max(0, $0) * Self.pointsPerUnit },
                          format: .number.precision(.fractionLength(0...(Self.isMetric ? 1 : 2))))
                    .labelsHidden()
                    .multilineTextAlignment(.trailing)
                    #if os(iOS)
                    .keyboardType(.decimalPad)
                    #endif
                    .frame(maxWidth: 90)
                Text(Self.isMetric ? "mm" : "in").foregroundStyle(.secondary)
            }
        }
    }
}

/// Crop: margins cut from each edge of the pages as displayed, as Preview's crop keeps the selected area.
private struct CropSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var scope = PageScope.current
    @State private var top = 36.0
    @State private var bottom = 36.0
    @State private var left = 36.0
    @State private var right = 36.0

    var body: some View {
        SheetForm(title: "Crop Pages", confirm: "Crop", isComplete: top + bottom + left + right > 0, isWriting: false, problem: nil,
                  cancel: { dismiss() }, confirmAction: apply) {
            Section {
                PageScopePicker(session: session, scope: $scope)
            }
            Section {
                LengthField(title: "Top", points: $top)
                LengthField(title: "Bottom", points: $bottom)
                LengthField(title: "Left", points: $left)
                LengthField(title: "Right", points: $right)
            } header: {
                Text("Remove from Each Edge")
            } footer: {
                Text("Cropping hides the edges; the content stays in the document.")
            }
        }
        .onAppear { scope = PageScope.initial(session) }
    }

    private func apply() {
        session.crop(scope.pages(session), left: left, top: top, right: right, bottom: bottom)
        dismiss()
    }
}

/// Page Size: a standard paper size, with the content fitted, filling or stretched to it.
private struct PageSizeSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var scope = PageScope.current
    @State private var paper = Paper.defaultPaper
    @State private var isLandscape = false
    @State private var mode = PageSizeEdit.Mode.fit

    enum Paper: String, CaseIterable, Identifiable {
        case a3, a4, a5, letter, legal, tabloid

        var id: Self { self }
        var size: CGSize {
            switch self {
            case .a3: CGSize(width: 842, height: 1191)
            case .a4: CGSize(width: 595, height: 842)
            case .a5: CGSize(width: 420, height: 595)
            case .letter: CGSize(width: 612, height: 792)
            case .legal: CGSize(width: 612, height: 1008)
            case .tabloid: CGSize(width: 792, height: 1224)
            }
        }
        var title: LocalizedStringKey {
            switch self {
            case .a3: "A3"
            case .a4: "A4"
            case .a5: "A5"
            case .letter: "US Letter"
            case .legal: "US Legal"
            case .tabloid: "Tabloid"
            }
        }
        static var defaultPaper: Paper { Locale.current.measurementSystem == .us ? .letter : .a4 }
    }

    var body: some View {
        SheetForm(title: "Page Size", confirm: "Change", isComplete: true, isWriting: false, problem: nil,
                  cancel: { dismiss() }, confirmAction: apply) {
            Section {
                PageScopePicker(session: session, scope: $scope)
            }
            Section {
                Picker("Paper Size", selection: $paper) {
                    ForEach(Paper.allCases) { paper in
                        Text("\(Text(paper.title)) – \(DocumentInspector.paperSize(paper.size))").tag(paper)
                    }
                }
                Picker("Orientation", selection: $isLandscape) {
                    Text("Portrait").tag(false)
                    Text("Landscape").tag(true)
                }
                .pickerStyle(.segmented)
                Picker("Content", selection: $mode) {
                    Text("Fit Inside").tag(PageSizeEdit.Mode.fit)
                    Text("Fill the Page").tag(PageSizeEdit.Mode.fill)
                    Text("Stretch").tag(PageSizeEdit.Mode.stretch)
                }
            } footer: {
                Text(mode == .fit ? "The content keeps its proportions and fits inside the new page."
                     : mode == .fill ? "The content keeps its proportions and covers the new page; edges may be cut."
                     : "The content stretches to the new page’s width and height.")
            }
        }
        .onAppear { scope = PageScope.initial(session) }
    }

    private func apply() {
        let size = paper.size
        session.resize(scope.pages(session), width: isLandscape ? size.height : size.width, height: isLandscape ? size.width : size.height,
                       mode: mode)
        dismiss()
    }
}

/// Header and Footer: text at the top and bottom of the pages, with page numbers and other fields filled in for each.
private struct HeaderFooterSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var scope = PageScope.all
    @State private var headerLeft = ""
    @State private var headerCenter = ""
    @State private var headerRight = ""
    @State private var footerLeft = ""
    @State private var footerCenter = String(localized: "Page {page} of {pages}")
    @State private var footerRight = ""
    @State private var fontSize = 9.0
    @State private var startNumber = 1

    private var isComplete: Bool {
        ![headerLeft, headerCenter, headerRight, footerLeft, footerCenter, footerRight]
            .allSatisfy { $0.trimmingCharacters(in: .whitespaces).isEmpty }
    }

    var body: some View {
        SheetForm(title: "Header and Footer", confirm: "Add", isComplete: isComplete, isWriting: false, problem: nil,
                  cancel: { dismiss() }, confirmAction: apply) {
            Section {
                PageScopePicker(session: session, scope: $scope)
            }
            Section("Header") {
                TextField("Left", text: $headerLeft)
                TextField("Center", text: $headerCenter)
                TextField("Right", text: $headerRight)
            }
            Section {
                TextField("Left", text: $footerLeft)
                TextField("Center", text: $footerCenter)
                TextField("Right", text: $footerRight)
            } header: {
                Text("Footer")
            } footer: {
                Text("{page} and {pages} become the page number and count, {date} today’s date, {file} the file name and {title} the document’s title.")
            }
            Section {
                Stepper("First Page Number: \(startNumber)", value: $startNumber, in: 0...99_999)
                Stepper("Text Size: \(Int(fontSize)) pt", value: $fontSize, in: 6...24)
            }
        }
    }

    private func apply() {
        func text(_ value: String) -> String? {
            let trimmed = value.trimmingCharacters(in: .whitespaces)
            return trimmed.isEmpty ? nil : trimmed
        }
        session.addHeaderFooter(HeaderFooterEdit(headerLeft: text(headerLeft), headerCenter: text(headerCenter), headerRight: text(headerRight),
                                                 footerLeft: text(footerLeft), footerCenter: text(footerCenter), footerRight: text(footerRight),
                                                 fontSize: fontSize, startNumber: startNumber,
                                                 pages: scope == .all ? nil : scope.pages(session)))
        dismiss()
    }
}

/// Watermark: text across the pages, over or behind their content.
private struct WatermarkSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var scope = PageScope.all
    @State private var text = String(localized: "CONFIDENTIAL")
    @State private var fontSize = 48.0
    @State private var opacity = 0.3
    @State private var rotation = -35.0
    @State private var behindContent = false
    @State private var color = Color.gray

    var body: some View {
        SheetForm(title: "Watermark", confirm: "Add", isComplete: !text.trimmingCharacters(in: .whitespaces).isEmpty, isWriting: false,
                  problem: nil, cancel: { dismiss() }, confirmAction: apply) {
            Section {
                PageScopePicker(session: session, scope: $scope)
            }
            Section {
                TextField("Text", text: $text)
                ColorPicker("Color", selection: $color, supportsOpacity: false)
                Stepper("Text Size: \(Int(fontSize)) pt", value: $fontSize, in: 12...144, step: 4)
                LabeledContent("Opacity") {
                    Slider(value: $opacity, in: 0.05...1) { Text("Opacity") }
                        .labelsHidden()
                }
                LabeledContent("Angle") {
                    Slider(value: $rotation, in: -90...90, step: 5) { Text("Angle") }
                        .labelsHidden()
                }
                Toggle("Behind the Content", isOn: $behindContent)
            }
        }
    }

    private func apply() {
        let resolved = color.resolve(in: EnvironmentValues())
        session.addWatermark(WatermarkEdit(text: text.trimmingCharacters(in: .whitespaces), fontSize: fontSize, opacity: opacity,
                                           rotation: rotation, behindContent: behindContent,
                                           color: EngineColor(red: Double(resolved.red), green: Double(resolved.green), blue: Double(resolved.blue)),
                                           pages: scope == .all ? nil : scope.pages(session)))
        dismiss()
    }
}

/// Bates Numbering: a running number on every page, as legal productions use.
private struct BatesSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var prefix = ""
    @State private var suffix = ""
    @State private var start: Int64 = 1
    @State private var digits = 6
    @State private var position = BatesEdit.Position.bottomRight

    private var example: String {
        prefix + String(repeating: "0", count: max(0, digits - String(start).count)) + String(start) + suffix
    }

    var body: some View {
        SheetForm(title: "Bates Numbering", confirm: "Add", isComplete: true, isWriting: false, problem: nil,
                  cancel: { dismiss() }, confirmAction: apply) {
            Section {
                TextField("Prefix", text: $prefix)
                TextField("Suffix", text: $suffix)
                LabeledContent("Start At") {
                    TextField("Start At", value: $start, format: .number.grouping(.never))
                        .labelsHidden()
                        .multilineTextAlignment(.trailing)
                        #if os(iOS)
                        .keyboardType(.numberPad)
                        #endif
                }
                Stepper("Digits: \(digits)", value: $digits, in: 1...15)
                Picker("Position", selection: $position) {
                    Text("Top Left").tag(BatesEdit.Position.topLeft)
                    Text("Top Center").tag(BatesEdit.Position.topCenter)
                    Text("Top Right").tag(BatesEdit.Position.topRight)
                    Text("Bottom Left").tag(BatesEdit.Position.bottomLeft)
                    Text("Bottom Center").tag(BatesEdit.Position.bottomCenter)
                    Text("Bottom Right").tag(BatesEdit.Position.bottomRight)
                }
            } footer: {
                Text("The first page shows \(example). Every page is numbered.")
            }
        }
    }

    private func apply() {
        session.addBatesNumbers(BatesEdit(prefix: prefix, suffix: suffix, start: max(0, start), digits: digits, position: position))
        dismiss()
    }
}

/// Split: the document as several PDFs, every few pages or one per top-level bookmark, saved together in a folder.
private struct SplitSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var atBookmarks = false
    @State private var pagesPerPart = 1
    @State private var problem: String?
    @State private var isWriting = false

    private var bookmarkParts: [SplitPart] {
        OutlineNode.parts(session.currentStructure?.outlineNodes ?? [], pageCount: session.pageCount)
    }
    private var partCount: Int {
        atBookmarks ? bookmarkParts.count : (session.pageCount + pagesPerPart - 1) / max(pagesPerPart, 1)
    }

    var body: some View {
        SheetForm(title: "Split Document", confirm: "Split", isComplete: partCount >= 2, isWriting: isWriting, problem: problem,
                  cancel: { dismiss() }, confirmAction: split) {
            Section {
                Picker("Split", selection: $atBookmarks) {
                    Text("Every Few Pages").tag(false)
                    Text("At Each Bookmark").tag(true)
                }
                .pickerStyle(.inline)
                .labelsHidden()
                if !atBookmarks {
                    Stepper("Pages in Each: \(pagesPerPart)", value: $pagesPerPart, in: 1...max(1, session.pageCount - 1))
                }
            } footer: {
                if atBookmarks && bookmarkParts.count < 2 {
                    Text("The document needs at least two bookmarks on different pages, or one after the first page.")
                } else {
                    Text("Makes \(partCount) PDFs in a folder you choose. This document stays as it is.")
                }
            }
        }
        .task { await session.loadStructure() }
    }

    private func split() {
        problem = nil
        isWriting = true
        Task {
            defer { isWriting = false }
            do {
                if atBookmarks {
                    try await session.split(parts: bookmarkParts)
                } else {
                    try await session.split(pagesPerPart: pagesPerPart)
                }
                dismiss()
            } catch {
                problem = error.localizedDescription
            }
        }
    }
}

/// Export: the document as a Word, Excel or PowerPoint file, a web page, its text, or images of its pages or the
/// pictures inside it. A conversion that couldn't keep everything lists what it left out before the file is saved.
private struct ExportSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var format = ExportFormat.word
    @State private var converted: ExportedFile?
    @State private var problem: String?
    @State private var isWriting = false

    var body: some View {
        SheetForm(title: "Export", confirm: converted == nil ? "Export" : "Save…", isComplete: converted != nil || session.canExport(format), isWriting: isWriting, problem: problem,
                  cancel: cancel, confirmAction: export) {
            if let converted, let notes = converted.notes {
                Section {
                    ForEach(notes.split(separator: "\n").map(String.init), id: \.self) { note in
                        Label(note, systemImage: "exclamationmark.triangle")
                    }
                } header: {
                    Text("Converted with Notes")
                } footer: {
                    Text("Some content couldn’t be kept exactly as it is in the PDF. Save the file, or cancel to keep the document only.")
                }
            } else {
                Section("Document") { options(ExportFormat.documents) }
                Section("Text") { options(ExportFormat.texts) }
                Section {
                    options(ExportFormat.pictures)
                } header: {
                    Text("Images")
                } footer: {
                    Text(format.detail)
                }
            }
        }
        .onChange(of: format) { problem = nil }
    }

    private func options(_ formats: [ExportFormat]) -> some View {
        Picker("Format", selection: $format) {
            ForEach(formats) { format in
                Text(format.title).tag(format)
                    .disabled(!session.canExport(format))
            }
        }
        .pickerStyle(.inline)
        .labelsHidden()
    }

    private func export() {
        if let converted {
            session.pendingExportedFile = converted
            dismiss()
            return
        }
        problem = nil
        isWriting = true
        Task {
            defer { isWriting = false }
            do {
                guard let file = try await session.export(format) else { return }
                if file.notes == nil {
                    session.pendingExportedFile = file
                    dismiss()
                } else {
                    converted = file
                }
            } catch {
                problem = error.localizedDescription
            }
        }
    }

    private func cancel() {
        if let converted { try? FileManager.default.removeItem(at: converted.folder) }
        dismiss()
    }
}

/// Flatten: the form's current values, or the markup, become part of the pages, so they can no longer be changed.
private struct FlattenSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss
    @State private var target = FlattenTarget.form

    enum FlattenTarget: Hashable {
        case form, markup
    }

    var body: some View {
        SheetForm(title: "Flatten", confirm: "Flatten", isComplete: isAvailable(target), isWriting: false, problem: nil,
                  cancel: { dismiss() }, confirmAction: {
                      if target == .form { session.flattenForm() } else { session.flattenMarkup() }
                      dismiss()
                  }) {
            Section {
                Picker("Flatten", selection: $target) {
                    Text("Form Fields").tag(FlattenTarget.form)
                        .selectionDisabled(!session.canFlattenForms)
                    Text("Markup").tag(FlattenTarget.markup)
                        .selectionDisabled(!session.canFlattenMarkup)
                }
                .pickerStyle(.inline)
                .labelsHidden()
            } footer: {
                Text(target == .form
                     ? "The form’s current values become part of the pages, and its fields can no longer be filled or changed. Undo restores the form."
                     : "Drawings, shapes, highlights, notes and stamps become part of the pages and can no longer be moved or changed. Links stay. Undo restores the markup.")
            }
        }
        .onAppear {
            if !session.canFlattenForms, session.canFlattenMarkup { target = .markup }
        }
    }

    private func isAvailable(_ target: FlattenTarget) -> Bool {
        target == .form ? session.canFlattenForms : session.canFlattenMarkup
    }
}
