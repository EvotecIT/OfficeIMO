import PDFKit
import SwiftUI

/// The window of a viewed Word, Excel, PowerPoint or other document: its pages as converted to PDF, read-only, with
/// Edit as PDF to work on them in a PDF window and Save as PDF to keep a copy.
struct ViewerWindow: View {
    @State private var model: ViewerModel
    @State private var isExporting = false
    @State private var exportError: String?
    @Environment(\.dismiss) private var dismiss
    #if os(macOS)
    @Environment(\.newDocument) private var newDocument
    #endif

    init(document: ViewerDocument) {
        _model = State(initialValue: ViewerModel(document: document))
    }

    var body: some View {
        content
            .frame(maxWidth: .infinity, maxHeight: .infinity)
            .task { await model.load() }
            .onDisappear { model.close() }
            .toolbar {
                ToolbarItemGroup(placement: .primaryAction) {
                    Group {
                        Button("Edit as PDF", systemImage: "pencil.and.outline", action: editAsPDF)
                        Button("Save as PDF…", systemImage: "square.and.arrow.down") { isExporting = true }
                    }
                    .disabled(model.pdf == nil)
                }
            }
            .fileExporter(isPresented: $isExporting, document: model.pdf.map(ConvertedPDF.init), contentType: .pdf,
                          defaultFilename: model.pdfName) { result in
                if case .failure(let error) = result, (error as? CocoaError)?.code != .userCancelled {
                    exportError = error.localizedDescription
                }
            }
            .alert("Couldn’t Save the PDF", isPresented: Binding { exportError != nil } set: { if !$0 { exportError = nil } }) {
                Button("OK") {}
            } message: {
                Text(exportError ?? "")
            }
            #if os(macOS)
            .frame(minWidth: 480, minHeight: 360)
            #endif
    }

    @ViewBuilder private var content: some View {
        switch model.phase {
        case .converting:
            ProgressView("Preparing “\(model.document.name)”…")
        case .failed(let message):
            ContentUnavailableView {
                Label("Can’t Show This Document", systemImage: "doc.questionmark")
            } description: {
                Text(message)
            }
        case .ready(let pages, _, let differs):
            ConvertedPagesView(document: pages)
                .ignoresSafeArea(edges: .bottom)
                .safeAreaInset(edge: .bottom) {
                    Text(differs ? "A PDF copy. Some content looks different from the original." : "A PDF copy of the original.")
                    .font(.footnote)
                    .lineLimit(2)
                    .foregroundStyle(.secondary)
                    .multilineTextAlignment(.center)
                    .padding(.horizontal)
                    .padding(.vertical, 8)
                    .frame(maxWidth: .infinity)
                    .background(.bar)
                }
        }
    }

    private func editAsPDF() {
        guard let pdf = model.pdf else { return }
        #if os(macOS)
        do {
            let document = try StudioDocument(contentsOf: pdf)
            newDocument { document }
        } catch {
            exportError = error.localizedDescription
        }
        #else
        // The PDF opens from the document browser once this window has closed, as the launch screen's Convert does,
        // so the browser never shows this document's title over the PDF.
        let copy = StudioTemporaryFiles.root.appending(path: "Converted/\(UUID().uuidString)", directoryHint: .isDirectory)
            .appending(path: pdf.lastPathComponent)
        do {
            try FileManager.default.createDirectory(at: copy.deletingLastPathComponent(), withIntermediateDirectories: true)
            try FileManager.default.copyItem(at: pdf, to: copy)
        } catch {
            exportError = error.localizedDescription
            return
        }
        dismiss()
        LaunchDocuments.open { copy }
        #endif
    }
}

/// The converted pages in the system's PDF view: scrolling, zooming and text selection, nothing that edits.
#if os(macOS)
private struct ConvertedPagesView: NSViewRepresentable {
    let document: PDFDocument

    func makeNSView(context: Context) -> PDFView {
        let view = PDFView()
        configure(view)
        return view
    }

    func updateNSView(_ view: PDFView, context: Context) {
        if view.document !== document { view.document = document }
    }

    private func configure(_ view: PDFView) {
        view.autoScales = true
        view.displayMode = .singlePageContinuous
        view.displaysPageBreaks = true
        view.document = document
    }
}
#else
private struct ConvertedPagesView: UIViewRepresentable {
    let document: PDFDocument

    func makeUIView(context: Context) -> PDFView {
        let view = PDFView()
        view.autoScales = true
        view.displayMode = .singlePageContinuous
        view.displaysPageBreaks = true
        view.document = document
        return view
    }

    func updateUIView(_ view: PDFView, context: Context) {
        if view.document !== document { view.document = document }
    }
}
#endif
