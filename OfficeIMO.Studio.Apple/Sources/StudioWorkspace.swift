import SwiftUI
import PDFKit

struct StudioWorkspace: View {
    @ObservedObject var document: StudioDocument
    let fileURL: URL?
    @Environment(\.undoManager) private var undoManager
    #if os(iOS)
    @Environment(\.horizontalSizeClass) private var sizeClass
    #endif
    @State private var selectedPage = 0
    @State private var visibility: NavigationSplitViewVisibility = .automatic
    @State private var showingPages = true
    @State private var showingCompactPages = false
    @State private var note = ""
    @State private var showingNote = false
    @State private var exporting = false
    @State private var busy = false
    @State private var noteTask: Task<Void, Never>?
    @State private var error: String?

    var body: some View {
        workspace
            .onDisappear { noteTask?.cancel() }
            .sheet(isPresented: $showingNote) { noteSheet }
            .sheet(isPresented: $showingCompactPages) {
                NavigationStack {
                    sidebar.navigationTitle("Pages and Notes")
                        .toolbar { ToolbarItem(placement: .confirmationAction) { Button("Done") { showingCompactPages = false } } }
                }
            }
            .fileExporter(isPresented: $exporting, document: PDFExport(data: document.data), contentType: .pdf,
                          defaultFilename: (fileURL?.deletingPathExtension().lastPathComponent ?? "Document") + " reviewed") { result in
                if case .failure(let failure) = result { error = failure.localizedDescription }
            }
            .alert("Unable to complete the action", isPresented: Binding(get: { error != nil }, set: { if !$0 { error = nil } })) {
                Button("OK", role: .cancel) { error = nil }
            } message: { Text(error ?? "") }
    }

    @ViewBuilder private var workspace: some View {
        #if os(macOS)
        NavigationSplitView(columnVisibility: $visibility) {
            sidebar.navigationTitle("Document")
                .navigationSplitViewColumnWidth(min: 180, ideal: 240, max: 340)
        } detail: { canvas }
        #else
        // DocumentGroup owns iOS navigation and the return-to-Files action.
        // Avoid nesting another navigation controller inside that document container.
        HStack(spacing: 0) {
            if sizeClass == .regular && showingPages {
                sidebar.frame(width: 220)
                Divider()
            }
            canvas
        }
        .toolbar {
            ToolbarItem(placement: .topBarLeading) {
                Button {
                    if sizeClass == .regular { showingPages.toggle() }
                    else { showingCompactPages = true }
                } label: { Label("Pages and Notes", systemImage: "sidebar.left") }
                .accessibilityIdentifier("showPages")
            }
        }
        #endif
    }

    private var sidebar: some View {
        List {
            Section("Pages") {
                ForEach(0..<document.pdf.pageCount, id: \.self) { index in
                    Button {
                        selectedPage = index
                        showingCompactPages = false
                    } label: {
                        HStack {
                            Label("Page \(index + 1)", systemImage: "doc.text")
                            Spacer()
                            if selectedPage == index { Image(systemName: "checkmark").accessibilityHidden(true) }
                        }
                        .padding(.vertical, 8).contentShape(Rectangle())
                    }
                    .buttonStyle(.plain)
                    .accessibilityAddTraits(selectedPage == index ? .isSelected : [])
                }
            }
            if !pageNotes.isEmpty {
                Section("Notes on page \(selectedPage + 1)") {
                    ForEach(Array(pageNotes.enumerated()), id: \.offset) { _, text in
                        Text(text).font(.callout).textSelection(.enabled).lineLimit(nil)
                            .fixedSize(horizontal: false, vertical: true).padding(.vertical, 6)
                    }
                }
            }
        }
        .listStyle(.sidebar)
    }

    private var canvas: some View {
        PDFViewport(document: document.pdf, page: $selectedPage)
            .accessibilityIdentifier("documentCanvas")
            .navigationTitle(fileURL?.deletingPathExtension().lastPathComponent ?? "Untitled")
            .safeAreaInset(edge: .bottom, spacing: 0) { pageControls.padding(.bottom, 12).padding(.top, 8) }
            .toolbar {
                ToolbarItemGroup(placement: .primaryAction) {
                    Button { showingNote = true } label: { Label("Add Note", systemImage: "square.and.pencil") }
                        .help("Add a note to the current page")
                        .accessibilityIdentifier("addNote").disabled(busy)
                    Button { exporting = true } label: { Label("Save a Copy", systemImage: "square.and.arrow.up") }
                        .accessibilityIdentifier("saveCopy").disabled(busy)
                }
            }
            .overlay(alignment: .top) {
                if busy { ProgressView("Adding note…").padding().glassEffect().padding() }
            }
    }

    private var pageNotes: [String] {
        (document.pdf.page(at: selectedPage)?.annotations ?? []).filter { $0.type == "Text" }.compactMap { $0.contents }
    }

    private var pageControls: some View {
        HStack(spacing: 20) {
            Button { selectedPage = max(0, selectedPage - 1) } label: {
                Image(systemName: "chevron.left").frame(width: 44, height: 44).contentShape(Rectangle())
            }.disabled(selectedPage == 0).accessibilityLabel("Previous page")
            Text("\(selectedPage + 1) of \(document.pdf.pageCount)").font(.subheadline.monospacedDigit())
            Button { selectedPage = min(document.pdf.pageCount - 1, selectedPage + 1) } label: {
                Image(systemName: "chevron.right").frame(width: 44, height: 44).contentShape(Rectangle())
            }.disabled(selectedPage >= document.pdf.pageCount - 1).accessibilityLabel("Next page")
        }
        .buttonStyle(.plain).padding(.horizontal, 22).frame(minHeight: 44)
        .glassEffect(.regular, in: .capsule).accessibilityElement(children: .contain)
    }

    private var noteSheet: some View {
        NavigationStack {
            Form {
                Section {
                    TextEditor(text: $note).frame(minHeight: 150).accessibilityLabel("Note text").accessibilityIdentifier("noteText")
                } header: { Text("Page \(selectedPage + 1)") } footer: { Text("Your note is placed near the bottom-left of this page.") }
            }
            .navigationTitle("Add a Note")
            .toolbar {
                ToolbarItem(placement: .cancellationAction) { Button("Cancel") { note = ""; showingNote = false } }
                ToolbarItem(placement: .confirmationAction) {
                    Button("Add") { addNote() }.disabled(note.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty || note.utf16.count > 4000)
                        .accessibilityIdentifier("confirmNote")
                }
            }
        }
        #if os(macOS)
        .frame(width: 440, height: 330)
        #else
        .presentationDetents([.medium, .large])
        #endif
    }

    private func addNote() {
        let original = document.data
        let page = selectedPage + 1
        let text = note
        showingNote = false
        busy = true
        noteTask = Task { @MainActor in
            defer { busy = false }
            do {
                let updated = try await Task.detached(priority: .userInitiated) {
                    try StudioEngine.addNote(to: original, page: page, text: text)
                }.value
                try Task.checkCancellation()
                guard document.data == original else {
                    throw StudioEngine.EngineError(message: "The document changed while the note was being added. Please try again.")
                }
                try document.replace(with: updated, undoManager: undoManager)
                note = ""
            } catch is CancellationError {
                // A closed workspace must not receive a late document edit.
            } catch {
                self.error = error.localizedDescription + " Your note has been kept. Choose Add Note to try again."
            }
        }
    }
}
