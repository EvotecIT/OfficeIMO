import SwiftUI

/// A document's window or scene. In regular width, and on the Mac, the page list is the split view's sidebar and the
/// facts are the inspector; in compact width the page fills the screen and both are sheets, as the system presents them.
struct DocumentWindow: View {
    @State private var session: DocumentSession
    @State private var password = ""
    @Environment(\.undoManager) private var undoManager
    #if os(iOS)
    @Environment(\.horizontalSizeClass) private var sizeClass
    #endif
    let fileURL: URL?
    /// A title for windows the system's document group does not title.
    var title: String?

    init(document: StudioDocument, fileURL: URL?, title: String? = nil) {
        _session = State(initialValue: DocumentSession(document: document))
        self.fileURL = fileURL
        self.title = title
    }

    private var isCompact: Bool {
        #if os(iOS)
        sizeClass == .compact
        #else
        false
        #endif
    }

    var body: some View {
        layout
            .focusedSceneValue(\.documentSession, session)
            .task {
                if session.isPageListShown == nil { session.isPageListShown = !isCompact }
                session.undoManager = undoManager
                await session.load()
            }
            .onChange(of: undoManager) { session.undoManager = undoManager }
            .alert("Password Required", isPresented: passwordPrompt) {
                SecureField("Password", text: $password)
                Button("Open") {
                    let entered = password
                    password = ""
                    Task { await session.load(password: entered) }
                }
                Button("Cancel", role: .cancel) { password = "" }
            } message: {
                if case .needsPassword(retry: true) = session.state {
                    Text("The password is incorrect. Try again.")
                } else {
                    Text("This document is protected. Enter its password to open it.")
                }
            }
            .alert("The Change Couldn’t Be Made", isPresented: errorShown) {
                Button("OK") { session.errorMessage = nil }
            } message: {
                Text(session.errorMessage ?? "")
            }
    }

    @ViewBuilder private var layout: some View {
        if isCompact {
            column
                .sheet(isPresented: pageListShown) {
                    PageListSheet(session: session)
                }
        } else {
            NavigationSplitView(columnVisibility: columnVisibility) {
                PageSidebar(session: session)
                    .navigationSplitViewColumnWidth(min: 150, ideal: 180, max: 260)
            } detail: {
                column
            }
        }
    }

    /// The page column with its bars and inspector, which belong to the column rather than the whole split view.
    private var column: some View {
        detail
            .inspector(isPresented: $session.isInspectorShown) {
                DocumentInspector(session: session)
                    .inspectorColumnWidth(min: 240, ideal: 280, max: 360)
                    // In compact width the inspector is a sheet over the page.
                    .presentationDetents([.medium, .large])
                    .presentationDragIndicator(.visible)
                    .presentationBackgroundInteraction(.enabled(upThrough: .medium))
            }
            .toolbar { DocumentToolbar(session: session, fileURL: fileURL, isCompact: isCompact) }
            .modifier(OptionalTitle(title: title))
            #if os(iOS)
            .toolbarTitleDisplayMode(.inline)
            #endif
    }

    @ViewBuilder private var detail: some View {
        switch session.state {
        case .loading, .needsPassword:
            ProgressView()
                .frame(maxWidth: .infinity, maxHeight: .infinity)
        case .failed(let message):
            ContentUnavailableView {
                Label("Can’t Open This Document", systemImage: "exclamationmark.triangle")
            } description: {
                Text(message)
            }
        case .ready:
            switch session.task {
            case .read: ReadView(session: session)
            case .pages: PagesTaskView(session: session, isCompact: isCompact)
            }
        }
    }

    /// The Pages task shows every page itself, so the page list steps aside while it is open.
    private var columnVisibility: Binding<NavigationSplitViewVisibility> {
        Binding {
            session.task == .pages || session.isPageListShown == false ? .detailOnly : .all
        } set: {
            if session.task == .read { session.isPageListShown = $0 != .detailOnly }
        }
    }

    private var pageListShown: Binding<Bool> {
        Binding { session.isPageListShown == true } set: { session.isPageListShown = $0 }
    }

    private var passwordPrompt: Binding<Bool> {
        Binding {
            if case .needsPassword = session.state { true } else { false }
        } set: { _ in }
    }

    private var errorShown: Binding<Bool> {
        Binding { session.errorMessage != nil } set: { if !$0 { session.errorMessage = nil } }
    }
}

/// The Read task: the pages, with the page number floating over them while they scroll.
struct ReadView: View {
    let session: DocumentSession
    @State private var isIndicatorShown = false

    var body: some View {
        PageCanvas(session: session)
            #if os(macOS)
            // The scroll view runs under the toolbar and insets itself for it; the sidebar keeps its own column.
            .ignoresSafeArea(edges: .vertical)
            #else
            // The scroll view runs under the bars and the floating sidebar and insets the pages by the safe area.
            .ignoresSafeArea()
            #endif
            .overlay(alignment: .bottom) {
                if isIndicatorShown {
                    Text("\(session.currentPage) of \(session.pageCount)")
                        .font(.callout.weight(.medium))
                        .monospacedDigit()
                        .padding(.horizontal, 14)
                        .padding(.vertical, 8)
                        .glassEffect(.regular, in: .capsule)
                        .padding(.bottom, 12)
                        .transition(.opacity)
                        .accessibilityHidden(true)
                }
            }
            .task(id: session.currentPage) {
                guard session.pageCount > 1 else { return }
                withAnimation(.easeOut(duration: 0.15)) { isIndicatorShown = true }
                try? await Task.sleep(for: .seconds(1.5))
                withAnimation(.easeIn(duration: 0.3)) { isIndicatorShown = false }
            }
            .accessibilityElement(children: .contain)
            .accessibilityLabel("Page \(session.currentPage) of \(session.pageCount)")
    }
}

private struct OptionalTitle: ViewModifier {
    let title: String?

    func body(content: Content) -> some View {
        if let title { content.navigationTitle(title) } else { content }
    }
}
