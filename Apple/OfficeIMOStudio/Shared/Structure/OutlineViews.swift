import SwiftUI

/// The document's table of contents: choosing an entry shows its page. Bookmarks the document lets the app change
/// can be added for the page in view, renamed, pointed at the page in view and deleted, each as one undoable change.
struct OutlineList: View {
    let session: DocumentSession
    @State private var renaming: OutlineNode?
    @State private var title = ""

    var body: some View {
        let nodes = session.currentStructure?.outlineNodes ?? []
        Group {
            if session.currentStructure == nil {
                ProgressView().frame(maxWidth: .infinity, maxHeight: .infinity)
            } else if nodes.isEmpty {
                // The sidebar is narrow, so the empty state stacks its text and buttons rather than using the wide system layout.
                VStack(spacing: 10) {
                    Image(systemName: "bookmark")
                        .font(.title2)
                        .foregroundStyle(.secondary)
                    Text("No Bookmarks")
                        .font(.headline)
                    Text("Bookmarks you add, or ones made from the document’s headings, appear here.")
                        .font(.callout)
                        .foregroundStyle(.secondary)
                        .multilineTextAlignment(.center)
                        .fixedSize(horizontal: false, vertical: true)
                    if session.canEditOutline {
                        Button("Add Bookmark") { session.addBookmark() }
                        Button("Make from Headings") { session.rebuildOutline() }
                    }
                }
                .controlSize(.small)
                .padding(.horizontal, 14)
                .frame(maxWidth: .infinity, maxHeight: .infinity)
            } else {
                List(selection: selection) {
                    OutlineGroup(nodes, children: \.children) { node in
                        Label(node.title, systemImage: node.page == nil ? "bookmark.slash" : "bookmark")
                            .labelStyle(.titleOnly)
                            .foregroundStyle(node.page == nil ? .secondary : .primary)
                            .badge(node.page.map { Text($0, format: .number) })
                            .tag(node.id)
                            .contextMenu { menu(node) }
                    }
                }
            }
        }
        .task(id: session.revision) { await session.loadStructure() }
        .alert("Rename Bookmark", isPresented: Binding { renaming != nil } set: { if !$0 { renaming = nil } }) {
            TextField("Title", text: $title)
            Button("Rename") {
                if let renaming { session.renameBookmark(renaming, to: title) }
                renaming = nil
            }
            Button("Cancel", role: .cancel) { renaming = nil }
        }
    }

    @ViewBuilder private func menu(_ node: OutlineNode) -> some View {
        if node.bookmark != nil && session.canEditOutline {
            Button("Rename…", systemImage: "pencil") {
                title = node.title
                renaming = node
            }
            Button("Use Page \(session.currentPage)", systemImage: "arrow.uturn.right") { session.retargetBookmark(node) }
                .disabled(node.page == session.currentPage)
            Divider()
            Button("Delete", systemImage: "trash", role: .destructive) { session.removeBookmark(node) }
        }
    }

    private var selection: Binding<String?> {
        Binding {
            nil
        } set: { key in
            guard let key, let node = Self.find(key, in: session.currentStructure?.outlineNodes ?? []), let page = node.page else { return }
            session.show(page: page)
        }
    }

    private static func find(_ key: String, in nodes: [OutlineNode]) -> OutlineNode? {
        for node in nodes {
            if node.id == key { return node }
            if let found = find(key, in: node.children ?? []) { return found }
        }
        return nil
    }
}

/// The sidebar's switch between thumbnails and the table of contents, for touch, where there is no View menu.
struct SidebarModePicker: View {
    @Bindable var session: DocumentSession

    var body: some View {
        Picker("Sidebar", selection: $session.sidebarMode) {
            ForEach(SidebarMode.allCases) { mode in
                Label(mode.title, systemImage: mode.systemImage).tag(mode)
            }
        }
        .pickerStyle(.segmented)
        .labelStyle(.iconOnly)
        .labelsHidden()
    }
}
