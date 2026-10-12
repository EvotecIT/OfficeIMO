import SwiftUI

/// The window toolbar on the Mac, the navigation and bottom bars on iPad and iPhone. The system draws them as glass.
struct DocumentToolbar: ToolbarContent {
    @Bindable var session: DocumentSession
    let fileURL: URL?
    let isCompact: Bool

    var body: some ToolbarContent {
        #if os(macOS)
        ToolbarItem(placement: .principal) {
            Picker("Task", selection: $session.task) {
                ForEach(StudioTask.allCases) { Text($0.title).tag($0) }
            }
            .pickerStyle(.segmented)
            .disabled(!session.isReady)
        }
        ToolbarItemGroup(placement: .primaryAction) {
            pageActions
            if session.task == .read {
                ControlGroup {
                    Button("Zoom Out", systemImage: "minus.magnifyingglass") { session.zoom(.zoomOut) }
                    Button("Zoom In", systemImage: "plus.magnifyingglass") { session.zoom(.zoomIn) }
                } label: {
                    Label("Zoom", systemImage: "magnifyingglass")
                }
                .disabled(!session.isReady)
            }
            share
        }
        ToolbarItem(placement: .primaryAction) {
            inspectorToggle
        }
        #else
        if session.task == .pages {
            ToolbarItemGroup(placement: .topBarTrailing) { undoRedo }
        }
        if isCompact {
            ToolbarItem(placement: .topBarTrailing) { inspectorToggle }
            // Separate items with flexible spacers, so each is its own glass control in the bottom bar.
            if session.task == .pages {
                ToolbarItem(placement: .bottomBar) { taskMenu }
                ToolbarSpacer(.flexible, placement: .bottomBar)
                ToolbarItemGroup(placement: .bottomBar) { pageActions }
            } else {
                ToolbarItem(placement: .bottomBar) {
                    Button("Pages", systemImage: "rectangle.grid.1x2") { session.isPageListShown = true }
                        .disabled(!session.isReady)
                }
                ToolbarSpacer(.flexible, placement: .bottomBar)
                ToolbarItem(placement: .bottomBar) { taskMenu }
                ToolbarSpacer(.flexible, placement: .bottomBar)
                ToolbarItem(placement: .bottomBar) { share }
            }
        } else {
            ToolbarItemGroup(placement: .topBarTrailing) {
                taskMenu
                share
                inspectorToggle
            }
            if session.task == .pages {
                ToolbarItemGroup(placement: .bottomBar) {
                    Spacer()
                    pageActions
                }
            }
        }
        #endif
    }

    #if os(iOS)
    /// Undo and Redo while editing, as Apple's editing apps show them; they act through the window's undo manager,
    /// which also serves the keyboard shortcuts and the three-finger gestures.
    @ViewBuilder private var undoRedo: some View {
        Button("Undo", systemImage: "arrow.uturn.backward") { session.undoManager?.undo() }
            .disabled(session.facts?.canUndo != true || session.isWorking)
        Button("Redo", systemImage: "arrow.uturn.forward") { session.undoManager?.redo() }
            .disabled(session.facts?.canRedo != true || session.isWorking)
    }
    #endif

    /// Rotation for the selected pages, shown while the Pages task is open.
    @ViewBuilder private var pageActions: some View {
        if session.task == .pages {
            Button("Rotate Left", systemImage: "rotate.left") { session.rotate(session.selection.sorted(), degrees: -90) }
                .disabled(session.selection.isEmpty)
            Button("Rotate Right", systemImage: "rotate.right") { session.rotate(session.selection.sorted(), degrees: 90) }
                .disabled(session.selection.isEmpty)
        }
    }

    /// The task menu: the current task's name, opening the list of tasks.
    private var taskMenu: some View {
        Menu {
            Picker("Task", selection: $session.task) {
                ForEach(StudioTask.allCases) { Label($0.title, systemImage: $0.systemImage).tag($0) }
            }
            .pickerStyle(.inline)
        } label: {
            HStack(spacing: 4) {
                Text(session.task.title)
                Image(systemName: "chevron.down")
                    .font(.caption.weight(.semibold))
                    .accessibilityHidden(true)
            }
            .accessibilityLabel(Text("Task: \(Text(session.task.title))"))
        }
        .menuIndicator(.hidden)
        .fixedSize()
        .disabled(!session.isReady)
    }

    @ViewBuilder private var share: some View {
        if let fileURL {
            ShareLink(item: fileURL)
        }
    }

    private var inspectorToggle: some View {
        Button("Inspector", systemImage: isCompact ? "info.circle" : "sidebar.trailing") {
            session.isInspectorShown.toggle()
        }
        .help(session.isInspectorShown ? "Hide Inspector" : "Show Inspector")
    }
}
