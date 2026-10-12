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
        if isCompact {
            ToolbarItem(placement: .topBarTrailing) { inspectorToggle }
            if session.task == .pages {
                ToolbarItemGroup(placement: .bottomBar) {
                    taskMenu
                    Spacer()
                    pageActions
                }
            } else {
                ToolbarItemGroup(placement: .bottomBar) {
                    Button("Pages", systemImage: "rectangle.grid.1x2") { session.isPageListShown = true }
                        .disabled(!session.isReady)
                    Spacer()
                    taskMenu
                    Spacer()
                    share
                }
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

    /// Rotation for the selected pages, shown while the Pages task is open.
    @ViewBuilder private var pageActions: some View {
        if session.task == .pages {
            Button("Rotate Left", systemImage: "rotate.left") { session.rotate(session.selection.sorted(), degrees: -90) }
                .disabled(session.selection.isEmpty)
            Button("Rotate Right", systemImage: "rotate.right") { session.rotate(session.selection.sorted(), degrees: 90) }
                .disabled(session.selection.isEmpty)
        }
    }

    private var taskMenu: some View {
        Picker(selection: $session.task) {
            ForEach(StudioTask.allCases) { Label($0.title, systemImage: $0.systemImage).tag($0) }
        } label: {
            Label(session.task.title, systemImage: session.task.systemImage)
        }
        .pickerStyle(.menu)
        .labelStyle(.titleOnly)
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
