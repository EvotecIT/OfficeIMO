import SwiftUI

/// The window toolbar on the Mac, the navigation and bottom bars on iPad and iPhone. The system draws them as glass.
struct DocumentToolbar: ToolbarContent {
    @Bindable var session: DocumentSession
    let fileURL: URL?
    let isCompact: Bool
    @Binding var isAddingNote: Bool

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
            if session.task == .markup {
                // Each segment shows its symbol and names its tool, rather than the symbol, to VoiceOver and in its tooltip.
                Picker("Tool", selection: markupChoice) {
                    Image(systemName: ShapeTool.select.systemImage).accessibilityLabel(Text(ShapeTool.select.title))
                        .help(Text(ShapeTool.select.title))
                        .tag(MarkupChoice.shape(.select))
                    ForEach(MarkupStyle.allCases) { style in
                        Image(systemName: style.systemImage).accessibilityLabel(Text(style.title)).help(Text(style.title))
                            .tag(MarkupChoice.style(style))
                    }
                }
                .pickerStyle(.segmented)
                // The shapes share one pull-down, as in Preview, which shows the shape a drag draws.
                Menu {
                    Picker("Shapes", selection: shapeChoice) {
                        ForEach(ShapeTool.shapes) { Label($0.title, systemImage: $0.systemImage).tag(Optional($0)) }
                    }
                    .pickerStyle(.inline)
                } label: {
                    Label("Shapes", systemImage: shapeChoice.wrappedValue?.systemImage ?? "square.on.circle")
                }
                .help("Shapes")
                ColorPicker("Colour", selection: markupColor, supportsOpacity: false)
                Button("Add Note", systemImage: "note.text.badge.plus") { isAddingNote = true }
            }
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
        if session.task == .markup {
            // The system tool picker takes the bottom of the screen, so Markup's commands sit at the top.
            ToolbarItemGroup(placement: .topBarTrailing) {
                undoRedo
                shapesMenu
                Button("Add Note", systemImage: "note.text.badge.plus") { isAddingNote = true }
            }
            ToolbarItem(placement: .confirmationAction) {
                Button("Done", systemImage: "checkmark") { session.task = .read }
                    .buttonStyle(.glassProminent)
            }
        } else if session.task == .pages {
            ToolbarItemGroup(placement: .topBarTrailing) { undoRedo }
        }
        // In Markup the bottom bar stays empty: the tool picker is there.
        if session.task != .markup && isCompact {
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
        } else if session.task != .markup {
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
        Button("Undo", systemImage: "arrow.uturn.backward") { session.undoEdit() }
            .disabled(!session.canUndoEdit || session.isWorking)
        Button("Redo", systemImage: "arrow.uturn.forward") { session.redoEdit() }
            .disabled(!session.canRedoEdit || session.isWorking)
    }
    #endif

    /// Actions on the selected pages, shown while the Pages task is open.
    @ViewBuilder private var pageActions: some View {
        if session.task == .pages {
            PageActionButtons(session: session, pages: session.selection.sorted())
        }
    }

    #if os(iOS)
    /// Selecting markup to move or resize it, and shapes and text boxes, which a drag on the page draws in place of the
    /// drawing tools until Done.
    private var shapesMenu: some View {
        Menu("Shapes", systemImage: "square.on.circle") {
            Picker("Shapes", selection: $session.shapeTool) {
                ForEach(ShapeTool.allCases) { Label($0.title, systemImage: $0.systemImage).tag(Optional($0)) }
            }
            .pickerStyle(.inline)
        }
        .disabled(!session.canEditAnnotations)
    }
    #endif

    #if os(macOS)
    /// The pen, the highlighter or a shape, as one choice.
    private var markupChoice: Binding<MarkupChoice> {
        Binding {
            session.shapeTool.map(MarkupChoice.shape) ?? .style(session.markupStyle)
        } set: { choice in
            switch choice {
            case .style(let style):
                session.markupStyle = style
                session.shapeTool = nil
            case .shape(let shape):
                session.shapeTool = shape
            }
        }
    }

    /// The shape a drag draws, if one is chosen.
    private var shapeChoice: Binding<ShapeTool?> {
        Binding {
            session.shapeTool == .select ? nil : session.shapeTool
        } set: { shape in
            session.shapeTool = shape
        }
    }

    private var markupColor: Binding<Color> {
        Binding {
            Color(red: session.markupColor.red, green: session.markupColor.green, blue: session.markupColor.blue)
        } set: { color in
            let resolved = NSColor(color).usingColorSpace(.sRGB) ?? .systemRed
            session.markupColor = EngineColor(red: Double(resolved.redComponent), green: Double(resolved.greenComponent),
                                              blue: Double(resolved.blueComponent))
        }
    }
    #endif

    /// The task menu: the current task's name, opening the list of tasks.
    private var taskMenu: some View {
        Menu {
            Picker("Task", selection: $session.task) {
                ForEach(StudioTask.allCases) { Label($0.title, systemImage: $0.systemImage).tag($0) }
            }
            .pickerStyle(.inline)
            Section {
                Button("Make Searchable", systemImage: "text.viewfinder") { session.makeSearchable() }
                    .disabled(!session.canMakeSearchable)
            }
            Section {
                StructureCommands(session: session)
            }
            Section {
                SignatureMenu(session: session)
                SecurityCommands(session: session)
            }
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
            #if os(iOS)
            DocumentShareButton(session: session, fileURL: fileURL)
            #else
            ShareLink(item: SharedDocument(session: session, fileURL: fileURL),
                      preview: SharePreview(fileURL.deletingPathExtension().lastPathComponent))
            #endif
        }
    }

    private var inspectorToggle: some View {
        Button("Inspector", systemImage: isCompact ? "info.circle" : "sidebar.trailing") {
            session.isInspectorShown.toggle()
        }
        .help(session.isInspectorShown ? "Hide Inspector" : "Show Inspector")
    }
}
