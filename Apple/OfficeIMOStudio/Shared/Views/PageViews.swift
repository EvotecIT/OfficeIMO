import SwiftUI

/// A page thumbnail from the engine, white until its image arrives.
struct PageThumbnail: View {
    let session: DocumentSession
    let page: DocumentFacts.Page
    let width: CGFloat
    @Environment(\.displayScale) private var displayScale
    @State private var image: (key: String, image: CGImage)?

    private var key: String { "\(page.number)-\(session.revision)-\(Int(width * displayScale))" }

    var body: some View {
        let size = page.displaySize
        ZStack {
            Color.white
            if let image, image.key == key {
                Image(decorative: image.image, scale: displayScale)
                    .resizable()
                    .interpolation(.high)
            }
        }
        .frame(width: width, height: (width * size.height / size.width).rounded())
        .shadow(color: .black.opacity(0.22), radius: 1.5, y: 1)
        .task(id: key) {
            let key = key
            if let rendered = await session.thumbnails.image(page: page.number, pixelWidth: Int(width * displayScale), session: session) {
                image = (key, rendered)
            }
        }
        .accessibilityElement()
        .accessibilityLabel("Page \(page.number)")
    }
}

/// The sidebar: page thumbnails or the table of contents, switched from the View menu or the picker above them.
struct PageSidebar: View {
    let session: DocumentSession

    var body: some View {
        Group {
            switch session.sidebarMode {
            case .thumbnails: PageThumbnailList(session: session)
            case .outline: OutlineList(session: session)
            }
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity)
        .safeAreaInset(edge: .top) {
            SidebarModePicker(session: session)
                .padding(.horizontal, 12)
                .padding(.vertical, 6)
        }
    }
}

/// The page list in the sidebar. Selecting a page shows it; the selection follows the page in view.
struct PageThumbnailList: View {
    let session: DocumentSession

    var body: some View {
        ScrollViewReader { proxy in
            List(selection: selection) {
                ForEach(session.facts?.pages ?? [], id: \.number) { page in
                    VStack(spacing: 6) {
                        PageThumbnail(session: session, page: page, width: 96)
                        Text(page.number, format: .number)
                            .font(.caption)
                            .foregroundStyle(.secondary)
                    }
                    .frame(maxWidth: .infinity)
                    .padding(.vertical, 6)
                    .tag(page.number)
                    .id(page.number)
                }
            }
            .onChange(of: session.currentPage) { _, page in
                withAnimation { proxy.scrollTo(page) }
            }
        }
    }

    private var selection: Binding<Int?> {
        Binding { session.currentPage } set: { if let page = $0, page != session.currentPage { session.show(page: page) } }
    }
}

/// The page list as a sheet in compact width, over the page it controls.
struct PageListSheet: View {
    let session: DocumentSession
    @Environment(\.dismiss) private var dismiss

    var body: some View {
        NavigationStack {
            Group {
                if session.sidebarMode == .outline {
                    OutlineList(session: session)
                } else {
                    grid
                }
            }
            .navigationTitle(session.sidebarMode == .outline ? "Table of Contents" : "Pages")
            #if os(iOS)
            .navigationBarTitleDisplayMode(.inline)
            #endif
            .toolbar {
                ToolbarItem(placement: .principal) {
                    SidebarModePicker(session: session).fixedSize()
                }
                ToolbarItem(placement: .confirmationAction) {
                    Button("Done", systemImage: "checkmark") { dismiss() }
                }
            }
        }
        .presentationDetents([.medium, .large])
        .presentationBackgroundInteraction(.enabled(upThrough: .medium))
    }

    private var grid: some View {
        ScrollViewReader { proxy in
            ScrollView {
                LazyVGrid(columns: [GridItem(.adaptive(minimum: 72), spacing: 16)], spacing: 16) {
                    ForEach(session.facts?.pages ?? [], id: \.number) { page in
                        Button {
                            session.show(page: page.number)
                        } label: {
                            PageCell(session: session, page: page, width: 72, isSelected: page.number == session.currentPage)
                        }
                        .buttonStyle(.plain)
                        .id(page.number)
                    }
                }
                .padding()
            }
            .onAppear { proxy.scrollTo(session.currentPage, anchor: .center) }
        }
    }
}

/// A thumbnail with its number, marked when selected.
struct PageCell: View {
    let session: DocumentSession
    let page: DocumentFacts.Page
    let width: CGFloat
    let isSelected: Bool

    var body: some View {
        VStack(spacing: 6) {
            PageThumbnail(session: session, page: page, width: width)
                .overlay {
                    if isSelected {
                        Rectangle().strokeBorder(.tint, lineWidth: 3).padding(-4)
                    }
                }
            Text(page.number, format: .number)
                .font(.caption.weight(isSelected ? .semibold : .regular))
                .foregroundStyle(isSelected ? AnyShapeStyle(.white) : AnyShapeStyle(.secondary))
                .padding(.horizontal, 7)
                .padding(.vertical, 1)
                .background {
                    if isSelected { Capsule().fill(.tint) }
                }
        }
        .accessibilityElement(children: .combine)
        .accessibilityAddTraits(isSelected ? .isSelected : [])
    }
}

/// The Pages task: every page in a grid to select, rotate, duplicate, delete, extract and drag into a new order.
struct PagesTaskView: View {
    @Bindable var session: DocumentSession
    let isCompact: Bool
    @State private var dropTarget: Int?

    var body: some View {
        let width: CGFloat = isCompact ? 96 : 140
        ScrollView {
            LazyVGrid(columns: [GridItem(.adaptive(minimum: width + 8), spacing: 24)], spacing: 24) {
                ForEach(session.facts?.pages ?? [], id: \.number) { page in
                    cell(page, width: width)
                }
            }
            .padding(24)
        }
        .background(.background.secondary)
        .safeAreaInset(edge: .top) {
            if session.isReady && !session.canEditPages {
                Label(session.canExtractPages
                      ? "This document’s pages can’t be changed yet: it has content, such as form fields or accessibility tags, that OfficeIMO can’t rewrite safely. You can still extract pages."
                      : "This document’s pages can’t be changed or extracted yet: it has content, such as form fields or accessibility tags, that OfficeIMO can’t rewrite safely.",
                      systemImage: "info.circle")
                    .font(.callout)
                    .foregroundStyle(.secondary)
                    .frame(maxWidth: .infinity, alignment: .leading)
                    .padding(.horizontal, 24)
                    .padding(.vertical, 10)
                    .background(.bar)
            }
        }
        .navigationTitle(session.selection.isEmpty ? "Select Pages" : "\(session.selection.count) Selected")
        #if os(macOS)
        .onDeleteCommand {
            if !session.selection.isEmpty { session.delete(session.selection.sorted()) }
        }
        #endif
        .toolbar {
            ToolbarItem(placement: .cancellationAction) {
                let all = session.selection.count == session.pageCount
                Button(all ? "Deselect All" : "Select All") {
                    session.selection = all ? [] : Set(1...max(session.pageCount, 1))
                }
            }
        }
    }

    private func cell(_ page: DocumentFacts.Page, width: CGFloat) -> some View {
        let number = page.number
        let isSelected = session.selection.contains(number)
        return Button {
            if isSelected { session.selection.remove(number) } else { session.selection.insert(number) }
        } label: {
            PageCell(session: session, page: page, width: width, isSelected: isSelected)
                .overlay(alignment: .leading) {
                    // Where dropped pages will go: before this page.
                    if dropTarget == number {
                        Capsule().fill(.tint).frame(width: 4).padding(.vertical, -6).offset(x: -14)
                    }
                }
        }
        .buttonStyle(.plain)
        .contextMenu { PageContextMenu(session: session, page: number) }
        .draggable(PageDrag(window: session.id, pages: isSelected ? session.selection.sorted() : [number])) {
            PageThumbnail(session: session, page: page, width: 72)
        }
        .dropDestination(for: PageDrag.self) { drops, _ in
            dropTarget = nil
            guard let drag = drops.first, drag.window == session.id, session.canEditPages else { return false }
            session.move(drag.pages, before: number)
            return true
        } isTargeted: { targeted in
            if targeted { dropTarget = number } else if dropTarget == number { dropTarget = nil }
        }
        .accessibilityAction(named: "Move Earlier") { session.move([number], before: number - 1) }
        .accessibilityAction(named: "Move Later") { session.move([number], before: number + 2) }
    }
}
