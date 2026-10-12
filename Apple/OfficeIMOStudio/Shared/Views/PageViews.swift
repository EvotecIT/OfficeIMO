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

/// The page list in the sidebar. Selecting a page shows it; the selection follows the page in view.
struct PageSidebar: View {
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
            .navigationTitle("Pages")
            #if os(iOS)
            .navigationBarTitleDisplayMode(.inline)
            #endif
            .toolbar {
                ToolbarItem(placement: .confirmationAction) {
                    Button("Done", systemImage: "checkmark") { dismiss() }
                }
            }
        }
        .presentationDetents([.medium, .large])
        .presentationBackgroundInteraction(.enabled(upThrough: .medium))
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

/// The Pages task: every page in a grid to select and rotate. The selection's actions are in the toolbar.
struct PagesTaskView: View {
    @Bindable var session: DocumentSession
    let isCompact: Bool

    var body: some View {
        let width: CGFloat = isCompact ? 96 : 140
        ScrollView {
            LazyVGrid(columns: [GridItem(.adaptive(minimum: width + 8), spacing: 24)], spacing: 24) {
                ForEach(session.facts?.pages ?? [], id: \.number) { page in
                    Button {
                        if session.selection.contains(page.number) {
                            session.selection.remove(page.number)
                        } else {
                            session.selection.insert(page.number)
                        }
                    } label: {
                        PageCell(session: session, page: page, width: width, isSelected: session.selection.contains(page.number))
                    }
                    .buttonStyle(.plain)
                }
            }
            .padding(24)
        }
        .background(.background.secondary)
        .navigationTitle(session.selection.isEmpty ? "Select Pages" : "\(session.selection.count) Selected")
        .toolbar {
            ToolbarItem(placement: .cancellationAction) {
                let all = session.selection.count == session.pageCount
                Button(all ? "Deselect All" : "Select All") {
                    session.selection = all ? [] : Set(1...max(session.pageCount, 1))
                }
            }
        }
    }
}
