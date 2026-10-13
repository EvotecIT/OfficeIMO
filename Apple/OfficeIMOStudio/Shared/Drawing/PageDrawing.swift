import CoreGraphics
import SwiftUI

/// Shapes a drag on the page draws in Markup, as Preview's Shapes and Text tools add them.
enum ShapeTool: String, CaseIterable, Identifiable {
    case rectangle
    case oval
    case line
    case arrow
    case textBox

    var id: Self { self }

    var title: LocalizedStringKey {
        switch self {
        case .rectangle: "Rectangle"
        case .oval: "Oval"
        case .line: "Line"
        case .arrow: "Arrow"
        case .textBox: "Text Box"
        }
    }

    var systemImage: String {
        switch self {
        case .rectangle: "rectangle"
        case .oval: "oval"
        case .line: "line.diagonal"
        case .arrow: "arrow.up.right"
        case .textBox: "character.textbox"
        }
    }

    /// What to do on the page, shown while the shape is chosen on iPad and iPhone.
    var hint: LocalizedStringKey {
        switch self {
        case .rectangle: "Drag on the page to draw a rectangle."
        case .oval: "Drag on the page to draw an oval."
        case .line: "Drag on the page to draw a line."
        case .arrow: "Drag on the page to draw an arrow."
        case .textBox: "Drag on the page where the text goes."
        }
    }

    /// A shape draws along the drag; a line from where it starts to where it ends.
    var isLine: Bool { self == .line || self == .arrow }
}

/// The Mac's Markup tools: the pen, the highlighter or a shape.
enum MarkupChoice: Hashable {
    case style(MarkupStyle)
    case shape(ShapeTool)
}

/// What a drag on the page does, when it doesn't scroll, select text or draw ink.
enum PageDragTool: Equatable {
    case shape(ShapeTool)
    case redaction
}

/// A text box drawn on a page, in the page's displayed points, waiting for its text.
struct PendingTextBox: Equatable {
    let page: Int
    let rect: CGRect
}

extension DocumentSession {
    /// The width of a drawn shape's outline, in points.
    static let shapeStrokeWidth = 2.0

    var pageDragTool: PageDragTool? {
        guard isReady, task != .pages else { return nil }
        if redactionAreas != nil, canRedact { return .redaction }
        if task == .markup, canEditAnnotations, let shapeTool { return .shape(shapeTool) }
        return nil
    }

    /// Adds a shape dragged on a page, from one point to another in the page's displayed points. A click without a
    /// drag adds the shape at a standard size; a text box first asks for its text.
    func addShape(_ tool: ShapeTool, page: Int, from start: CGPoint, to end: CGPoint, failed: (() -> Void)? = nil) {
        guard canEditAnnotations, let facts, facts.pages.indices.contains(page - 1) else {
            failed?()
            return
        }
        let size = facts.pages[page - 1].displaySize
        let request: AnnotationRequest
        if tool.isLine {
            var end = end
            if hypot(end.x - start.x, end.y - start.y) < 4 { end = CGPoint(x: min(start.x + 120, size.width), y: start.y) }
            guard hypot(end.x - start.x, end.y - start.y) >= 4 else {
                failed?()
                return
            }
            request = AnnotationRequest(tool: .line, page: page,
                                        strokes: [[EnginePoint(x: start.x, y: start.y), EnginePoint(x: end.x, y: end.y)]],
                                        author: Self.author, color: markupColor, strokeWidth: Self.shapeStrokeWidth,
                                        lineEnd: tool == .arrow ? "OpenArrow" : nil)
        } else {
            let standard = tool == .textBox ? CGSize(width: 180, height: 36) : CGSize(width: 120, height: 80)
            let rect = Self.dragRect(from: start, to: end, standard: standard, page: size)
            if tool == .textBox {
                failed?()
                pendingTextBox = PendingTextBox(page: page, rect: rect)
                return
            }
            request = AnnotationRequest(tool: tool == .oval ? .ellipse : .rectangle, page: page,
                                        bounds: .init(x: rect.minX, y: rect.minY, width: rect.width, height: rect.height),
                                        author: Self.author, color: markupColor, strokeWidth: Self.shapeStrokeWidth)
        }
        addMarkup([request], name: String(localized: "Add Shape"), failed: failed)
    }

    /// Adds the text box waiting for its text.
    func addTextBox(_ text: String) {
        guard let box = pendingTextBox else { return }
        pendingTextBox = nil
        let text = text.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !text.isEmpty, canEditAnnotations else { return }
        let request = AnnotationRequest(tool: .freeText, page: box.page,
                                        bounds: .init(x: box.rect.minX, y: box.rect.minY, width: box.rect.width, height: box.rect.height),
                                        text: text, author: Self.author, color: markupColor)
        addMarkup([request], name: String(localized: "Add Text Box"))
    }

    // MARK: Redacting areas

    /// Starts marking areas to redact: each drag across a page marks one, and Redact removes them together.
    func beginMarkingRedactions() {
        guard canRedact else { return }
        if task != .read { task = .read }
        textSelection = nil
        redactionAreas = []
    }

    /// Marks an area dragged across a page, in its displayed points; a click on a marked area unmarks it.
    func markRedaction(page: Int, from start: CGPoint, to end: CGPoint) {
        guard var areas = redactionAreas else { return }
        if hypot(end.x - start.x, end.y - start.y) < 4 {
            areas.removeAll { $0.page == page && $0.bounds.cgRect.contains(start) }
        } else {
            let rect = CGRect(x: min(start.x, end.x), y: min(start.y, end.y), width: abs(end.x - start.x), height: abs(end.y - start.y))
            guard rect.width >= 2, rect.height >= 2 else { return }
            areas.append(RedactionMark(page: page, bounds: .init(x: rect.minX, y: rect.minY, width: rect.width, height: rect.height)))
        }
        redactionAreas = areas
    }

    /// Removes everything under the marked areas as one change.
    func redactMarkedAreas() {
        guard let areas = redactionAreas, !areas.isEmpty else { return }
        redactionAreas = nil
        redact(.init(areas: areas))
    }

    /// A rectangle from a drag, or a standard size around the point of a click, kept on the page.
    private static func dragRect(from start: CGPoint, to end: CGPoint, standard: CGSize, page: CGSize) -> CGRect {
        var rect = CGRect(x: min(start.x, end.x), y: min(start.y, end.y), width: abs(end.x - start.x), height: abs(end.y - start.y))
        if rect.width < 4 && rect.height < 4 {
            rect = CGRect(x: start.x - standard.width / 2, y: start.y - standard.height / 2, width: standard.width, height: standard.height)
        }
        rect.size.width = max(min(rect.width, page.width), 4)
        rect.size.height = max(min(rect.height, page.height), 4)
        rect.origin.x = min(max(rect.minX, 0), page.width - rect.width)
        rect.origin.y = min(max(rect.minY, 0), page.height - rect.height)
        return rect
    }
}

/// The outline of a shape or an area while it is dragged, in the page view's coordinates, which grow downward.
enum DragOutline {
    static func path(for tool: PageDragTool, from start: CGPoint, to end: CGPoint) -> CGPath {
        let path = CGMutablePath()
        let rect = CGRect(x: min(start.x, end.x), y: min(start.y, end.y), width: abs(end.x - start.x), height: abs(end.y - start.y))
        switch tool {
        case .shape(.oval):
            path.addEllipse(in: rect)
        case .shape(.line), .shape(.arrow):
            path.move(to: start)
            path.addLine(to: end)
            if tool == .shape(.arrow), hypot(end.x - start.x, end.y - start.y) > 1 {
                let angle = atan2(end.y - start.y, end.x - start.x)
                for side in [-1.0, 1.0] {
                    let wing = angle + .pi + side * .pi / 6
                    path.move(to: end)
                    path.addLine(to: CGPoint(x: end.x + 10 * cos(wing), y: end.y + 10 * sin(wing)))
                }
            }
        default:
            path.addRect(rect)
        }
        return path
    }

    /// The marked areas, offset to their pages in the page layout.
    static func path(for areas: [RedactionMark], layout: PageLayout) -> CGPath {
        let path = CGMutablePath()
        for area in areas where layout.frames.indices.contains(area.page - 1) {
            let origin = layout.frames[area.page - 1].origin
            path.addRect(area.bounds.cgRect.offsetBy(dx: origin.x, dy: origin.y))
        }
        return path
    }
}

/// Over the pages while marking areas to redact: how many are marked, Cancel and Redact.
struct RedactionBanner: View {
    let session: DocumentSession

    var body: some View {
        let count = session.redactionAreas?.count ?? 0
        HStack(spacing: 10) {
            Image(systemName: "rectangle.dashed")
                .accessibilityHidden(true)
            if count == 0 {
                Text("Drag across what to remove.")
            } else {
                Text("\(count) area(s) marked")
            }
            Button("Cancel") { session.redactionAreas = nil }
                .buttonStyle(.borderless)
            Button("Redact") { session.redactMarkedAreas() }
                .buttonStyle(.glassProminent)
                .tint(.red)
                .disabled(count == 0 || session.isWorking)
        }
        .font(.callout.weight(.medium))
        .padding(.horizontal, 14)
        .padding(.vertical, 8)
        .glassEffect(.regular, in: .capsule)
    }
}

/// Over the pages while a shape is chosen on iPad and iPhone, where the drawing tools take the bottom of the screen
/// otherwise: what to do, the colour and Done, which goes back to drawing.
struct ShapeBanner: View {
    @Bindable var session: DocumentSession
    let tool: ShapeTool

    var body: some View {
        HStack(spacing: 10) {
            Image(systemName: tool.systemImage)
                .accessibilityHidden(true)
            Text(tool.hint)
            ColorPicker("Colour", selection: markupColor, supportsOpacity: false)
                .labelsHidden()
            Button("Done") { session.shapeTool = nil }
                .buttonStyle(.borderless)
        }
        .font(.callout.weight(.medium))
        .padding(.horizontal, 14)
        .padding(.vertical, 8)
        .glassEffect(.regular, in: .capsule)
    }

    private var markupColor: Binding<Color> {
        Binding {
            Color(red: session.markupColor.red, green: session.markupColor.green, blue: session.markupColor.blue)
        } set: { color in
            let resolved = color.resolve(in: EnvironmentValues())
            session.markupColor = EngineColor(red: Double(resolved.red), green: Double(resolved.green), blue: Double(resolved.blue))
        }
    }
}

/// Asks for a drawn text box's text.
struct TextBoxPrompt: ViewModifier {
    @Bindable var session: DocumentSession
    @State private var text = ""

    func body(content: Content) -> some View {
        content.alert("Add Text Box", isPresented: Binding { session.pendingTextBox != nil } set: { if !$0 { session.pendingTextBox = nil } }) {
            TextField("Text", text: $text)
            Button("Cancel", role: .cancel) {
                session.pendingTextBox = nil
                text = ""
            }
            Button("Add") {
                session.addTextBox(text)
                text = ""
            }
        } message: {
            Text("The text shows in a box on the page, in the Markup colour.")
        }
    }
}
