import CoreGraphics
import SwiftUI

/// What a drag on the page does in Markup, as Preview's Select, Shapes and Text tools do: select, move and resize
/// annotations, or draw a shape.
enum ShapeTool: String, CaseIterable, Identifiable {
    case select
    case rectangle
    case oval
    case line
    case arrow
    case polygon
    case textBox

    /// The tools that draw.
    static var shapes: [ShapeTool] { allCases.filter { $0 != .select } }

    var id: Self { self }

    var title: LocalizedStringKey {
        switch self {
        case .select: "Select"
        case .rectangle: "Rectangle"
        case .oval: "Oval"
        case .line: "Line"
        case .arrow: "Arrow"
        case .polygon: "Polygon"
        case .textBox: "Text Box"
        }
    }

    var systemImage: String {
        switch self {
        case .select: "cursorarrow"
        case .rectangle: "rectangle"
        case .oval: "oval"
        case .line: "line.diagonal"
        case .arrow: "arrow.up.right"
        case .polygon: "pentagon"
        case .textBox: "character.textbox"
        }
    }

    /// What to do on the page, shown while the shape is chosen on iPad and iPhone.
    var hint: LocalizedStringKey {
        switch self {
        case .select: "Tap markup to select it, then drag it or a corner."
        case .rectangle: "Drag on the page to draw a rectangle."
        case .oval: "Drag on the page to draw an oval."
        case .line: "Drag on the page to draw a line."
        case .arrow: "Drag on the page to draw an arrow."
        case .polygon: "Drag around an area to outline it."
        case .textBox: "Drag on the page where the text goes."
        }
    }

    /// A shape draws along the drag; a line from where it starts to where it ends.
    var isLine: Bool { self == .line || self == .arrow }
}

/// The Mac's Markup tools: selecting, the pen, the highlighter or a shape.
enum MarkupChoice: Hashable {
    case style(MarkupStyle)
    case shape(ShapeTool)
}

/// What a drag on the page does, when it doesn't scroll, select text or draw ink.
enum PageDragTool: Equatable {
    case shape(ShapeTool)
    case redaction
    /// Selects markup, then moves or resizes it.
    case select
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
        if task == .markup, canEditAnnotations, let shapeTool { return shapeTool == .select ? .select : .shape(shapeTool) }
        return nil
    }

    // MARK: Selecting markup

    /// The selected annotation, while the current page's annotations are read for the revision shown.
    var selectedPageAnnotation: AnnotationList.Annotation? {
        guard let selectedAnnotation, let list = pageAnnotations, list.page == currentPage, list.revision == revision else { return nil }
        return list.annotations.first { $0.objectNumber == selectedAnnotation }
    }

    /// The topmost annotation under a point of a page, in its displayed points; thin lines are easier to hit.
    func annotation(at point: CGPoint, page: Int) -> AnnotationList.Annotation? {
        guard let list = pageAnnotations, list.page == page, list.revision == revision else { return nil }
        return list.annotations.last { $0.bounds.cgRect.insetBy(dx: -4, dy: -4).contains(point) }
    }

    /// Moves an annotation to a rectangle of its page, in displayed points, or resizes it to that rectangle.
    func setAnnotationBounds(_ objectNumber: Int, _ rect: CGRect, resizing: Bool, failed: (() -> Void)? = nil) {
        guard canEditAnnotations else {
            failed?()
            return
        }
        let subtype = pageAnnotations?.annotations.first { $0.objectNumber == objectNumber }?.subtype
        let update = AnnotationUpdate(objectNumber: objectNumber, bounds: .init(x: rect.minX, y: rect.minY, width: rect.width, height: rect.height))
        edit(resizing ? String(localized: "Resize Annotation") : String(localized: "Move Annotation"), { try await $0.updateAnnotation(update) },
             then: { session in
                 // The changed annotation can be written as a new object, so it stays selected by where it now is: the
                 // one of its kind that covers most of the rectangle it went to.
                 Task {
                     await session.loadAnnotations()
                     let candidates = session.pageAnnotations?.annotations.filter { $0.subtype == subtype } ?? []
                     let best = candidates.max { Self.overlap($0.bounds.cgRect, rect) < Self.overlap($1.bounds.cgRect, rect) }
                     if let best, Self.overlap(best.bounds.cgRect, rect) > 0.5 { session.selectedAnnotation = best.objectNumber }
                 }
             }, failed: failed)
    }

    /// How much two rectangles share, as a share of the smaller one.
    private static func overlap(_ a: CGRect, _ b: CGRect) -> CGFloat {
        let shared = a.intersection(b)
        guard !shared.isNull else { return 0 }
        return shared.width * shared.height / max(min(a.width * a.height, b.width * b.height), 1)
    }

    func removeSelectedAnnotation() {
        if let annotation = selectedPageAnnotation { removeAnnotation(annotation) }
    }

    /// Adds a polygon outlined by a drag, from points in the page's displayed points.
    func addPolygon(page: Int, points: [CGPoint], failed: (() -> Void)? = nil) {
        guard canEditAnnotations, points.count >= 3 else {
            failed?()
            return
        }
        let request = AnnotationRequest(tool: .polygon, page: page, strokes: [points.map { EnginePoint(x: $0.x, y: $0.y) }],
                                        author: Self.author, color: markupColor, strokeWidth: Self.shapeStrokeWidth)
        addMarkup([request], name: String(localized: "Add Shape"), failed: failed)
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

/// One drag across a page with a page tool, in the page layout's coordinates, which grow downward: drawing a shape,
/// marking an area to redact, or moving or resizing the selected annotation. The Mac and touch page views feed it
/// their pointer and show its outline.
@MainActor
struct PageDragGesture {
    enum Action {
        case draw(PageDragTool)
        case move(objectNumber: Int, original: CGRect)
        case resize(objectNumber: Int, fixed: CGPoint, original: CGRect)
    }

    /// How the outline looks while dragging.
    enum Look {
        case shape(dashed: Bool)
        case area
        case selection
    }

    let action: Action
    let page: Int
    let frame: CGRect
    /// The layout the drag started in, which turns layout points into points of the upright page.
    let layout: PageLayout
    private(set) var points: [CGPoint]

    private var start: CGPoint { points[0] }
    private var end: CGPoint { points[points.count - 1] }

    /// Starts a drag at a point of the page layout, or returns nil where the tool does nothing there. The select tool
    /// first selects the annotation under the point, or grabs a corner of the selected one.
    static func begin(_ tool: PageDragTool, at point: CGPoint, layout: PageLayout, session: DocumentSession) -> PageDragGesture? {
        guard let index = layout.frames.firstIndex(where: { $0.contains(point) }) else {
            if tool == .select { session.selectedAnnotation = nil }
            return nil
        }
        let frame = layout.frames[index]
        let page = index + 1
        guard tool == .select else { return PageDragGesture(action: .draw(tool), page: page, frame: frame, layout: layout, points: [point]) }
        let local = layout.pagePoint(point, page: page)
        if let selected = session.selectedPageAnnotation, selected.page == page,
           let fixed = SelectionOutline.fixedCorner(grabbing: local, of: selected.bounds.cgRect) {
            return PageDragGesture(action: .resize(objectNumber: selected.objectNumber, fixed: fixed, original: selected.bounds.cgRect),
                                   page: page, frame: frame, layout: layout, points: [point])
        }
        guard let hit = session.annotation(at: local, page: page) else {
            session.selectedAnnotation = nil
            return nil
        }
        session.selectedAnnotation = hit.objectNumber
        return PageDragGesture(action: .move(objectNumber: hit.objectNumber, original: hit.bounds.cgRect), page: page, frame: frame, layout: layout, points: [point])
    }

    /// Follows the pointer, kept on the page the drag started on.
    mutating func move(to point: CGPoint) {
        let point = CGPoint(x: min(max(point.x, frame.minX), frame.maxX), y: min(max(point.y, frame.minY), frame.maxY))
        if case .draw(.shape(.polygon)) = action {
            guard hypot(point.x - end.x, point.y - end.y) >= 2 else { return }
            points.append(point)
        } else if points.count == 1 {
            points.append(point)
        } else {
            points[1] = point
        }
    }

    var look: Look {
        switch action {
        case .draw(.redaction): .area
        case .draw(let tool): .shape(dashed: tool == .shape(.textBox))
        case .move, .resize: .selection
        }
    }

    /// The outline while dragging.
    var outline: CGPath {
        switch action {
        case .draw(let tool):
            return Self.path(for: tool, points: points)
        case .move, .resize:
            return CGPath(rect: layout.layoutRect(targetRect ?? .zero, page: page), transform: nil)
        }
    }

    /// Finishes the drag: adds the shape, marks the area, or moves or resizes the annotation. Returns whether the
    /// outline stays until the page's images show the change; `failed` removes it if the change doesn't go through.
    func finish(session: DocumentSession, failed: @escaping () -> Void) -> Bool {
        let from = local(start), to = local(end)
        switch action {
        case .draw(.redaction):
            session.markRedaction(page: page, from: from, to: to)
            return false
        case .draw(.shape(.polygon)):
            guard points.count >= 3 else { return false }
            session.addPolygon(page: page, points: points.map(local), failed: failed)
            return true
        case .draw(.shape(let shape)):
            session.addShape(shape, page: page, from: from, to: to, failed: failed)
            return true
        case .draw(.select):
            return false
        case .move(let number, let original), .resize(let number, _, let original):
            guard let rect = targetRect, abs(rect.minX - original.minX) >= 0.5 || abs(rect.minY - original.minY) >= 0.5
                    || abs(rect.width - original.width) >= 0.5 || abs(rect.height - original.height) >= 0.5 else { return false }
            var resizing = false
            if case .resize = action { resizing = true }
            session.setAnnotationBounds(number, rect, resizing: resizing, failed: failed)
            return true
        }
    }

    /// The rectangle, in the page's displayed points, that a move or resize puts the annotation in.
    private var targetRect: CGRect? {
        switch action {
        case .draw:
            return nil
        case .move(_, let original):
            let moved = layout.pageVector(CGVector(dx: end.x - start.x, dy: end.y - start.y))
            let size = layout.pageSize(page)
            var rect = original.offsetBy(dx: moved.dx, dy: moved.dy)
            rect.origin.x = min(max(rect.minX, 0), max(size.width - rect.width, 0))
            rect.origin.y = min(max(rect.minY, 0), max(size.height - rect.height, 0))
            return rect
        case .resize(_, let fixed, _):
            let corner = local(end)
            return CGRect(x: min(fixed.x, corner.x), y: min(fixed.y, corner.y),
                          width: max(abs(corner.x - fixed.x), 4), height: max(abs(corner.y - fixed.y), 4))
        }
    }

    private func local(_ point: CGPoint) -> CGPoint {
        layout.pagePoint(point, page: page)
    }

    private static func path(for tool: PageDragTool, points: [CGPoint]) -> CGPath {
        let path = CGMutablePath()
        let start = points[0], end = points[points.count - 1]
        let rect = CGRect(x: min(start.x, end.x), y: min(start.y, end.y), width: abs(end.x - start.x), height: abs(end.y - start.y))
        switch tool {
        case .shape(.oval):
            path.addEllipse(in: rect)
        case .shape(.polygon):
            path.addLines(between: points)
            if points.count > 2 { path.closeSubpath() }
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
}

/// The selected annotation's outline and corner handles, which a drag resizes from.
enum SelectionOutline {
    /// How far from a corner, in page points, a press grabs it.
    static let reach: CGFloat = 10

    /// The corner opposite the one a point grabs, which stays put while resizing.
    static func fixedCorner(grabbing point: CGPoint, of rect: CGRect) -> CGPoint? {
        let corners = corners(of: rect)
        guard let index = corners.firstIndex(where: { hypot($0.x - point.x, $0.y - point.y) <= reach }) else { return nil }
        return corners[3 - index]
    }

    /// The outline and the handles of a rectangle of an upright page, where the page layout shows it.
    static func paths(for rect: CGRect, page: Int, layout: PageLayout) -> (outline: CGPath, handles: CGPath) {
        let rect = layout.layoutRect(rect, page: page)
        let handles = CGMutablePath()
        for corner in corners(of: rect) { handles.addRect(CGRect(x: corner.x - 3.5, y: corner.y - 3.5, width: 7, height: 7)) }
        return (CGPath(rect: rect, transform: nil), handles)
    }

    /// Top left, top right, bottom left and bottom right, so the opposite of corner i is corner 3 - i.
    private static func corners(of rect: CGRect) -> [CGPoint] {
        [CGPoint(x: rect.minX, y: rect.minY), CGPoint(x: rect.maxX, y: rect.minY),
         CGPoint(x: rect.minX, y: rect.maxY), CGPoint(x: rect.maxX, y: rect.maxY)]
    }
}

/// Marked areas waiting to be redacted, in the page view's coordinates.
enum DragOutline {
    /// The marked areas, offset to their pages in the page layout.
    static func path(for areas: [RedactionMark], layout: PageLayout) -> CGPath {
        let path = CGMutablePath()
        for area in areas where layout.frames.indices.contains(area.page - 1) {
            path.addRect(layout.layoutRect(area.bounds.cgRect, page: area.page))
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
            if tool == .select {
                Button("Delete", systemImage: "trash", role: .destructive) { session.removeSelectedAnnotation() }
                    .labelStyle(.iconOnly)
                    .disabled(session.selectedPageAnnotation == nil)
            } else {
                ColorPicker("Colour", selection: markupColor, supportsOpacity: false)
                    .labelsHidden()
            }
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
