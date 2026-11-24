import CoreGraphics
import Foundation

/// One thing VoiceOver finds on a page: the page itself, a line of its text, a form field or a comment. Rectangles are
/// in points of the page as displayed; each platform's page view places them through the page layout.
struct PageAccessibilityItem: Equatable {
    enum Kind: Equatable {
        case page
        case text
        case field(FormField.Kind, name: String, widget: Int)
        case comment(id: String)
    }

    let kind: Kind
    let page: Int
    let rect: CGRect
    var label: String
    var value: String?
    var hint: String?
    /// Whether a check box or radio button is on.
    var isOn: Bool?
}

extension DocumentSession {
    /// What VoiceOver reads on the pages around the one in view, in reading order: each page, then its lines of text
    /// with its form fields and comments where they sit. Pages whose text isn't read yet offer the page alone; pages the
    /// layout doesn't show, such as the others in Single Page, offer nothing.
    func accessibilityItems(around page: Int, in layout: PageLayout) -> [PageAccessibilityItem] {
        guard isReady, let facts else { return [] }
        var items: [PageAccessibilityItem] = []
        for number in [page - 1, page, page + 1] where facts.pages.indices.contains(number - 1) && layout.shows(number) {
            let size = facts.pages[number - 1].displaySize
            items.append(PageAccessibilityItem(kind: .page, page: number, rect: CGRect(origin: .zero, size: size),
                                               label: String(localized: "Page \(number) of \(facts.pages.count)"),
                                               value: pageText(number)?.count == 0 ? String(localized: "No text") : nil))
            var lines: [(rect: CGRect, text: String)] = []
            if let text = pageText(number) {
                for range in ReadingOrder.lines(breaks: text.breaks.map(\.rawValue)) {
                    let string = text.string(range)
                    let rects = text.lineRects(range)
                    guard !string.isEmpty, let first = rects.first else { continue }
                    lines.append((rects.dropFirst().reduce(first) { $0.union($1) }, string))
                }
            }
            let others = fieldItems(page: number) + commentItems(page: number)
            for item in ReadingOrder.merge(lines: lines.map(\.rect), others: others.map(\.rect)) {
                switch item {
                case .line(let index): items.append(PageAccessibilityItem(kind: .text, page: number, rect: lines[index].rect, label: lines[index].text))
                case .other(let index): items.append(others[index])
                }
            }
        }
        return items
    }

    private func fieldItems(page: Int) -> [PageAccessibilityItem] {
        guard let fields = formFields(page) else { return [] }
        return fields.fields.flatMap { field in
            field.widgets.map { widget in
                var item = PageAccessibilityItem(kind: .field(field.kind, name: field.name, widget: widget.objectNumber), page: page,
                                                 rect: widget.bounds.cgRect, label: field.displayName)
                switch field.kind {
                case .text:
                    item.value = field.password ? String(localized: "Secure text") : field.text
                case .checkbox, .radio:
                    item.value = widget.isOn ? String(localized: "On") : String(localized: "Off")
                    item.isOn = widget.isOn
                    if field.kind == .radio, let on = widget.onState {
                        item.label = "\(field.displayName), \(field.options.first { $0.value == on }?.text ?? on)"
                    }
                case .combo, .list:
                    item.value = field.values.map { value in field.options.first { $0.value == value }?.text ?? value }
                        .formatted(.list(type: .and))
                case .signature:
                    item.value = String(localized: "Signature field")
                default:
                    break
                }
                if field.readOnly { item.hint = String(localized: "Read only") }
                else if field.required { item.hint = String(localized: "Required") }
                return item
            }
        }
    }

    private func commentItems(page: Int) -> [PageAccessibilityItem] {
        guard let comments = currentComments else { return [] }
        return comments.threads.compactMap { thread in
            guard thread.page == page, let box = thread.bounds, let comment = thread.comment else { return nil }
            let author = comment.author.flatMap { $0.isEmpty ? nil : $0 } ?? String(localized: "Unknown Author")
            let replies = thread.replies.count
            var item = PageAccessibilityItem(kind: .comment(id: thread.id), page: page,
                                             rect: CGRect(x: box.x, y: box.y, width: box.width, height: box.height),
                                             label: String(localized: "Comment from \(author)"), value: comment.contents)
            let notes = [replies > 0 ? String(localized: "Replies: \(replies)") : nil, thread.resolved ? String(localized: "Resolved") : nil]
            item.hint = notes.contains { $0 != nil } ? notes.compactMap { $0 }.joined(separator: ", ") : nil
            return item
        }
    }
}
