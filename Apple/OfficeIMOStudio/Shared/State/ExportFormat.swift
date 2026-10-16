/// The formats supported by the document export sheet and native engine.
enum ExportFormat: String, CaseIterable, Identifiable, Sendable {
    case word, excel, powerpoint, html, text, markdown, json, pageImages, images

    var id: Self { self }

    /// Accessibility text extraction does not grant permission to copy graphics or convert the complete document.
    func isAllowed(canExtractText: Bool, canExtractContent: Bool) -> Bool {
        switch self {
        case .text: canExtractText
        default: canExtractContent
        }
    }
}
