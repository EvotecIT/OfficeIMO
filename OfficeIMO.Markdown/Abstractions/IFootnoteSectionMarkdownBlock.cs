namespace OfficeIMO.Markdown;

internal interface IFootnoteSectionMarkdownBlock {
    string FootnoteLabel { get; }
    string RenderFootnoteSectionItemHtml();
}
