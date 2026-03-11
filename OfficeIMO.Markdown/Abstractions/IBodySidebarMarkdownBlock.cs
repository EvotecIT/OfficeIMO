namespace OfficeIMO.Markdown;

internal interface IBodySidebarMarkdownBlock : IContextualHtmlMarkdownBlock {
    bool UsesSidebarLayout();
    bool SuppressesPrecedingHeadingTitle();
    string WrapSidebarLayoutHtml(string navigationHtml, string contentHtml);
}
