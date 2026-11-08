namespace OfficeIMO.Html;

public sealed partial class HtmlConversionDocument {
    private const string ThemeStyleAttribute = "data-officeimo-theme";

    /// <summary>Returns an independent document with theme defaults for unstyled content. Authored CSS and inline formatting take precedence.</summary>
    /// <remarks>Content, resource policies, base URI and limits are retained. Calling this again replaces only the earlier theme stylesheet.</remarks>
    public HtmlConversionDocument WithTheme(OfficeTheme theme) => WithTheme(theme, string.Empty);

    /// <summary>Retains format-owned page defaults in the same lower-priority theme layer.</summary>
    internal HtmlConversionDocument WithTheme(OfficeTheme theme, string pageDefaults) {
        ArgumentNullException.ThrowIfNull(theme);
        return Edit(document => {
            foreach (Dom.HtmlElement previous in document.QuerySelectorAll("style[" + ThemeStyleAttribute + "]")) previous.Remove();
            Dom.HtmlElement? root = document.DocumentElement;
            if (root == null) {
                root = document.CreateElement("html");
                document.AppendChild(root);
            }
            if (root.NamespaceUri != Dom.HtmlElement.HtmlNamespace || root.LocalName != "html")
                throw new NotSupportedException("HTML theme defaults require an HTML document or fragment.");
            Dom.HtmlElement? head = document.Head;
            if (head == null) {
                head = document.CreateElement("head");
                Dom.HtmlNode[] content = root.ChildNodes.ToArray();
                root.AppendChild(head);
                foreach (Dom.HtmlNode node in content) root.AppendChild(node);
            }
            if (document.Body == null) root.AppendChild(document.CreateElement("body"));
            Dom.HtmlElement style = document.CreateElement("style");
            style.SetAttribute(ThemeStyleAttribute, theme.Name);
            // The first, lowest-priority layer lets every authored stylesheet retain precedence,
            // including selectors less specific than the semantic defaults below.
            style.TextContent = "@layer officeimo-theme {\n" + (theme.IsDark ? ":where(html){color-scheme:dark;}\n" : string.Empty) +
                OfficeHtmlThemeDefaults.Build(theme, ":where(html)") + pageDefaults + "\n}";
            Dom.HtmlNode[] authored = head.ChildNodes.ToArray();
            head.AppendChild(style);
            foreach (Dom.HtmlNode node in authored) head.AppendChild(node);
        });
    }
}
