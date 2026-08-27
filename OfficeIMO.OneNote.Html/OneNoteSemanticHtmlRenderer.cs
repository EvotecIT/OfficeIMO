using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OneNote.Markdown;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;

namespace OfficeIMO.OneNote.Html;

internal static class OneNoteSemanticHtmlRenderer {
    internal static MarkdownDoc CreateDocument(OneNoteSection section, OneNoteMarkdownOptions options) {
        var html = new StringBuilder(4096);
        html.Append("<article class=\"officeimo-onenote-semantic\">");
        AppendHeading(html, options.HeadingLevel, Name(section.Name, "OneNote section"));
        foreach (OneNotePage page in section.Pages) AppendPageWithRelated(html, page, Math.Min(6, options.HeadingLevel + 1), null, options, 0);
        html.Append("</article>");
        return MarkdownDoc.Create().Add(new GeneratedHtmlBlock(html.ToString()));
    }

    internal static MarkdownDoc CreateDocument(OneNoteNotebook notebook, OneNoteMarkdownOptions options) {
        var html = new StringBuilder(4096);
        html.Append("<article class=\"officeimo-onenote-semantic\">");
        AppendHeading(html, options.HeadingLevel, Name(notebook.Name, "OneNote notebook"));
        AppendHierarchy(html, notebook.Sections, notebook.SectionGroups, Math.Min(6, options.HeadingLevel + 1), options, 0);
        html.Append("</article>");
        return MarkdownDoc.Create().Add(new GeneratedHtmlBlock(html.ToString()));
    }

    private static void AppendHierarchy(StringBuilder html, IList<OneNoteSection> sections, IList<OneNoteSectionGroup> groups,
        int headingLevel, OneNoteMarkdownOptions options, int depth) {
        if (depth >= options.MaxSectionGroupDepth) return;
        foreach (HierarchyItem item in Order(sections, groups)) {
            if (item.Section != null) {
                AppendHeading(html, headingLevel, Name(item.Section.Name, "OneNote section"));
                foreach (OneNotePage page in item.Section.Pages) AppendPageWithRelated(html, page, Math.Min(6, headingLevel + 1), null, options, 0);
            } else {
                OneNoteSectionGroup group = item.Group!;
                AppendHeading(html, headingLevel, Name(group.Name, "Section group"));
                AppendHierarchy(html, group.Sections, group.SectionGroups, Math.Min(6, headingLevel + 1), options, depth + 1);
            }
        }
    }

    private static IEnumerable<HierarchyItem> Order(IList<OneNoteSection> sections, IList<OneNoteSectionGroup> groups) {
        int sequence = 0;
        var items = new List<HierarchyItem>(sections.Count + groups.Count);
        foreach (OneNoteSection section in sections) items.Add(new HierarchyItem(section, sequence++));
        foreach (OneNoteSectionGroup group in groups) items.Add(new HierarchyItem(group, sequence++));
        return items.OrderBy(item => item.Order.HasValue ? 0 : 1).ThenBy(item => item.Order).ThenBy(item => item.Sequence);
    }

    private static void AppendPageWithRelated(StringBuilder html, OneNotePage page, int headingLevel, string? prefix,
        OneNoteMarkdownOptions options, int depth) {
        if (depth >= options.MaxPageRelationshipDepth) return;
        AppendHeading(html, headingLevel, string.IsNullOrEmpty(prefix) ? Name(page.Title, "Untitled page") : prefix + ": " + Name(page.Title, "Untitled page"));
        html.Append("<section class=\"officeimo-onenote-page\">");
        AppendElements(html, page.Outlines.Cast<OneNoteElement>().Concat(page.DirectContent), options, 0);
        html.Append("</section>");
        int relatedLevel = Math.Min(6, headingLevel + 1);
        if (options.IncludeConflictPages) {
            foreach (OneNotePage conflict in page.ConflictPages) AppendPageWithRelated(html, conflict, relatedLevel, "Conflict", options, depth + 1);
        }
        if (options.IncludeVersionHistory) {
            foreach (OneNotePage version in page.VersionHistory) AppendPageWithRelated(html, version, relatedLevel, "Version", options, depth + 1);
        }
    }

    private static void AppendElement(StringBuilder html, OneNoteElement element, OneNoteMarkdownOptions options, int depth) {
        if (depth >= options.MaxContentDepth) return;
        switch (element) {
            case OneNoteOutline outline:
                html.Append("<section class=\"officeimo-onenote-outline\">");
                AppendElements(html, outline.Children, options, depth + 1);
                html.Append("</section>");
                break;
            case OneNoteParagraph paragraph:
                AppendParagraph(html, paragraph, options, depth);
                break;
            case OneNoteTable table:
                AppendTable(html, table, options, depth);
                break;
            case OneNoteImage image:
                AppendImage(html, image, options.AssetUriResolver);
                break;
            case OneNoteEmbeddedFile file:
                AppendBinary(html, file, file.FileName ?? "attachment", options.AssetUriResolver);
                break;
            case OneNoteMedia media:
                AppendBinary(html, media, media.FileName ?? "recording", options.AssetUriResolver);
                break;
            case OneNoteInk ink:
                AppendBinary(html, ink, "Ink", options.AssetUriResolver);
                break;
            case OneNoteMath math:
                html.Append("<pre class=\"officeimo-onenote-math\"><code>").Append(Text(math.Latex ?? math.Text ?? math.MathMl)).Append("</code></pre>");
                break;
        }
    }

    private static void AppendParagraph(StringBuilder html, OneNoteParagraph paragraph, OneNoteMarkdownOptions options, int depth) {
        if (paragraph.List != null) {
            AppendList(html, new[] { paragraph }, options, depth);
            return;
        }
        string tag = HeadingTag(paragraph.Style.StyleId);
        html.Append('<').Append(tag).Append(ParagraphStyle(paragraph.Style)).Append('>');
        foreach (OneNoteTextRun run in paragraph.Runs) AppendRun(html, run);
        html.Append("</").Append(tag).Append('>');
        AppendElements(html, paragraph.Children, options, depth + 1);
    }

    private static void AppendElements(
        StringBuilder html,
        IEnumerable<OneNoteElement> elements,
        OneNoteMarkdownOptions options,
        int depth) {
        if (depth >= options.MaxContentDepth) return;
        OneNoteElement[] items = elements.ToArray();
        for (int index = 0; index < items.Length;) {
            if (items[index] is OneNoteParagraph paragraph && paragraph.List != null) {
                bool ordered = paragraph.List.Ordered;
                int level = paragraph.List.Level;
                var group = new List<OneNoteParagraph>();
                while (index < items.Length
                       && items[index] is OneNoteParagraph candidate
                       && candidate.List != null
                       && candidate.List.Ordered == ordered
                       && candidate.List.Level == level) {
                    group.Add(candidate);
                    index++;
                }
                AppendList(html, group, options, depth);
                continue;
            }
            AppendElement(html, items[index], options, depth);
            index++;
        }
    }

    private static void AppendList(
        StringBuilder html,
        IReadOnlyList<OneNoteParagraph> paragraphs,
        OneNoteMarkdownOptions options,
        int depth) {
        if (paragraphs.Count == 0) return;
        OneNoteListInfo list = paragraphs[0].List!;
        string tag = list.Ordered ? "ol" : "ul";
        html.Append('<').Append(tag).Append(" data-level=\"")
            .Append(list.Level.ToString(CultureInfo.InvariantCulture)).Append("\">");
        foreach (OneNoteParagraph paragraph in paragraphs) {
            html.Append("<li>");
            foreach (OneNoteTextRun run in paragraph.Runs) AppendRun(html, run);
            AppendElements(html, paragraph.Children, options, depth + 1);
            html.Append("</li>");
        }
        html.Append("</").Append(tag).Append('>');
    }

    private static void AppendRun(StringBuilder html, OneNoteTextRun run) {
        bool linked = !string.IsNullOrWhiteSpace(run.Hyperlink);
        string hyperlink = linked ? SafeLink(run.Hyperlink) : string.Empty;
        if (hyperlink.Length > 0) html.Append("<a href=\"").Append(Attribute(hyperlink)).Append("\">");
        if (run.Style.Bold == true) html.Append("<strong>");
        if (run.Style.Italic == true) html.Append("<em>");
        if (run.Style.Underline == true) html.Append("<u>");
        if (run.Style.Strikethrough == true) html.Append("<s>");
        string style = RunStyle(run.Style);
        string? tag = run.Style.Superscript == true ? "sup" : run.Style.Subscript == true ? "sub" : style.Length > 0 ? "span" : null;
        if (tag != null) {
            html.Append('<').Append(tag);
            if (style.Length > 0) html.Append(" style=\"").Append(Attribute(style)).Append('"');
            html.Append('>');
        }
        if (run.MathExpression != null || run.Style.IsMath == true) {
            string math = run.MathExpression != null
                ? OfficeIMO.Drawing.OfficeMathMarkup.ToLatex(run.MathExpression)
                : OneNoteTextProjection.Normalize(run.Text);
            html.Append("<code class=\"officeimo-onenote-math\" data-officeimo-math-format=\"latex\">")
                .Append(Text(math))
                .Append("</code>");
        } else {
            string normalized = OneNoteTextProjection.Normalize(run.Text);
            html.Append(Text(normalized).Replace("\r\n", "<br>").Replace("\r", "<br>").Replace("\n", "<br>"));
        }
        if (tag != null) html.Append("</").Append(tag).Append('>');
        if (run.Style.Strikethrough == true) html.Append("</s>");
        if (run.Style.Underline == true) html.Append("</u>");
        if (run.Style.Italic == true) html.Append("</em>");
        if (run.Style.Bold == true) html.Append("</strong>");
        if (hyperlink.Length > 0) html.Append("</a>");
    }

    private static void AppendTable(StringBuilder html, OneNoteTable table, OneNoteMarkdownOptions options, int depth) {
        html.Append("<table").Append(table.BordersVisible ? " style=\"border-collapse:collapse\"" : string.Empty).Append('>');
        foreach (OneNoteTableRow row in table.Rows) {
            html.Append("<tr>");
            foreach (OneNoteTableCell cell in row.Cells) {
                html.Append("<td");
                if (cell.ShadingColorArgb.HasValue) html.Append(" style=\"background-color:").Append(Color(cell.ShadingColorArgb.Value)).Append("\"");
                html.Append('>');
                AppendElements(html, cell.Content, options, depth + 1);
                html.Append("</td>");
            }
            html.Append("</tr>");
        }
        html.Append("</table>");
    }

    private static void AppendImage(StringBuilder html, OneNoteImage image, Func<OneNoteBinaryElement, string?>? resolver) {
        string source = image.Payload == null ? string.Empty : SafeResource(resolver?.Invoke(image));
        string label = image.AltText ?? image.FileName ?? "image";
        string hyperlink = SafeLink(image.Hyperlink);
        if (hyperlink.Length > 0) html.Append("<a href=\"").Append(Attribute(hyperlink)).Append("\">");
        if (source.Length == 0) html.Append("<span class=\"officeimo-onenote-image-placeholder\">[Image: ").Append(Text(label)).Append("]</span>");
        else html.Append("<img src=\"").Append(Attribute(source)).Append("\" alt=\"").Append(Attribute(label)).Append("\">");
        if (hyperlink.Length > 0) html.Append("</a>");
    }

    private static void AppendBinary(StringBuilder html, OneNoteBinaryElement element, string label, Func<OneNoteBinaryElement, string?>? resolver) {
        string target = element.Payload == null ? string.Empty : SafeResource(resolver?.Invoke(element));
        if (target.Length == 0) html.Append("<span class=\"officeimo-onenote-attachment\">[").Append(Text(label)).Append("]</span>");
        else html.Append("<a class=\"officeimo-onenote-attachment\" href=\"").Append(Attribute(target)).Append("\">[").Append(Text(label)).Append("]</a>");
    }

    private static string RunStyle(OneNoteTextStyle style) {
        var css = new List<string>();
        if (!string.IsNullOrWhiteSpace(style.FontFamily)) css.Add("font-family:\"" + CssString(style.FontFamily!) + "\"");
        if (style.FontSize.HasValue) css.Add("font-size:" + style.FontSize.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt");
        if (style.ColorArgb.HasValue) css.Add("color:" + Color(style.ColorArgb.Value));
        if (style.HighlightColorArgb.HasValue) css.Add("background-color:" + Color(style.HighlightColorArgb.Value));
        return string.Join(";", css);
    }

    private static string ParagraphStyle(OneNoteParagraphStyle style) {
        if (!style.Alignment.HasValue) return string.Empty;
        string value = style.Alignment.Value.ToString().ToLowerInvariant();
        return " style=\"text-align:" + value + "\"";
    }

    private static string HeadingTag(string? styleId) {
        if (!string.IsNullOrWhiteSpace(styleId) && styleId!.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(styleId.Substring(7), NumberStyles.Integer, CultureInfo.InvariantCulture, out int level)) return "h" + Math.Max(1, Math.Min(6, level));
        return "p";
    }

    private static void AppendHeading(StringBuilder html, int level, string value) {
        string tag = "h" + Math.Max(1, Math.Min(6, level)).ToString(CultureInfo.InvariantCulture);
        html.Append('<').Append(tag).Append('>').Append(Text(value)).Append("</").Append(tag).Append('>');
    }

    private static string SafeLink(string? value) => HtmlUrlPolicyEvaluator.ResolveUrl(EncodeUrl(value), null, HtmlUrlPolicy.CreateHyperlinkProfile());
    private static string SafeResource(string? value) => HtmlUrlPolicyEvaluator.ResolveUrl(EncodeUrl(value), null, HtmlUrlPolicy.CreateOfficeIMOProfile());
    private static string Name(string? value, string fallback) => OneNoteTextProjection.Normalize(string.IsNullOrWhiteSpace(value) ? fallback : value);
    private static string Text(string? value) => WebUtility.HtmlEncode(OneNoteTextProjection.Normalize(value));
    private static string Attribute(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);
    private static string CssString(string value) => value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ");
    private static string Color(uint argb) => "#" + (argb & 0x00FFFFFFU).ToString("X6", CultureInfo.InvariantCulture);

    private static string EncodeUrl(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var encoded = new StringBuilder(value!.Length);
        foreach (char character in value) {
            if (character <= 0x20 || character == '\'' || character == '"' || character == '<' || character == '>'
                || character == '\\' || character == '[' || character == ']' || character == '^' || character == '`'
                || character == '{' || character == '|' || character == '}') {
                foreach (byte item in Encoding.UTF8.GetBytes(character.ToString())) encoded.Append('%').Append(item.ToString("X2", CultureInfo.InvariantCulture));
            } else {
                encoded.Append(character);
            }
        }
        return encoded.ToString();
    }

    private sealed class GeneratedHtmlBlock : IMarkdownBlock {
        private readonly string _html;
        internal GeneratedHtmlBlock(string html) => _html = html;
        public string RenderMarkdown() => _html;
        public string RenderHtml() => _html;
    }

    private sealed class HierarchyItem {
        internal HierarchyItem(OneNoteSection section, int sequence) { Section = section; Order = section.TableOfContentsOrder; Sequence = sequence; }
        internal HierarchyItem(OneNoteSectionGroup group, int sequence) { Group = group; Order = group.TableOfContentsOrder; Sequence = sequence; }
        internal OneNoteSection? Section { get; }
        internal OneNoteSectionGroup? Group { get; }
        internal uint? Order { get; }
        internal int Sequence { get; }
    }
}
