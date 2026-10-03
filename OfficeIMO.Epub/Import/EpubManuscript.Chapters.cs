using OfficeIMO.Html;

namespace OfficeIMO.Epub;

public static partial class EpubManuscript {
    private sealed class Chapter {
        internal string Title = string.Empty;
        internal string Path = string.Empty;
        internal XElement Body = new XElement(Xhtml + "body");
    }

    private static List<Chapter> SplitChapters(XElement body, int headingLevel, string title,
        List<OfficeConversionFidelityDiagnostic> diagnostics) {
        var chapters = new List<Chapter>();
        var ancestors = new List<XElement>();
        var targets = new List<XElement>();
        var current = new Chapter { Title = title, Body = new XElement(body.Name, body.Attributes()) };
        chapters.Add(current);
        targets.Add(current.Body);
        foreach (XNode node in body.Nodes()) Append(node);
        foreach (Chapter chapter in chapters) {
            foreach (XElement empty in chapter.Body.Descendants().Reverse().Where(element =>
                !element.Nodes().Any() && new[] { "section", "article", "main", "div" }.Contains(element.Name.LocalName)).ToArray()) empty.Remove();
        }
        chapters.RemoveAll(chapter => !HasContent(chapter.Body));
        if (chapters.Count == 0) throw new InvalidDataException("The manuscript has no readable body content.");
        for (int index = 0; index < chapters.Count; index++) chapters[index].Path = "EPUB/text/chapter-" + (index + 1).ToString("D4") + ".xhtml";
        return chapters;

        void Append(XNode node) {
            if (node is XElement element) {
                int level = HeadingLevel(element);
                bool canSplit = ancestors.All(parent => parent.Name.Namespace == Xhtml &&
                    new[] { "section", "article", "main", "div" }.Contains(parent.Name.LocalName));
                if (headingLevel != 0 && level != 0 && level <= headingLevel && canSplit) {
                    string label = element.Value.Trim();
                    if (HasContent(current.Body)) {
                        current = new Chapter { Title = label.Length == 0 ? title : label, Body = new XElement(body.Name, body.Attributes()) };
                        chapters.Add(current);
                        targets.Clear(); targets.Add(current.Body);
                        foreach (XElement parent in ancestors) {
                            var clone = new XElement(parent.Name, parent.Attributes().Where(attribute => attribute.Name != "id"));
                            targets[targets.Count - 1].Add(clone); targets.Add(clone);
                        }
                    } else if (label.Length > 0) current.Title = label;
                }
                var target = new XElement(element.Name, element.Attributes());
                targets[targets.Count - 1].Add(target);
                ancestors.Add(element); targets.Add(target);
                foreach (XNode child in element.Nodes()) Append(child);
                ancestors.RemoveAt(ancestors.Count - 1); targets.RemoveAt(targets.Count - 1);
            } else if (node is XText text) targets[targets.Count - 1].Add(new XText(text.Value));
        }
    }

    private static bool HasContent(XElement body) => !string.IsNullOrWhiteSpace(body.Value) ||
        body.Descendants().Any(element => new[] { "img", "svg", "math", "audio", "video", "hr" }.Contains(element.Name.LocalName));
    private static int HeadingLevel(XElement element) => element.Name.Namespace == Xhtml &&
        element.Name.LocalName.Length == 2 && element.Name.LocalName[0] == 'h' && element.Name.LocalName[1] >= '1' && element.Name.LocalName[1] <= '6'
            ? element.Name.LocalName[1] - '0' : 0;

    private static void AssignAnchors(List<Chapter> chapters, List<OfficeConversionFidelityDiagnostic> diagnostics) {
        var used = new HashSet<string>(StringComparer.Ordinal);
        int next = 0;
        foreach (XElement element in chapters.SelectMany(chapter => chapter.Body.DescendantsAndSelf())) {
            XAttribute? id = element.Attribute("id");
            if (id != null && (id.Value.Length == 0 || !used.Add(id.Value))) {
                AddDiagnostic(diagnostics, "EPUB_IMPORT_DUPLICATE_ANCHOR", "A duplicate or empty source anchor was removed; existing links select its first occurrence.", id.Value, OfficeConversionLossKind.Approximation);
                id.Remove();
            }
        }
        foreach (XElement heading in chapters.SelectMany(chapter => chapter.Body.Descendants()).Where(element => HeadingLevel(element) != 0)) {
            if (heading.Attribute("id") != null) continue;
            string id;
            do { id = "epub-heading-" + ++next; } while (!used.Add(id));
            heading.SetAttributeValue("id", id);
        }
    }

    private static void RewriteChapterLinks(List<Chapter> chapters, HtmlConversionDocument manuscript,
        List<OfficeConversionFidelityDiagnostic> diagnostics) {
        var anchors = chapters.SelectMany(chapter => chapter.Body.DescendantsAndSelf()
            .Where(element => element.Attribute("id") != null).Select(element => new { Id = (string)element.Attribute("id")!, Chapter = chapter }))
            .ToDictionary(item => item.Id, item => item.Chapter, StringComparer.Ordinal);
        foreach (Chapter chapter in chapters) {
            foreach (XAttribute href in chapter.Body.Descendants()
                .Where(element => element.Name == Xhtml + "a" || element.Name == Xhtml + "area" ||
                    element.Name == XName.Get("a", "http://www.w3.org/2000/svg"))
                .SelectMany(element => element.Attributes().Where(attribute => attribute.Name == "href" ||
                    attribute.Name == XName.Get("href", "http://www.w3.org/1999/xlink"))).ToArray()) {
                string value = href.Value.Trim();
                if (!value.StartsWith("#", StringComparison.Ordinal) && manuscript.BaseUri != null &&
                    Uri.TryCreate(manuscript.BaseUri, value, out Uri? absolute) && absolute.Fragment.Length != 0 &&
                    Uri.Compare(absolute, manuscript.BaseUri, UriComponents.SchemeAndServer | UriComponents.PathAndQuery,
                        UriFormat.SafeUnescaped, StringComparison.Ordinal) == 0) value = absolute.Fragment;
                if (value.StartsWith("#", StringComparison.Ordinal)) {
                    string id = Uri.UnescapeDataString(value.Substring(1));
                    if (anchors.TryGetValue(id, out Chapter? target)) href.Value = (target == chapter ? string.Empty : System.IO.Path.GetFileName(target.Path)) + "#" + Uri.EscapeDataString(id);
                    else {
                        AddDiagnostic(diagnostics, "EPUB_IMPORT_ANCHOR_MISSING", "A source hyperlink targets a missing anchor.", value, OfficeConversionLossKind.Failure);
                        href.Remove();
                    }
                } else {
                    string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(value, manuscript.BaseUri, manuscript.HyperlinkUrlPolicy);
                    if (string.IsNullOrWhiteSpace(resolved)) {
                        AddDiagnostic(diagnostics, "EPUB_IMPORT_HYPERLINK_OMITTED", "The source hyperlink was rejected by the shared URL policy.", value);
                        href.Remove();
                    } else href.Value = resolved;
                }
            }
        }
    }

    private sealed class TocNode {
        internal int Level;
        internal string Label = string.Empty;
        internal string Href = string.Empty;
        internal List<TocNode> Children = new List<TocNode>();
        internal EpubNavigationEntry Freeze() => new EpubNavigationEntry(Label, Href, Children.Select(child => child.Freeze()));
    }

    private static IEnumerable<EpubNavigationEntry> BuildNavigation(List<Chapter> chapters) {
        foreach (Chapter chapter in chapters) {
            var root = new TocNode { Label = chapter.Title, Href = chapter.Path };
            var stack = new List<TocNode> { root };
            XElement[] headings = chapter.Body.Descendants().Where(element => HeadingLevel(element) != 0).ToArray();
            bool first = true;
            foreach (XElement heading in headings) {
                string label = heading.Value.Trim();
                if (label.Length == 0) continue;
                int level = HeadingLevel(heading);
                string href = chapter.Path + "#" + Uri.EscapeDataString((string)heading.Attribute("id")!);
                if (first && label == chapter.Title) {
                    root.Level = level; root.Href = href; first = false; continue;
                }
                first = false;
                while (stack.Count > 1 && stack[stack.Count - 1].Level >= level) stack.RemoveAt(stack.Count - 1);
                var node = new TocNode { Level = level, Label = label, Href = href };
                stack[stack.Count - 1].Children.Add(node); stack.Add(node);
            }
            yield return root.Freeze();
        }
    }
}
