using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

/// <summary>Builds the canonical typed semantic document consumed by generic target adapters.</summary>
internal static class HtmlSemanticDocumentBuilder {
    internal static HtmlSemanticDocument FromDocument(
        IHtmlDocument document,
        HtmlCssMediaContext mediaContext,
        HtmlConversionLimits limits) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (limits == null) throw new ArgumentNullException(nameof(limits));

        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles =
            HtmlComputedStyleEngine.Compute(document, mediaContext, limits);
        var sections = new List<HtmlSemanticSection>();
        var resources = new List<HtmlSemanticResource>();
        foreach (HtmlGenericSectionProjection projection in HtmlGenericDocumentProjector.CreateSections(document)) {
            var blocks = new List<HtmlSemanticBlock>();
            foreach (IElement element in HtmlGenericDocumentProjector.EnumerateBlocks(projection)) {
                HtmlSemanticBlock block = BuildBlock(document, element, styles, resources, null);
                blocks.Add(block);
            }

            HtmlSemanticSourceLocation? location = blocks.FirstOrDefault()?.SourceLocation;
            sections.Add(new HtmlSemanticSection(projection.Title, blocks.AsReadOnly(), location));
        }

        var rootTables = new List<HtmlSemanticBlock>();
        int tableIndex = 0;
        foreach (IElement table in HtmlGenericDocumentProjector.SelectRootTables(document)) {
            tableIndex++;
            rootTables.Add(BuildBlock(
                document,
                table,
                styles,
                resources,
                HtmlGenericDocumentProjector.GetTableTitle(document, table, tableIndex)));
        }

        IReadOnlyDictionary<string, string> metadata = ReadMetadata(document);
        string language = Normalize(document.DocumentElement?.GetAttribute("lang"));
        return new HtmlSemanticDocument(
            Normalize(document.Title),
            language,
            metadata,
            sections.AsReadOnly(),
            rootTables.AsReadOnly(),
            DeduplicateResources(resources));
    }

    private static HtmlSemanticBlock BuildBlock(
        IHtmlDocument document,
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<HtmlSemanticResource> resources,
        string? tableTitle) {
        HtmlSemanticBlockKind kind = GetKind(element);
        HtmlSemanticSourceLocation? location = HtmlSemanticSourceLocation.FromElement(element);
        styles.TryGetValue(element, out HtmlComputedStyle? style);
        IReadOnlyList<HtmlSemanticRun> runs = IsTextual(kind)
            ? BuildRuns(element, styles)
            : Array.Empty<HtmlSemanticRun>();
        var children = new List<HtmlSemanticBlock>();
        bool ordered = Is(element, "ol");
        int level = kind == HtmlSemanticBlockKind.Heading ? GetHeadingLevel(element)
            : kind == HtmlSemanticBlockKind.List ? GetListDepth(element)
            : 0;

        if (kind == HtmlSemanticBlockKind.List) {
            foreach (IElement item in element.Children.Where(child => Is(child, "li") || Is(child, "dt") || Is(child, "dd"))) {
                children.Add(BuildListItem(document, item, styles, resources, level));
            }
        }
        if (kind == HtmlSemanticBlockKind.Form) {
            foreach (IElement control in element.QuerySelectorAll("input, select, textarea, button")) {
                if (ReferenceEquals(control, element)) continue;
                children.Add(BuildBlock(document, control, styles, resources, null));
            }
        }

        HtmlSemanticTable? table = kind == HtmlSemanticBlockKind.Table
            ? BuildTable(element, styles, resources, tableTitle ?? HtmlGenericDocumentProjector.GetTableTitle(document, element, 1))
            : null;
        HtmlSemanticResource? resource = BuildResource(element, kind, style, location);
        if (resource != null) resources.Add(resource);
        IReadOnlyList<HtmlSemanticResource> inlineResources = kind == HtmlSemanticBlockKind.Table
            || kind == HtmlSemanticBlockKind.List || kind == HtmlSemanticBlockKind.Form
            ? Array.Empty<HtmlSemanticResource>()
            : BuildInlineResources(element, styles, resources);
        if (kind == HtmlSemanticBlockKind.Media) {
            foreach (IElement sourceElement in element.QuerySelectorAll("source[src], track[src]")) {
                string childSource = (sourceElement.GetAttribute("src") ?? string.Empty).Trim();
                if (childSource.Length == 0) continue;
                string mediaType = (sourceElement.GetAttribute("type") ?? string.Empty).Trim();
                resources.Add(new HtmlSemanticResource(
                    HtmlResourceKind.Media,
                    childSource,
                    HtmlAccessibilitySemantics.GetAccessibleName(element, includeTextFallback: true),
                    mediaType,
                    null,
                    null,
                    HtmlSemanticSourceLocation.FromElement(sourceElement)));
            }
        }
        HtmlSemanticFormControl? form = BuildFormControl(element);

        return new HtmlSemanticBlock(
            kind,
            IsTextual(kind) ? string.Concat(runs.Select(run => run.Text)) : HtmlGenericDocumentProjector.GetBlockText(element),
            level,
            ordered,
            runs,
            children.AsReadOnly(),
            table,
            resource,
            inlineResources,
            form,
            style,
            location,
            element);
    }

    private static HtmlSemanticBlock BuildListItem(
        IHtmlDocument document,
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<HtmlSemanticResource> resources,
        int level) {
        styles.TryGetValue(element, out HtmlComputedStyle? style);
        IReadOnlyList<HtmlSemanticRun> runs = BuildRuns(element, styles, skipNestedLists: true);
        IReadOnlyList<HtmlSemanticResource> inlineResources = BuildInlineResources(element, styles, resources,
            skipNestedLists: true);
        var children = new List<HtmlSemanticBlock>();
        foreach (IElement nestedList in element.Children.Where(child => Is(child, "ul") || Is(child, "ol") || Is(child, "dl"))) {
            children.Add(BuildBlock(document, nestedList, styles, resources, null));
        }

        return new HtmlSemanticBlock(
            HtmlSemanticBlockKind.ListItem,
            string.Concat(runs.Select(run => run.Text)),
            level,
            ordered: false,
            runs,
            children.AsReadOnly(),
            table: null,
            resource: null,
            inlineResources,
            formControl: null,
            style,
            HtmlSemanticSourceLocation.FromElement(element),
            element);
    }

    private static HtmlSemanticTable BuildTable(
        IElement table,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<HtmlSemanticResource> resources,
        string title) {
        var rows = new List<HtmlSemanticTableRow>();
        foreach (IElement rowElement in DirectRows(table)) {
            var cells = new List<HtmlSemanticTableCell>();
            foreach (IElement cell in rowElement.Children.Where(IsTableCell)) {
                styles.TryGetValue(cell, out HtmlComputedStyle? style);
                IReadOnlyList<HtmlSemanticRun> runs = BuildRuns(cell, styles);
                IReadOnlyList<HtmlSemanticResource> cellResources = BuildInlineResources(cell, styles, resources);
                cells.Add(new HtmlSemanticTableCell(
                    string.Concat(runs.Select(run => run.Text)),
                    Is(cell, "th"),
                    ReadSpan(cell, "rowspan"),
                    ReadSpan(cell, "colspan"),
                    runs,
                    cellResources,
                    style,
                    HtmlSemanticSourceLocation.FromElement(cell)));
            }
            if (cells.Count > 0) rows.Add(new HtmlSemanticTableRow(cells.AsReadOnly(), HtmlSemanticSourceLocation.FromElement(rowElement)));
        }
        return new HtmlSemanticTable(title, rows.AsReadOnly());
    }

    private static IReadOnlyList<HtmlSemanticRun> BuildRuns(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        bool skipNestedLists = false) {
        var runs = new List<HtmlSemanticRun>();
        AppendRuns(element, styles, runs, default, skipNestedLists, isRoot: true);
        styles.TryGetValue(element, out HtmlComputedStyle? rootStyle);
        string whiteSpace = (rootStyle?.GetValue("white-space") ?? string.Empty).Trim().ToLowerInvariant();
        bool preserve = Is(element, "pre") || whiteSpace == "pre" || whiteSpace == "pre-wrap" || whiteSpace == "break-spaces";
        bool preserveLines = preserve || whiteSpace == "pre-line";
        return NormalizeRuns(runs, preserve, preserveLines);
    }

    private static void AppendRuns(
        INode node,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<HtmlSemanticRun> runs,
        InlineState state,
        bool skipNestedLists,
        bool isRoot = false) {
        if (node is IText text) {
            if (text.Data.Length == 0) return;
            runs.Add(new HtmlSemanticRun(text.Data, state.Hyperlink, state.Bold, state.Italic,
                state.Underline, state.Strikethrough, state.Superscript, state.Subscript,
                state.Style, state.Location, isLineBreak: false));
            return;
        }
        if (!(node is IElement element)) return;
        if (!isRoot && skipNestedLists && (Is(element, "ul") || Is(element, "ol") || Is(element, "dl"))) return;
        if (Is(element, "br")) {
            runs.Add(new HtmlSemanticRun("\n", state.Hyperlink, state.Bold, state.Italic,
                state.Underline, state.Strikethrough, state.Superscript, state.Subscript,
                state.Style, HtmlSemanticSourceLocation.FromElement(element), isLineBreak: true));
            return;
        }

        styles.TryGetValue(element, out HtmlComputedStyle? style);
        InlineState nested = state.With(element, style, HtmlSemanticSourceLocation.FromElement(element));
        foreach (INode child in element.ChildNodes) {
            AppendRuns(child, styles, runs, nested, skipNestedLists);
        }
    }

    private static HtmlSemanticRun CopyRun(HtmlSemanticRun run, string text) =>
        new HtmlSemanticRun(text, run.Hyperlink, run.Bold, run.Italic, run.Underline,
            run.Strikethrough, run.Superscript, run.Subscript, run.Style, run.SourceLocation, run.IsLineBreak);

    private static IReadOnlyList<HtmlSemanticRun> NormalizeRuns(
        IReadOnlyList<HtmlSemanticRun> source,
        bool preserveWhitespace,
        bool preserveLines) {
        if (preserveWhitespace) return source.Where(run => run.Text.Length > 0).ToList().AsReadOnly();
        var result = new List<HtmlSemanticRun>();
        bool hasVisibleText = false;
        bool pendingSpace = false;
        foreach (HtmlSemanticRun run in source) {
            if (run.IsLineBreak) {
                TrimTrailingSpace(result);
                if (result.Count > 0 && result[result.Count - 1].Text != "\n") result.Add(CopyRun(run, "\n"));
                hasVisibleText = false;
                pendingSpace = false;
                continue;
            }
            var builder = new StringBuilder(run.Text.Length);
            foreach (char character in run.Text) {
                bool lineBreak = character == '\r' || character == '\n';
                if (lineBreak && preserveLines) {
                    while (builder.Length > 0 && builder[builder.Length - 1] == ' ') builder.Length--;
                    if (builder.Length > 0 || result.Count > 0) builder.Append('\n');
                    hasVisibleText = false;
                    pendingSpace = false;
                } else if (char.IsWhiteSpace(character)) {
                    pendingSpace = hasVisibleText;
                } else {
                    if (pendingSpace && builder.Length == 0 && result.Count > 0
                        && !EndsWithWhitespace(result[result.Count - 1].Text)) builder.Append(' ');
                    else if (pendingSpace && builder.Length > 0 && builder[builder.Length - 1] != '\n') builder.Append(' ');
                    builder.Append(character);
                    hasVisibleText = true;
                    pendingSpace = false;
                }
            }
            if (pendingSpace && hasVisibleText) {
                if (builder.Length > 0 && builder[builder.Length - 1] != '\n') builder.Append(' ');
                else if (builder.Length == 0 && (result.Count == 0 || !EndsWithWhitespace(result[result.Count - 1].Text))) {
                    result.Add(CopyRun(run, " "));
                }
                pendingSpace = false;
            }
            if (builder.Length > 0) result.Add(CopyRun(run, builder.ToString()));
        }
        TrimTrailingSpace(result);
        while (result.Count > 0 && result[result.Count - 1].Text == "\n") result.RemoveAt(result.Count - 1);
        return result.AsReadOnly();
    }

    private static void TrimTrailingSpace(IList<HtmlSemanticRun> runs) {
        if (runs.Count == 0) return;
        HtmlSemanticRun last = runs[runs.Count - 1];
        string text = last.Text.TrimEnd(' ', '\t', '\r');
        if (text.Length == 0) runs.RemoveAt(runs.Count - 1);
        else if (text.Length != last.Text.Length) runs[runs.Count - 1] = CopyRun(last, text);
    }

    private static bool EndsWithWhitespace(string text) =>
        text.Length > 0 && char.IsWhiteSpace(text[text.Length - 1]);

    private static IReadOnlyList<HtmlSemanticResource> BuildInlineResources(
        IElement owner,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<HtmlSemanticResource> resources,
        bool skipNestedLists = false) {
        var result = new List<HtmlSemanticResource>();
        foreach (IElement element in owner.QuerySelectorAll("img, video, audio, object, embed")) {
            if (skipNestedLists && HasAncestorBetween(element, owner, "ul", "ol", "dl")) continue;
            styles.TryGetValue(element, out HtmlComputedStyle? style);
            HtmlSemanticResource? resource = BuildResourceFromElement(element, style);
            if (resource == null) continue;
            result.Add(resource);
            resources.Add(resource);
        }
        return result.AsReadOnly();
    }

    private static bool HasAncestorBetween(IElement element, IElement owner, params string[] names) {
        for (IElement? current = element.ParentElement; current != null && !ReferenceEquals(current, owner); current = current.ParentElement) {
            if (names.Any(name => Is(current, name))) return true;
        }
        return false;
    }

    private static HtmlSemanticResource? BuildResource(
        IElement element,
        HtmlSemanticBlockKind kind,
        HtmlComputedStyle? style,
        HtmlSemanticSourceLocation? location) {
        string source;
        HtmlResourceKind resourceKind;
        if (kind == HtmlSemanticBlockKind.Image) {
            source = (element.GetAttribute("src") ?? element.GetAttribute("data") ?? string.Empty).Trim();
            resourceKind = HtmlResourceKind.Image;
        } else if (kind == HtmlSemanticBlockKind.Media) {
            source = (element.GetAttribute("src") ?? element.GetAttribute("data") ?? element.GetAttribute("poster") ?? string.Empty).Trim();
            resourceKind = HtmlResourceKind.Media;
        } else {
            return null;
        }
        if (source.Length == 0) return null;

        return BuildResourceFromElement(element, style, resourceKind, source, location);
    }

    private static HtmlSemanticResource? BuildResourceFromElement(IElement element, HtmlComputedStyle? style) {
        HtmlResourceKind kind = Is(element, "img") ? HtmlResourceKind.Image : HtmlResourceKind.Media;
        string source = (element.GetAttribute("src") ?? element.GetAttribute("data")
            ?? element.GetAttribute("poster") ?? string.Empty).Trim();
        if (source.Length == 0) return null;
        return BuildResourceFromElement(element, style, kind, source, HtmlSemanticSourceLocation.FromElement(element));
    }

    private static HtmlSemanticResource BuildResourceFromElement(
        IElement element,
        HtmlComputedStyle? style,
        HtmlResourceKind resourceKind,
        string source,
        HtmlSemanticSourceLocation? location) {
        string mediaType = (element.GetAttribute("type") ?? string.Empty).Trim();
        if (mediaType.Length == 0 && HtmlDataUri.TryParse(source, out HtmlDataUri dataUri)) mediaType = dataUri.MediaType;
        string alternateText = HtmlAccessibilitySemantics.GetAccessibleName(element, includeTextFallback: true);
        return new HtmlSemanticResource(resourceKind, source, alternateText, mediaType,
            ReadPixels(element.GetAttribute("width") ?? style?.GetValue("width")),
            ReadPixels(element.GetAttribute("height") ?? style?.GetValue("height")),
            location);
    }

    private static double? ReadPixels(string? value) {
        string text = (value ?? string.Empty).Trim();
        if (text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) text = text.Substring(0, text.Length - 2).Trim();
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double pixels) && pixels > 0D
            ? pixels
            : null;
    }

    private static HtmlSemanticFormControl? BuildFormControl(IElement element) {
        if (!Is(element, "input") && !Is(element, "select") && !Is(element, "textarea")
            && !Is(element, "button") && !Is(element, "option")) return null;
        string type = Normalize(element.GetAttribute("type")).ToLowerInvariant();
        if (type.Length == 0) type = element.LocalName.ToLowerInvariant();
        string value = element.GetAttribute("value") ?? Normalize(element.TextContent);
        return new HtmlSemanticFormControl(
            type,
            element.GetAttribute("name") ?? string.Empty,
            value,
            element.HasAttribute("checked") || element.HasAttribute("selected"),
            element.HasAttribute("disabled"));
    }

    private static IReadOnlyDictionary<string, string> ReadMetadata(IHtmlDocument document) {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (IElement meta in document.QuerySelectorAll("meta[name][content], meta[property][content]")) {
            string name = Normalize(meta.GetAttribute("name") ?? meta.GetAttribute("property")).ToLowerInvariant();
            string value = Normalize(meta.GetAttribute("content"));
            if (name.Length > 0 && value.Length > 0 && !result.ContainsKey(name)) result[name] = value;
        }
        if (!string.IsNullOrWhiteSpace(document.Title)) result["title"] = Normalize(document.Title);
        return new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(result);
    }

    private static IReadOnlyList<HtmlSemanticResource> DeduplicateResources(IEnumerable<HtmlSemanticResource> resources) =>
        resources.GroupBy(resource => resource.Kind + "\0" + resource.Source, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First()).ToList().AsReadOnly();

    private static HtmlSemanticBlockKind GetKind(IElement element) {
        if (HtmlGenericDocumentProjector.IsHeading(element)) return HtmlSemanticBlockKind.Heading;
        if (HtmlGenericDocumentProjector.IsTable(element)) return HtmlSemanticBlockKind.Table;
        if (HtmlGenericDocumentProjector.IsImage(element)) return HtmlSemanticBlockKind.Image;
        if (Is(element, "ul") || Is(element, "ol") || Is(element, "dl")) return HtmlSemanticBlockKind.List;
        if (Is(element, "pre")) return HtmlSemanticBlockKind.Code;
        if (Is(element, "blockquote")) return HtmlSemanticBlockKind.Quote;
        if (Is(element, "video") || Is(element, "audio") || Is(element, "object") || Is(element, "embed")) return HtmlSemanticBlockKind.Media;
        if (Is(element, "form") || Is(element, "input") || Is(element, "select") || Is(element, "textarea") || Is(element, "button")) return HtmlSemanticBlockKind.Form;
        if (HtmlAccessibilitySemantics.HasRole(element, "doc-footnote") || HtmlAccessibilitySemantics.HasRole(element, "doc-endnote") || HtmlAccessibilitySemantics.HasRole(element, "note")) return HtmlSemanticBlockKind.Note;
        if (Is(element, "p") || Is(element, "address")) return HtmlSemanticBlockKind.Paragraph;
        return HtmlSemanticBlockKind.Other;
    }

    private static bool IsTextual(HtmlSemanticBlockKind kind) =>
        kind == HtmlSemanticBlockKind.Heading || kind == HtmlSemanticBlockKind.Paragraph
        || kind == HtmlSemanticBlockKind.Code || kind == HtmlSemanticBlockKind.Quote
        || kind == HtmlSemanticBlockKind.Note || kind == HtmlSemanticBlockKind.Form;

    private static int GetHeadingLevel(IElement element) {
        if (HtmlAccessibilitySemantics.TryGetHeadingLevel(element, out int level)) return level;
        return element.LocalName.Length == 2 && char.IsDigit(element.LocalName[1])
            ? element.LocalName[1] - '0'
            : 0;
    }

    private static int GetListDepth(IElement element) {
        int depth = 1;
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (Is(parent, "ul") || Is(parent, "ol") || Is(parent, "dl")) depth++;
        }
        return depth;
    }

    private static int ReadSpan(IElement element, string attribute) =>
        int.TryParse(element.GetAttribute(attribute), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value > 0
            ? value
            : 1;

    private static IEnumerable<IElement> DirectRows(IElement table) {
        foreach (IElement child in table.Children) {
            if (Is(child, "tr")) yield return child;
            else if (Is(child, "thead") || Is(child, "tbody") || Is(child, "tfoot")) {
                foreach (IElement row in child.Children.Where(candidate => Is(candidate, "tr"))) yield return row;
            }
        }
    }

    private static bool IsTableCell(IElement element) => Is(element, "th") || Is(element, "td");

    private static bool Is(IElement element, string localName) =>
        string.Equals(element.LocalName, localName, StringComparison.OrdinalIgnoreCase);

    private static string Normalize(string? value) =>
        string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private readonly struct InlineState {
        internal bool Bold { get; }
        internal bool Italic { get; }
        internal bool Underline { get; }
        internal bool Strikethrough { get; }
        internal bool Superscript { get; }
        internal bool Subscript { get; }
        internal string? Hyperlink { get; }
        internal HtmlComputedStyle? Style { get; }
        internal HtmlSemanticSourceLocation? Location { get; }

        private InlineState(bool bold, bool italic, bool underline, bool strikethrough,
            bool superscript, bool subscript, string? hyperlink, HtmlComputedStyle? style,
            HtmlSemanticSourceLocation? location) {
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strikethrough = strikethrough;
            Superscript = superscript;
            Subscript = subscript;
            Hyperlink = hyperlink;
            Style = style;
            Location = location;
        }

        internal InlineState With(IElement element, HtmlComputedStyle? style, HtmlSemanticSourceLocation? location) {
            string name = element.LocalName.ToLowerInvariant();
            string fontWeight = style?.GetValue("font-weight") ?? string.Empty;
            bool cssBold = string.Equals(fontWeight, "bold", StringComparison.OrdinalIgnoreCase)
                || (int.TryParse(fontWeight, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight) && weight >= 600);
            string decoration = style?.GetValue("text-decoration-line") ?? style?.GetValue("text-decoration") ?? string.Empty;
            string vertical = style?.GetValue("vertical-align") ?? string.Empty;
            return new InlineState(
                Bold || name == "strong" || name == "b" || cssBold,
                Italic || name == "em" || name == "i" || (style?.GetValue("font-style") ?? string.Empty).IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0,
                Underline || name == "u" || decoration.IndexOf("underline", StringComparison.OrdinalIgnoreCase) >= 0,
                Strikethrough || name == "s" || name == "strike" || name == "del" || decoration.IndexOf("line-through", StringComparison.OrdinalIgnoreCase) >= 0,
                Superscript || name == "sup" || string.Equals(vertical, "super", StringComparison.OrdinalIgnoreCase),
                Subscript || name == "sub" || string.Equals(vertical, "sub", StringComparison.OrdinalIgnoreCase),
                name == "a" ? element.GetAttribute("href") : Hyperlink,
                style ?? Style,
                location ?? Location);
        }
    }
}
