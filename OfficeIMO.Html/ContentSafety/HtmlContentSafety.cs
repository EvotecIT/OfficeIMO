using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Inspects machine-readable HTML that is concealed from an ordinary rendered view.</summary>
public static class HtmlContentSafety {
    /// <summary>Inspects HTML using OfficeIMO's computed CSS cascade.</summary>
    public static OfficeContentSafetyReport Inspect(string html, OfficeContentSafetyOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateText(html, effective);
        HtmlConversionDocument conversion = HtmlConversionDocument.Parse(html);
        IHtmlDocument document = conversion.CreateSourceDocumentForConversion();
        return InspectDocument(document, effective, targets: null);
    }

    /// <summary>Inspects a UTF-8 HTML file.</summary>
    public static OfficeContentSafetyReport InspectFile(string filePath, OfficeContentSafetyOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return Inspect(OfficeContentSafetyInputGuard.ReadUtf8Text(filePath, effective), effective);
    }

    /// <summary>Removes only exact, current findings selected by the caller and then reinspects the result.</summary>
    public static OfficeContentCleanupResult RemoveSelected(
        string html,
        OfficeContentCleanupSelection selection,
        OfficeContentSafetyOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateText(html, effective);
        HtmlConversionDocument conversion = HtmlConversionDocument.Parse(html);
        IHtmlDocument document = conversion.CreateSourceDocumentForConversion();
        var targets = new Dictionary<string, HtmlCleanupTarget>(StringComparer.Ordinal);
        OfficeContentSafetyReport before = InspectDocument(document, effective, targets);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) return new OfficeContentCleanupResult(Encoding.UTF8.GetBytes(html), before, before, Array.Empty<OfficeContentCleanupChange>());
        foreach (IGrouping<HtmlCleanupTarget, OfficeContentSafetyFinding> group in selected
            .OrderByDescending(item => item.SourceTextOffset ?? -1)
            .GroupBy(item => targets[item.Id])) {
            group.Key.Remove();
        }
        string outputHtml = document.ToHtml();
        byte[] output = Encoding.UTF8.GetBytes(outputHtml);
        OfficeContentSafetyReport after = Inspect(outputHtml, effective);
        OfficeContentCleanupChange[] changes = selected.Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability)).ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    /// <summary>Atomically writes an explicitly cleaned HTML artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedFile(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentSafetyOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        byte[] input = OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, effective);
        string html = OfficeContentSafetyInputGuard.DecodeText(input, effective);
        OfficeContentCleanupResult result;
        if (selection.FindingIds.Count == 0) {
            OfficeContentSafetyReport report = Inspect(html, effective);
            result = new OfficeContentCleanupResult((byte[])input.Clone(), report, report, Array.Empty<OfficeContentCleanupChange>());
        } else {
            result = RemoveSelected(html, selection, effective);
        }
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static OfficeContentSafetyReport InspectDocument(
        IHtmlDocument document,
        OfficeContentSafetyOptions? options,
        IDictionary<string, HtmlCleanupTarget>? targets) {
        var builder = new OfficeContentSafetyBuilder("HTML", options);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);
        IElement? root = document.DocumentElement ?? document.Body;
        if (root != null) Traverse(root, styles, builder, targets, ancestorConcealed: false);
        InspectComments(document, builder, targets);
        return builder.Build();
    }

    private static void Traverse(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, HtmlCleanupTarget>? targets,
        bool ancestorConcealed) {
        string location = BuildLocation(element);
        styles.TryGetValue(element, out HtmlComputedStyle? style);
        InspectMachineOnlyAttributes(element, location, builder, targets);

        Concealment? concealment = ancestorConcealed ? null : FindElementConcealment(element, style, styles, builder.Options);
        if (concealment != null) {
            string text = element.TextContent ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text)) {
                OfficeContentCleanupCapability capability = CanRemoveElement(element)
                    ? OfficeContentCleanupCapability.RemoveElement
                    : OfficeContentCleanupCapability.RemoveText;
                OfficeContentSafetyFinding finding = builder.Add(
                    concealment.Kind,
                    concealment.Risk,
                    location,
                    concealment.Evidence,
                    text,
                    capability,
                    inspectTextIntegrityEvidence: false);
                if (targets != null) targets[finding.Id] = capability == OfficeContentCleanupCapability.RemoveElement
                    ? HtmlCleanupTarget.ForElement(element)
                    : HtmlCleanupTarget.ForText(element);
                InspectHtmlTextIntegrity(element, location, builder, targets, alreadyCharged: true);
            }
            return;
        }

        if (IsNonTextElement(element)) {
            string machineText = element.TextContent ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(machineText) && builder.Options.IncludeNonPrimaryContent) {
                bool safeToRemove = !string.Equals(element.LocalName, "style", StringComparison.OrdinalIgnoreCase) && CanRemoveElement(element);
                OfficeContentCleanupCapability capability = safeToRemove
                    ? OfficeContentCleanupCapability.RemoveElement
                    : OfficeContentCleanupCapability.ReportOnly;
                OfficeContentSafetyFinding finding = builder.Add(
                    OfficeContentConcealmentKind.NonPrimaryContent,
                    OfficeContentSafetyRisk.ContextDependent,
                    location,
                    "The " + element.LocalName + " payload is machine-readable source content but is not ordinary rendered body text.",
                    machineText,
                    capability,
                    inspectTextIntegrityEvidence: false);
                if (targets != null && safeToRemove) targets[finding.Id] = HtmlCleanupTarget.ForElement(element);
                InspectHtmlTextIntegrity(element, location, builder, targets, alreadyCharged: true);
            }
            return;
        }

        if (!ancestorConcealed) {
            string directText = string.Concat(element.ChildNodes.Where(node => node.NodeType == NodeType.Text).Select(node => node.TextContent));
            if (!string.IsNullOrWhiteSpace(directText)) {
                Concealment? lowContrast = FindLowContrast(element, style, styles, builder.Options);
                if (lowContrast != null) {
                    OfficeContentSafetyFinding finding = builder.Add(
                        lowContrast.Kind,
                        lowContrast.Risk,
                        location + "/text()",
                        lowContrast.Evidence,
                        directText,
                        OfficeContentCleanupCapability.RemoveText,
                        inspectTextIntegrityEvidence: false);
                    if (targets != null) targets[finding.Id] = HtmlCleanupTarget.ForDirectText(element);
                    InspectHtmlDirectTextIntegrity(element, location, builder, targets, alreadyCharged: true);
                } else {
                    InspectHtmlDirectTextIntegrity(element, location, builder, targets, alreadyCharged: false);
                }
            }
        }

        if (string.Equals(element.LocalName, "template", StringComparison.OrdinalIgnoreCase)) {
            string templateText = NormalizePayload(element.TextContent);
            if (templateText.Length > 0 && builder.Options.IncludeNonPrimaryContent) {
                AddAttributeOrNonPrimaryFinding(builder, targets, element, null, location, "HTML template content is not part of the ordinary rendered document.", templateText);
            }
            return;
        }

        foreach (IElement child in element.Children) Traverse(child, styles, builder, targets, ancestorConcealed: false);
    }

    private static void InspectHtmlTextIntegrity(
        IElement element,
        string location,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, HtmlCleanupTarget>? targets,
        bool alreadyCharged) {
        var textNodes = new List<IText>();
        CollectHtmlTextNodes(element, textNodes);
        InspectHtmlTextNodes(textNodes, location, builder, targets, alreadyCharged);
    }

    private static void InspectHtmlDirectTextIntegrity(
        IElement element,
        string location,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, HtmlCleanupTarget>? targets,
        bool alreadyCharged) => InspectHtmlTextNodes(element.ChildNodes.OfType<IText>(), location, builder, targets, alreadyCharged);

    private static void InspectHtmlTextNodes(
        IEnumerable<IText> textNodes,
        string location,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, HtmlCleanupTarget>? targets,
        bool alreadyCharged) {
        int textIndex = 0;
        foreach (IText textNode in textNodes) {
            string nodeText = textNode.Data ?? string.Empty;
            if (nodeText.Length == 0) continue;
            string nodeLocation = location + "/text()[" + (++textIndex).ToString(CultureInfo.InvariantCulture) + "]";
            IReadOnlyList<OfficeContentSafetyFinding> unicode = alreadyCharged
                ? builder.InspectChargedTextIntegrity(nodeLocation, nodeText, OfficeContentCleanupCapability.RemoveText)
                : builder.InspectVisibleText(nodeLocation, nodeText, OfficeContentCleanupCapability.RemoveText);
            if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = HtmlCleanupTarget.ForTextRange(textNode, item);
        }
    }

    private static void CollectHtmlTextNodes(INode node, ICollection<IText> textNodes) {
        foreach (INode child in node.ChildNodes) {
            if (child is IText text) textNodes.Add(text);
            else CollectHtmlTextNodes(child, textNodes);
        }
    }

    private static Concealment? FindElementConcealment(
        IElement element,
        HtmlComputedStyle? style,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        OfficeContentSafetyOptions options) {
        if (element.HasAttribute("hidden") || string.Equals(element.LocalName, "input", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(element.GetAttribute("type"), "hidden", StringComparison.OrdinalIgnoreCase)) {
            return new Concealment(OfficeContentConcealmentKind.HiddenByProperty, "The HTML hidden state prevents ordinary rendering.");
        }
        if (style == null) return null;
        string display = style.GetValue("display").Trim();
        if (string.Equals(display, "none", StringComparison.OrdinalIgnoreCase)) {
            return new Concealment(OfficeContentConcealmentKind.HiddenByProperty, "Computed CSS display is none.");
        }
        string visibility = style.GetValue("visibility").Trim();
        if (string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase) || string.Equals(visibility, "collapse", StringComparison.OrdinalIgnoreCase)) {
            return new Concealment(OfficeContentConcealmentKind.HiddenByProperty, "Computed CSS visibility is " + visibility + ".");
        }
        if (TryParseScalar(style.GetValue("opacity"), out double opacity) && opacity <= 0.01D) {
            return new Concealment(OfficeContentConcealmentKind.TransparentText, "Computed CSS opacity is " + opacity.ToString("0.###", CultureInfo.InvariantCulture) + ".");
        }
        if (TryParseCssColor(style.GetValue("color"), out OfficeColor textColor) && textColor.A <= 3) {
            return new Concealment(OfficeContentConcealmentKind.TransparentText, "Computed CSS text color is fully or nearly transparent.");
        }
        string filter = style.GetValue("filter");
        if (TryGetCssFilterOpacity(filter, out double filterOpacity) && filterOpacity <= 0.01D) {
            return new Concealment(OfficeContentConcealmentKind.TransparentText, "Computed CSS filter applies zero opacity.");
        }
        if (TryParseLengthPoints(style.GetValue("font-size"), out double fontPoints) && fontPoints <= options.MaximumTinyFontSizePoints) {
            return new Concealment(OfficeContentConcealmentKind.TinyText, "Computed font size is " + fontPoints.ToString("0.###", CultureInfo.InvariantCulture) + "pt.");
        }
        bool zeroWidth = IsZeroLength(style.GetValue("width")) || IsZeroLength(style.GetValue("max-width"));
        bool zeroHeight = IsZeroLength(style.GetValue("height")) || IsZeroLength(style.GetValue("max-height"));
        string overflow = style.GetValue("overflow") + " " + style.GetValue("overflow-x") + " " + style.GetValue("overflow-y");
        if ((zeroWidth || zeroHeight) && (overflow.IndexOf("hidden", StringComparison.OrdinalIgnoreCase) >= 0 || overflow.IndexOf("clip", StringComparison.OrdinalIgnoreCase) >= 0)) {
            return new Concealment(OfficeContentConcealmentKind.ZeroDimension, "Computed zero-size geometry is combined with clipped overflow.");
        }
        string clipPath = style.GetValue("clip-path");
        string clip = style.GetValue("clip");
        if (IsZeroClip(clipPath) || IsZeroClip(clip)) {
            return new Concealment(OfficeContentConcealmentKind.ClippedContent, "Computed CSS clip removes the content from the visible area.");
        }
        string position = style.GetValue("position");
        if (string.Equals(position, "absolute", StringComparison.OrdinalIgnoreCase) || string.Equals(position, "fixed", StringComparison.OrdinalIgnoreCase)) {
            if (IsFarNegative(style.GetValue("left")) || IsFarNegative(style.GetValue("top")) ||
                IsFarNegativeTextIndent(style.GetValue("text-indent")) || IsFarTranslation(style.GetValue("transform"))) {
                return new Concealment(OfficeContentConcealmentKind.OffCanvas, "Computed positioned geometry moves the content far outside the ordinary viewport.");
            }
        }
        return null;
    }

    private static Concealment? FindLowContrast(
        IElement element,
        HtmlComputedStyle? style,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        OfficeContentSafetyOptions options) {
        if (style == null || !TryParseCssColor(style.GetValue("color"), out OfficeColor foreground)) return null;
        OfficeColor background = OfficeColor.White;
        for (IElement? current = element; current != null; current = current.ParentElement) {
            if (!styles.TryGetValue(current, out HtmlComputedStyle? currentStyle)) continue;
            if (HasNonSolidBackground(currentStyle)) return null;
            if (!TryParseBackgroundColor(currentStyle, out OfficeColor parsed)) continue;
            if (parsed.A < byte.MaxValue) return null;
            background = parsed;
            break;
        }
        foreground = CompositeOver(foreground, background);
        double ratio = OfficeColorContrast.ContrastRatio(foreground, background);
        if (ratio + 0.000001D >= options.MinimumVisibleContrastRatio) return null;
        return new Concealment(
            OfficeContentConcealmentKind.LowContrastText,
            "Computed foreground #" + foreground.ToRgbHex() + " against background #" + background.ToRgbHex() +
            " has contrast ratio " + ratio.ToString("0.###", CultureInfo.InvariantCulture) + ".");
    }

    private static void InspectMachineOnlyAttributes(
        IElement element,
        string location,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, HtmlCleanupTarget>? targets) {
        if (!builder.Options.IncludeNonPrimaryContent) return;
        foreach (string attribute in new[] { "alt", "aria-label", "title", "data-ai", "data-prompt" }) {
            string value = element.GetAttribute(attribute) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(value)) continue;
            AddAttributeOrNonPrimaryFinding(builder, targets, element, attribute, location + "/@" + attribute,
                "The " + attribute + " attribute is machine-readable but not ordinary body text.", value);
        }
        if (string.Equals(element.LocalName, "meta", StringComparison.OrdinalIgnoreCase)) {
            string value = element.GetAttribute("content") ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(value)) {
                AddAttributeOrNonPrimaryFinding(builder, targets, element, "content", location + "/@content",
                    "HTML metadata is machine-readable but not rendered as ordinary body text.", value);
            }
        }
        if (string.Equals(element.LocalName, "input", StringComparison.OrdinalIgnoreCase)) {
            string value = element.GetAttribute("value") ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(value) && string.Equals(element.GetAttribute("type"), "hidden", StringComparison.OrdinalIgnoreCase)) {
                AddAttributeOrNonPrimaryFinding(builder, targets, element, "value", location + "/@value",
                    "A hidden form control retains a machine-readable value.", value);
            }
        }
    }

    private static void AddAttributeOrNonPrimaryFinding(
        OfficeContentSafetyBuilder builder,
        IDictionary<string, HtmlCleanupTarget>? targets,
        IElement element,
        string? attribute,
        string location,
        string evidence,
        string value) {
        OfficeContentSafetyFinding finding = builder.Add(
            OfficeContentConcealmentKind.NonPrimaryContent,
            OfficeContentSafetyRisk.ContextDependent,
            location,
            evidence,
            NormalizePayload(value),
            attribute == null ? OfficeContentCleanupCapability.RemoveElement : OfficeContentCleanupCapability.RemoveText);
        if (targets != null) targets[finding.Id] = attribute == null
            ? HtmlCleanupTarget.ForElement(element)
            : HtmlCleanupTarget.ForAttribute(element, attribute);
    }

    private static void InspectComments(
        IHtmlDocument document,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, HtmlCleanupTarget>? targets) {
        if (!builder.Options.IncludeNonPrimaryContent) return;
        IComment[] comments = document.Descendants<IComment>().ToArray();
        for (int index = 0; index < comments.Length; index++) {
            string value = NormalizePayload(comments[index].Data);
            if (value.Length == 0) continue;
            OfficeContentSafetyFinding finding = builder.Add(
                OfficeContentConcealmentKind.NonPrimaryContent,
                OfficeContentSafetyRisk.ContextDependent,
                "HTML/comment()[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]",
                "An HTML comment is machine-readable source content but is not rendered.",
                value,
                OfficeContentCleanupCapability.RemoveElement);
            if (targets != null) targets[finding.Id] = HtmlCleanupTarget.ForNode(comments[index]);
        }
    }

    private static string BuildLocation(IElement element) {
        var segments = new Stack<string>();
        for (IElement? current = element; current != null; current = current.ParentElement) {
            int index = 1;
            for (IElement? sibling = current.PreviousElementSibling; sibling != null; sibling = sibling.PreviousElementSibling) {
                if (string.Equals(sibling.LocalName, current.LocalName, StringComparison.OrdinalIgnoreCase)) index++;
            }
            segments.Push(current.LocalName + "[" + index.ToString(CultureInfo.InvariantCulture) + "]");
        }
        return "HTML/" + string.Join("/", segments);
    }

    private static bool TryParseBackgroundColor(HtmlComputedStyle style, out OfficeColor color) {
        if (TryParseCssColor(style.GetValue("background-color"), out color) && color.A > 0) return true;
        string background = style.GetValue("background");
        foreach (string token in background.Split(new[] { ' ', '\t', '\r', '\n', '/', ',' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (TryParseCssColor(token, out color) && color.A > 0) return true;
        }
        color = default;
        return false;
    }

    private static bool HasNonSolidBackground(HtmlComputedStyle style) {
        string image = style.GetValue("background-image").Trim();
        string shorthand = style.GetValue("background");
        return ContainsCssImage(image) || ContainsCssImage(shorthand);
    }

    private static bool ContainsCssImage(string value) =>
        value.IndexOf("url(", StringComparison.OrdinalIgnoreCase) >= 0 ||
        value.IndexOf("gradient(", StringComparison.OrdinalIgnoreCase) >= 0 ||
        value.IndexOf("image(", StringComparison.OrdinalIgnoreCase) >= 0 ||
        value.IndexOf("image-set(", StringComparison.OrdinalIgnoreCase) >= 0 ||
        value.IndexOf("cross-fade(", StringComparison.OrdinalIgnoreCase) >= 0 ||
        value.IndexOf("element(", StringComparison.OrdinalIgnoreCase) >= 0 ||
        value.IndexOf("paint(", StringComparison.OrdinalIgnoreCase) >= 0;

    private static OfficeColor CompositeOver(OfficeColor foreground, OfficeColor background) {
        if (foreground.A == byte.MaxValue) return foreground;
        double alpha = foreground.A / 255D;
        return OfficeColor.FromRgb(
            (byte)Math.Round(foreground.R * alpha + background.R * (1D - alpha), MidpointRounding.AwayFromZero),
            (byte)Math.Round(foreground.G * alpha + background.G * (1D - alpha), MidpointRounding.AwayFromZero),
            (byte)Math.Round(foreground.B * alpha + background.B * (1D - alpha), MidpointRounding.AwayFromZero));
    }

    private static bool TryParseCssColor(string value, out OfficeColor color) {
        string normalized = value?.Trim() ?? string.Empty;
        if (normalized.Length == 0 || string.Equals(normalized, "transparent", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "currentcolor", StringComparison.OrdinalIgnoreCase)) {
            color = default;
            return false;
        }
        return OfficeColor.TryParseCss(normalized, out color);
    }

    private static bool TryParseLengthPoints(string value, out double points) {
        string normalized = value?.Trim().ToLowerInvariant() ?? string.Empty;
        if (normalized == "0") { points = 0; return true; }
        double multiplier;
        int suffix;
        if (normalized.EndsWith("pt", StringComparison.Ordinal)) { multiplier = 1D; suffix = 2; }
        else if (normalized.EndsWith("px", StringComparison.Ordinal)) { multiplier = 0.75D; suffix = 2; }
        else { points = 0; return false; }
        if (!double.TryParse(normalized.Substring(0, normalized.Length - suffix), NumberStyles.Float, CultureInfo.InvariantCulture, out double valueNumber)) {
            points = 0;
            return false;
        }
        points = valueNumber * multiplier;
        return true;
    }

    private static bool TryParseScalar(string value, out double result) =>
        double.TryParse(value?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out result);

    private static bool TryGetCssFilterOpacity(string value, out double opacity) {
        opacity = 1D;
        string source = value ?? string.Empty;
        int cursor = 0;
        bool found = false;
        while (cursor < source.Length) {
            if (!SkipCssWhitespaceAndComments(source, ref cursor)) return false;
            if (cursor >= source.Length) break;
            int nameStart = cursor;
            while (cursor < source.Length && (char.IsLetter(source[cursor]) || source[cursor] == '-')) cursor++;
            if (cursor == nameStart) return false;
            string functionName = source.Substring(nameStart, cursor - nameStart);
            while (cursor < source.Length && char.IsWhiteSpace(source[cursor])) cursor++;
            if (cursor >= source.Length || source[cursor] != '(' || !TryFindCssFunctionClose(source, cursor, out int close)) return false;
            if (string.Equals(functionName, "opacity", StringComparison.OrdinalIgnoreCase)) {
                string lexical = source.Substring(cursor + 1, close - cursor - 1).Trim();
                bool percent = lexical.EndsWith("%", StringComparison.Ordinal);
                if (percent) lexical = lexical.Substring(0, lexical.Length - 1).Trim();
                if (!double.TryParse(lexical, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) || double.IsNaN(parsed) || double.IsInfinity(parsed)) return false;
                double component = percent ? parsed / 100D : parsed;
                if (component < 0D) return false;
                opacity *= Math.Min(1D, component);
                found = true;
            }
            cursor = close + 1;
        }
        return found;
    }

    private static bool SkipCssWhitespaceAndComments(string source, ref int cursor) {
        while (cursor < source.Length) {
            if (char.IsWhiteSpace(source[cursor])) { cursor++; continue; }
            if (cursor + 1 >= source.Length || source[cursor] != '/' || source[cursor + 1] != '*') return true;
            int close = source.IndexOf("*/", cursor + 2, StringComparison.Ordinal);
            if (close < 0) return false;
            cursor = close + 2;
        }
        return true;
    }

    private static bool TryFindCssFunctionClose(string source, int open, out int close) {
        int depth = 1;
        char quote = '\0';
        bool escaped = false;
        for (int index = open + 1; index < source.Length; index++) {
            char current = source[index];
            if (escaped) { escaped = false; continue; }
            if (current == '\\') { escaped = true; continue; }
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') { quote = current; continue; }
            if (current == '(') depth++;
            else if (current == ')' && --depth == 0) { close = index; return true; }
        }
        close = -1;
        return false;
    }

    private static bool IsZeroLength(string value) => TryParseLengthPoints(value, out double points) && Math.Abs(points) <= 0.000001D;

    private static bool IsZeroClip(string value) {
        string normalized = (value ?? string.Empty).Replace(" ", string.Empty).ToLowerInvariant();
        return normalized.Contains("rect(0px,0px,0px,0px)") || normalized.Contains("rect(0,0,0,0)") ||
               normalized.Contains("inset(50%)") || normalized == "circle(0)" || normalized == "circle(0px)";
    }

    private static bool IsFarNegative(string value) => TryParseLengthPoints(value, out double points) && points <= -1000D;
    private static bool IsFarNegativeTextIndent(string value) => IsFarNegative(value);
    private static bool IsFarTranslation(string value) {
        string normalized = value?.ToLowerInvariant() ?? string.Empty;
        int marker = normalized.IndexOf("translate", StringComparison.Ordinal);
        if (marker < 0) return false;
        return normalized.IndexOf("-999", marker, StringComparison.Ordinal) >= 0 || normalized.IndexOf("-1000", marker, StringComparison.Ordinal) >= 0;
    }

    private static bool IsNonTextElement(IElement element) => element.LocalName.ToLowerInvariant() is "script" or "style" or "noscript";
    private static bool CanRemoveElement(IElement element) => element.ParentElement != null && element.LocalName.ToLowerInvariant() is not "html" and not "body";
    private static string NormalizePayload(string value) => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private sealed class Concealment {
        internal Concealment(OfficeContentConcealmentKind kind, string evidence, OfficeContentSafetyRisk risk = OfficeContentSafetyRisk.ContextDependent) {
            Kind = kind; Evidence = evidence; Risk = risk;
        }
        internal OfficeContentConcealmentKind Kind { get; }
        internal string Evidence { get; }
        internal OfficeContentSafetyRisk Risk { get; }
    }

    private sealed class HtmlCleanupTarget : IEquatable<HtmlCleanupTarget> {
        private readonly INode _node;
        private readonly string? _attribute;
        private readonly TextRemovalMode _textRemoval;
        private readonly int? _offset;
        private readonly int? _length;
        private readonly string? _expected;
        private HtmlCleanupTarget(INode node, string? attribute, TextRemovalMode textRemoval, int? offset = null, int? length = null, string? expected = null) { _node = node; _attribute = attribute; _textRemoval = textRemoval; _offset = offset; _length = length; _expected = expected; }
        internal static HtmlCleanupTarget ForElement(IElement element) => new HtmlCleanupTarget(element, null, TextRemovalMode.None);
        internal static HtmlCleanupTarget ForText(IElement element) => new HtmlCleanupTarget(element, null, TextRemovalMode.Descendants);
        internal static HtmlCleanupTarget ForDirectText(IElement element) => new HtmlCleanupTarget(element, null, TextRemovalMode.DirectChildren);
        internal static HtmlCleanupTarget ForAttribute(IElement element, string attribute) => new HtmlCleanupTarget(element, attribute, TextRemovalMode.None);
        internal static HtmlCleanupTarget ForNode(INode node) => new HtmlCleanupTarget(node, null, TextRemovalMode.None);
        internal static HtmlCleanupTarget ForTextRange(IText node, OfficeContentSafetyFinding finding) => new HtmlCleanupTarget(
            node, null, TextRemovalMode.Range, finding.SourceTextOffset, finding.SourceTextLength,
            node.Data.Substring(finding.SourceTextOffset!.Value, finding.SourceTextLength!.Value));
        internal void Remove() {
            if (_textRemoval == TextRemovalMode.Range && _node is IText text && _offset.HasValue && _length.HasValue) {
                string current = text.Data ?? string.Empty;
                if (_offset.Value > current.Length - _length.Value || !string.Equals(current.Substring(_offset.Value, _length.Value), _expected, StringComparison.Ordinal)) {
                    throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected HTML text node.");
                }
                text.Replace(_offset.Value, _length.Value, string.Empty);
                return;
            }
            if (_attribute != null && _node is IElement element) { element.RemoveAttribute(_attribute); return; }
            if (_textRemoval != TextRemovalMode.None && _node is IElement textOwner) {
                IEnumerable<INode> textNodes = _textRemoval == TextRemovalMode.Descendants
                    ? textOwner.Descendants().Where(node => node.NodeType == NodeType.Text)
                    : textOwner.ChildNodes.Where(node => node.NodeType == NodeType.Text);
                foreach (INode child in textNodes.ToArray()) child.Parent?.RemoveChild(child);
                return;
            }
            _node.Parent?.RemoveChild(_node);
        }
        public bool Equals(HtmlCleanupTarget? other) => other != null && ReferenceEquals(_node, other._node) && string.Equals(_attribute, other._attribute, StringComparison.Ordinal) && _textRemoval == other._textRemoval && _offset == other._offset && _length == other._length;
        public override bool Equals(object? obj) => Equals(obj as HtmlCleanupTarget);
        public override int GetHashCode() { unchecked { return (_node.GetHashCode() * 397) ^ (_attribute?.GetHashCode() ?? 0) ^ (int)_textRemoval ^ (_offset ?? 0); } }
        private enum TextRemovalMode { None, DirectChildren, Descendants, Range }
    }
}
