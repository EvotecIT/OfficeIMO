using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private const int LongDocumentParagraphCount = 20;

    /// <summary>Inspects the document body against an accessibility policy: alternative text, headings, tables, links, contrast, floating objects, spacing, title and language.</summary>
    public OfficeAccessibilityReport InspectAccessibility(WordAccessibilityOptions? options = null) {
        options ??= new WordAccessibilityOptions();
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart
            ?? throw new InvalidOperationException("MainDocumentPart is missing.");
        W.Body? body = main.Document?.Body;
        var findings = new List<OfficeAccessibilityFinding>();
        var context = new AccessibilityContext(main, options);

        if (options.RequireDocumentTitle && string.IsNullOrWhiteSpace(_wordprocessingDocument.PackageProperties.Title)) {
            findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle,
                "The document does not define a title.", "Document"));
        }
        if (options.RequireLanguage && GetDefaultLanguage(main) == null) {
            findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityMissingLanguage,
                "The document does not define a default text language.", "Document"));
        }
        if (body == null) return new OfficeAccessibilityReport("Word", options.Profile, findings);

        List<W.Paragraph> paragraphs = body.Descendants<W.Paragraph>().ToList();
        var paragraphIndexes = new Dictionary<W.Paragraph, int>(paragraphs.Count);
        for (int index = 0; index < paragraphs.Count; index++) paragraphIndexes[paragraphs[index]] = index;
        context.ParagraphIndexes = paragraphIndexes;

        if (options.CheckHeadings) InspectHeadings(paragraphs, context, findings);
        int tableIndex = 0;
        foreach (W.Table table in body.Descendants<W.Table>().ToList()) InspectTable(table, ++tableIndex, context, findings);
        for (int index = 0; index < paragraphs.Count; index++) InspectParagraph(paragraphs[index], index, context, findings);
        if (options.CheckBlankParagraphs) InspectBlankParagraphs(body, context, findings);

        return new OfficeAccessibilityReport("Word", options.Profile, findings);
    }

    /// <summary>Applies the selected safe fixes (title, language, table header rows) and returns a fresh inspection with default options.</summary>
    public OfficeAccessibilityReport FixAccessibility(WordAccessibilityFixOptions options) {
        ArgumentNullException.ThrowIfNull(options);
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart
            ?? throw new InvalidOperationException("MainDocumentPart is missing.");
        if (!string.IsNullOrWhiteSpace(options.Title) && string.IsNullOrWhiteSpace(_wordprocessingDocument.PackageProperties.Title)) {
            _wordprocessingDocument.PackageProperties.Title = options.Title!.Trim();
        }
        if (!string.IsNullOrWhiteSpace(options.Language) && GetDefaultLanguage(main) == null) {
            SetDefaultLanguage(main, options.Language!.Trim());
        }
        if (options.MarkTableHeaderRows && main.Document?.Body != null) {
            foreach (W.Table table in main.Document.Body.Descendants<W.Table>().ToList()) {
                List<W.TableRow> rows = table.Elements<W.TableRow>().ToList();
                if (rows.Count < 2 || HasHeaderRow(rows) || IsLikelyLayoutTable(table, rows)) continue;
                W.TableRowProperties properties = rows[0].TableRowProperties ?? rows[0].PrependChild(new W.TableRowProperties());
                properties.AppendChild(new W.TableHeader());
            }
        }
        return InspectAccessibility();
    }

    private sealed class AccessibilityContext {
        internal AccessibilityContext(MainDocumentPart main, WordAccessibilityOptions options) {
            Options = options;
            Styles = main.StyleDefinitionsPart?.Styles?.Elements<W.Style>()
                .Where(style => style.StyleId?.Value != null)
                .GroupBy(style => style.StyleId!.Value!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal)
                ?? new Dictionary<string, W.Style>(StringComparer.Ordinal);
            StyleResolver = new WordStyleResolver(main);
            string? pageColor = main.Document?.DocumentBackground?.Color?.Value;
            PageBackground = TryParseWordColor(pageColor) ?? OfficeColor.White;
        }

        internal WordAccessibilityOptions Options { get; }
        internal Dictionary<string, W.Style> Styles { get; }
        internal WordStyleResolver StyleResolver { get; }
        internal OfficeColor PageBackground { get; }
        internal Dictionary<W.Paragraph, int> ParagraphIndexes { get; set; } = new();

        /// <summary>Strict promotes structural warnings to errors; heuristics stay warnings.</summary>
        internal OfficeAccessibilitySeverity Structural =>
            Options.Profile == OfficeAccessibilityPolicyProfile.Strict ? OfficeAccessibilitySeverity.Error : OfficeAccessibilitySeverity.Warning;

        internal int IndexOf(OpenXmlElement element) {
            W.Paragraph? paragraph = element as W.Paragraph ?? element.Descendants<W.Paragraph>().FirstOrDefault() ?? element.Ancestors<W.Paragraph>().FirstOrDefault();
            return paragraph != null && ParagraphIndexes.TryGetValue(paragraph, out int index) ? index : -1;
        }
    }

    private static void InspectHeadings(List<W.Paragraph> paragraphs, AccessibilityContext context, List<OfficeAccessibilityFinding> findings) {
        int previousLevel = 0;
        int headingCount = 0;
        int textParagraphCount = 0;
        for (int index = 0; index < paragraphs.Count; index++) {
            W.Paragraph paragraph = paragraphs[index];
            if (!string.IsNullOrWhiteSpace(GetParagraphText(paragraph))) textParagraphCount++;
            int level = GetHeadingLevel(paragraph, context.Styles);
            if (level == 0) continue;
            headingCount++;
            if (previousLevel > 0 && level > previousLevel + 1) {
                findings.Add(new OfficeAccessibilityFinding(context.Structural, OfficeDiagnosticCodes.AccessibilityHeadingLevelSkipped,
                    "Heading level " + level.ToString(CultureInfo.InvariantCulture) + " follows level " + previousLevel.ToString(CultureInfo.InvariantCulture) + "; heading levels should not be skipped.",
                    ParagraphLocation(index), index));
            }
            previousLevel = level;
        }

        if (headingCount == 0 && textParagraphCount >= LongDocumentParagraphCount) {
            findings.Add(new OfficeAccessibilityFinding(context.Structural, "accessibility-missing-headings",
                "The document has " + textParagraphCount.ToString(CultureInfo.InvariantCulture) + " paragraphs of text but no headings to navigate by.", "Document"));
        }
    }

    private static void InspectTable(W.Table table, int tableNumber, AccessibilityContext context, List<OfficeAccessibilityFinding> findings) {
        List<W.TableRow> rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count == 0) return;
        int paragraphIndex = context.IndexOf(table);
        string location = "Table " + tableNumber.ToString(CultureInfo.InvariantCulture);
        if (IsLikelyLayoutTable(table, rows)) {
            findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Warning, "accessibility-layout-table",
                "The table looks like it is used for layout; screen readers announce it as a data table.", location,
                paragraphIndex < 0 ? null : paragraphIndex, isHeuristic: true));
            return;
        }
        if (context.Options.RequireTableHeaders && rows.Count > 1 && !HasHeaderRow(rows)) {
            findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityMissingTableHeader,
                "The table does not mark its first row as a header row.", location, paragraphIndex < 0 ? null : paragraphIndex));
        }
    }

    private static void InspectParagraph(W.Paragraph paragraph, int index, AccessibilityContext context, List<OfficeAccessibilityFinding> findings) {
        WordAccessibilityOptions options = context.Options;
        foreach (OpenXmlElement child in paragraph.Descendants().Where(element => element is W.Drawing || element is W.Picture).ToList()) {
            if (child.Ancestors<W.Paragraph>().FirstOrDefault() != paragraph) continue;
            if (child is W.Drawing drawing) InspectDrawing(drawing, index, context, findings);
            else InspectVmlPicture((W.Picture)child, index, context, findings);
        }

        if (options.CheckMeaningfulLinks) {
            foreach (W.Hyperlink hyperlink in paragraph.Elements<W.Hyperlink>()) {
                string label = string.Concat(hyperlink.Descendants<W.Text>().Select(text => text.Text));
                bool hasImage = hyperlink.Descendants<W.Drawing>().Any();
                if (string.IsNullOrWhiteSpace(label) && !hasImage) {
                    findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityLinkTextMissing,
                        "The hyperlink has no visible text.", ParagraphLocation(index), index));
                } else if (!hasImage && !OfficeAccessibilityText.IsMeaningfulLinkLabel(label)) {
                    findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityUnclearLinkLabel,
                        "Hyperlink text '" + label.Trim() + "' does not describe its destination without surrounding context.", ParagraphLocation(index), index));
                }
            }
        }

        if (options.CheckContrast) InspectParagraphContrast(paragraph, index, context, findings);
    }

    private static void InspectDrawing(W.Drawing drawing, int index, AccessibilityContext context, List<OfficeAccessibilityFinding> findings) {
        DW.Anchor? anchor = drawing.GetFirstChild<DW.Anchor>();
        DW.DocProperties? properties = anchor?.GetFirstChild<DW.DocProperties>() ?? drawing.GetFirstChild<DW.Inline>()?.GetFirstChild<DW.DocProperties>();
        if (properties == null) return;
        string? name = properties.Name?.Value;
        uint? objectId = properties.Id?.Value;
        if (IsDecorative(properties)) return;
        string kind = DescribeGraphic(drawing.Descendants<A.GraphicData>().FirstOrDefault()?.Uri?.Value);
        bool textShape = kind == "shape" && drawing.Descendants<W.TextBoxContent>().Any(content => !string.IsNullOrWhiteSpace(content.InnerText));

        if (context.Options.RequireAlternativeText && !textShape &&
            string.IsNullOrWhiteSpace(properties.Description?.Value) && string.IsNullOrWhiteSpace(properties.Title?.Value)) {
            findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityMissingAlternativeText,
                "The " + kind + " is neither decorative nor described with alternative text.", ParagraphLocation(index), index, objectId, name));
        }
        if (context.Options.CheckFloatingObjects && anchor != null) {
            findings.Add(new OfficeAccessibilityFinding(context.Structural, "accessibility-floating-object",
                "The " + kind + " floats over the text instead of sitting inline, so its reading order is unclear.", ParagraphLocation(index), index, objectId, name));
        }
    }

    private static void InspectVmlPicture(W.Picture picture, int index, AccessibilityContext context, List<OfficeAccessibilityFinding> findings) {
        foreach (OpenXmlElement shape in picture.Elements().Where(element => element.LocalName == "shape")) {
            if (!shape.Descendants().Any(element => element.LocalName == "imagedata")) continue;
            string? alt = shape.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "alt").Value;
            if (context.Options.RequireAlternativeText && string.IsNullOrWhiteSpace(alt)) {
                findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityMissingAlternativeText,
                    "The image is neither decorative nor described with alternative text.", ParagraphLocation(index), index));
            }
        }
    }

    private static void InspectParagraphContrast(W.Paragraph paragraph, int index, AccessibilityContext context, List<OfficeAccessibilityFinding> findings) {
        foreach (W.Run run in paragraph.Descendants<W.Run>()) {
            if (run.Ancestors<W.Paragraph>().FirstOrDefault() != paragraph) continue;
            string text = string.Concat(run.Elements<W.Text>().Select(item => item.Text));
            if (string.IsNullOrWhiteSpace(text)) continue;
            EffectiveWordRunStyle style = context.StyleResolver.Resolve(run);
            if (style.Hidden) continue;
            OfficeColor? foreground = TryParseWordColor(style.Foreground);
            if (!foreground.HasValue) continue;
            OfficeColor background = TryParseHighlight(run.RunProperties?.Highlight?.Val?.InnerText)
                ?? TryParseWordColor(style.Background) ?? context.PageBackground;
            double size = style.FontSizePoints ?? 10D;
            bool large = size >= 18D || style.Bold && size >= 14D;
            double required = large ? context.Options.MinimumLargeTextContrastRatio : context.Options.MinimumTextContrastRatio;
            double measured = OfficeColorContrast.ContrastRatio(foreground.Value, background);
            if (measured + 0.000001D >= required) continue;
            findings.Add(new OfficeAccessibilityFinding(OfficeAccessibilitySeverity.Error, OfficeDiagnosticCodes.AccessibilityLowContrast,
                "Text contrast is " + measured.ToString("0.##", CultureInfo.InvariantCulture) + ":1; the policy requires at least " +
                required.ToString("0.##", CultureInfo.InvariantCulture) + ":1.", ParagraphLocation(index), index,
                measuredValue: measured, requiredValue: required));
            return;
        }
    }

    private static void InspectBlankParagraphs(W.Body body, AccessibilityContext context, List<OfficeAccessibilityFinding> findings) {
        W.Paragraph? first = null;
        int count = 0;
        foreach (OpenXmlElement element in body.Elements().Append(null!)) {
            if (element is W.Paragraph paragraph && IsBlankParagraph(paragraph)) {
                if (count++ == 0) first = paragraph;
                continue;
            }
            if (count >= 2 && first != null) {
                int index = context.IndexOf(first);
                findings.Add(new OfficeAccessibilityFinding(context.Structural, "accessibility-blank-paragraph-spacing",
                    count.ToString(CultureInfo.InvariantCulture) + " consecutive empty paragraphs are used for spacing; use paragraph spacing or a page break instead.",
                    ParagraphLocation(index), index));
            }
            count = 0;
            first = null;
        }
    }

    private static bool IsBlankParagraph(W.Paragraph paragraph) {
        if (paragraph.ParagraphProperties?.SectionProperties != null) return false;
        foreach (OpenXmlElement element in paragraph.Descendants()) {
            if (element is W.Text text && text.Text.Length > 0) return false;
            if (element is W.Drawing || element is W.Picture || element is W.EmbeddedObject || element is W.Break ||
                element is W.FieldChar || element is W.SimpleField || element is W.TabChar || element is W.SymbolChar) return false;
            if (element.LocalName == "oMath" || element.LocalName == "oMathPara") return false;
        }
        return true;
    }

    private static bool HasHeaderRow(List<W.TableRow> rows) {
        W.TableHeader? header = rows[0].TableRowProperties?.GetFirstChild<W.TableHeader>();
        return header != null && header.Val?.Value != W.OnOffOnlyValues.Off;
    }

    /// <summary>A table with one row or one column, or a borderless table holding images or nested tables, is probably used for layout.</summary>
    private static bool IsLikelyLayoutTable(W.Table table, List<W.TableRow> rows) {
        int columns = rows.Max(row => row.Elements<W.TableCell>().Count());
        if (rows.Count == 1 && columns > 1) return true;
        if (table.Descendants<W.Table>().Any()) return true;
        bool borderless = IsBorderless(table.GetFirstChild<W.TableProperties>());
        bool holdsGraphics = table.Descendants<W.Drawing>().Any();
        return borderless && holdsGraphics;
    }

    private static bool IsBorderless(W.TableProperties? properties) {
        if (properties?.TableStyle?.Val?.Value != null) return false;
        W.TableBorders? borders = properties?.TableBorders;
        if (borders == null) return true;
        return borders.Elements<W.BorderType>().All(border => border.Val == null || border.Val.Value == W.BorderValues.None || border.Val.Value == W.BorderValues.Nil);
    }

    private static int GetHeadingLevel(W.Paragraph paragraph, Dictionary<string, W.Style> styles) {
        int? outline = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (outline == null && styleId != null && visited.Add(styleId) && styles.TryGetValue(styleId, out W.Style? style)) {
            outline = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (outline == null) {
                string name = style.StyleName?.Val?.Value ?? styleId;
                if (name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) &&
                    int.TryParse(name.Substring(8), NumberStyles.Integer, CultureInfo.InvariantCulture, out int level) && level is >= 1 and <= 9) {
                    return level;
                }
            }
            styleId = style.BasedOn?.Val?.Value;
        }
        return outline is >= 0 and <= 8 ? outline.Value + 1 : 0;
    }

    private static string? GetDefaultLanguage(MainDocumentPart main) {
        W.Styles? styles = main.StyleDefinitionsPart?.Styles;
        string? language = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<W.Languages>()?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(language)) return language;
        W.Style? normal = styles?.Elements<W.Style>().FirstOrDefault(style => style.Default?.Value == true && style.Type?.Value == W.StyleValues.Paragraph);
        language = normal?.StyleRunProperties?.GetFirstChild<W.Languages>()?.Val?.Value;
        return string.IsNullOrWhiteSpace(language) ? null : language;
    }

    private static void SetDefaultLanguage(MainDocumentPart main, string language) {
        StyleDefinitionsPart stylesPart = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        W.Styles styles = stylesPart.Styles ??= new W.Styles();
        W.DocDefaults defaults = styles.DocDefaults ?? styles.PrependChild(new W.DocDefaults());
        W.RunPropertiesDefault runDefaults = defaults.RunPropertiesDefault ?? defaults.PrependChild(new W.RunPropertiesDefault());
        W.RunPropertiesBaseStyle properties = runDefaults.RunPropertiesBaseStyle ?? runDefaults.AppendChild(new W.RunPropertiesBaseStyle());
        W.Languages? languages = properties.GetFirstChild<W.Languages>();
        if (languages == null) {
            languages = new W.Languages();
            // Schema order puts lang after the size and highlight elements and before eastAsianLayout.
            OpenXmlElement? before = properties.ChildElements.FirstOrDefault(element => element.LocalName is "eastAsianLayout" or "specVanish" or "oMath");
            if (before != null) properties.InsertBefore(languages, before); else properties.AppendChild(languages);
        }
        languages.Val = language;
    }

    private static bool IsDecorative(DW.DocProperties properties) =>
        properties.Descendants().Any(element => element.LocalName == "decorative" &&
            element.GetAttributes().Any(attribute => attribute.LocalName == "val" && (attribute.Value == "1" || string.Equals(attribute.Value, "true", StringComparison.OrdinalIgnoreCase))));

    private static string DescribeGraphic(string? uri) {
        if (uri == null) return "object";
        if (uri.EndsWith("/picture", StringComparison.Ordinal)) return "image";
        if (uri.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0) return "chart";
        if (uri.EndsWith("/diagram", StringComparison.Ordinal)) return "SmartArt graphic";
        if (uri.EndsWith("wordprocessingShape", StringComparison.Ordinal)) return "shape";
        if (uri.EndsWith("wordprocessingGroup", StringComparison.Ordinal) || uri.EndsWith("wordprocessingCanvas", StringComparison.Ordinal)) return "group";
        return "object";
    }


    private static string ParagraphLocation(int index) => index < 0 ? "Document" : "Paragraph " + (index + 1).ToString(CultureInfo.InvariantCulture);

    private static OfficeColor? TryParseWordColor(string? value) {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) return null;
        return OfficeColor.TryParseHex("#" + value, out OfficeColor color) ? color : null;
    }

    // ST_HighlightColor names Word's fixed palette; several names (green, darkBlue, darkMagenta...) differ from CSS.
    private static OfficeColor? TryParseHighlight(string? value) => value switch {
        "black" => OfficeColor.FromRgb(0, 0, 0),
        "blue" => OfficeColor.FromRgb(0, 0, 255),
        "cyan" => OfficeColor.FromRgb(0, 255, 255),
        "green" => OfficeColor.FromRgb(0, 255, 0),
        "magenta" => OfficeColor.FromRgb(255, 0, 255),
        "red" => OfficeColor.FromRgb(255, 0, 0),
        "yellow" => OfficeColor.FromRgb(255, 255, 0),
        "white" => OfficeColor.FromRgb(255, 255, 255),
        "darkBlue" => OfficeColor.FromRgb(0, 0, 128),
        "darkCyan" => OfficeColor.FromRgb(0, 128, 128),
        "darkGreen" => OfficeColor.FromRgb(0, 128, 0),
        "darkMagenta" => OfficeColor.FromRgb(128, 0, 128),
        "darkRed" => OfficeColor.FromRgb(128, 0, 0),
        "darkYellow" => OfficeColor.FromRgb(128, 128, 0),
        "darkGray" => OfficeColor.FromRgb(128, 128, 128),
        "lightGray" => OfficeColor.FromRgb(192, 192, 192),
        _ => null
    };
}
