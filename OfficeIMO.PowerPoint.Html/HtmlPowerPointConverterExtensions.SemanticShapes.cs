using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Drawing;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static void ImportSemanticShapes(
        IElement section,
        HtmlSemanticSection? semanticSection,
        PptCore.PowerPointSlide slide,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        var items = new List<PowerPointSemanticImportItem>();
        int fallbackOrder = 0;

        foreach (IElement element in section.Children) {
            if (IsElement(element, "p")) {
                items.Add(CreateSemanticImportItem(element, PowerPointSemanticImportKind.TextBox, fallbackOrder++));
            } else if (options.ImportTables && IsElement(element, "table")) {
                items.Add(CreateSemanticImportItem(element, PowerPointSemanticImportKind.Table, fallbackOrder++));
            }
        }

        if (options.ImportPictures) {
            foreach (IElement item in section.QuerySelectorAll("section.officeimo-images li")) {
                items.Add(CreateSemanticImportItem(item, PowerPointSemanticImportKind.Picture, fallbackOrder++));
            }
        }

        if (options.ImportChartInventory) {
            foreach (IElement item in section.QuerySelectorAll("section.officeimo-charts li")) {
                items.Add(CreateSemanticImportItem(item, PowerPointSemanticImportKind.Chart, fallbackOrder++));
            }
        }

        double contentTop = 48D;
        double pictureTop = 140D;
        double chartTop = 220D;
        HtmlSemanticBlock[] semanticTextBlocks = semanticSection?.Blocks
            .Where(block => block.Kind is HtmlSemanticBlockKind.Paragraph or HtmlSemanticBlockKind.Heading)
            .ToArray() ?? Array.Empty<HtmlSemanticBlock>();
        int semanticTextIndex = 0;
        foreach (PowerPointSemanticImportItem item in items
            .OrderBy(item => item.LayerIndex ?? item.FallbackOrder)
            .ThenBy(item => item.FallbackOrder)) {
            switch (item.Kind) {
                case PowerPointSemanticImportKind.TextBox:
                    HtmlSemanticBlock? semanticBlock = semanticTextIndex < semanticTextBlocks.Length
                        ? semanticTextBlocks[semanticTextIndex]
                        : null;
                    semanticTextIndex++;
                    contentTop = ImportSemanticTextBox(item.Element, semanticBlock, slide, contentTop, result, budget);
                    break;
                case PowerPointSemanticImportKind.Table:
                    contentTop = ImportTable(item.Element, slide, contentTop, result, budget);
                    break;
                case PowerPointSemanticImportKind.Picture:
                    ImportPicture(item.Element, slide, result, budget, ref pictureTop);
                    break;
                case PowerPointSemanticImportKind.Chart:
                    ImportChart(item.Element, slide, result, budget, ref chartTop);
                    break;
            }
        }
    }

    private static PowerPointSemanticImportItem CreateSemanticImportItem(
        IElement element,
        PowerPointSemanticImportKind kind,
        int fallbackOrder) =>
        new(element, kind, ReadOptionalIntAttribute(element, "data-officeimo-layer-index"), fallbackOrder);

    private static double ImportSemanticTextBox(
        IElement paragraph,
        HtmlSemanticBlock? semanticBlock,
        PptCore.PowerPointSlide slide,
        double fallbackTop,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        string text = PreserveText(paragraph.TextContent);
        return ImportTextBox(paragraph, text, slide, fallbackTop, result, budget, 48D, semanticBlock);
    }

    private static double ImportTextBox(
        IElement? source,
        string text,
        PptCore.PowerPointSlide slide,
        double fallbackTop,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        double fallbackHeight,
        HtmlSemanticBlock? semanticBlock = null) {
        if (text.Length == 0) {
            return fallbackTop;
        }

        if (!budget.IsMetadataWithinLimit(text, out string metadataLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "A slide text block was omitted because it exceeded the shared field limit.",
                lossKind: OfficeConversionLossKind.Omission, detail: metadataLimit);
            return fallbackTop;
        }

        if (!budget.TryReserveShape(out string shapeLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "A slide text block was omitted because the shared shape limit was reached.",
                lossKind: OfficeConversionLossKind.Omission, detail: shapeLimit);
            return fallbackTop;
        }

        double left = 64D;
        double top = fallbackTop;
        double width = 620D;
        double height = fallbackHeight;
        if (source != null) {
            ReadSemanticShapeGeometry(source, left, top, width, height, budget, result,
                out left, out top, out width, out height);
        }
        PptCore.PowerPointTextBox textBox = slide.AddTextBoxPoints(text, left, top, width, height);
        if (semanticBlock?.Kind == HtmlSemanticBlockKind.List) {
            ApplySemanticList(textBox, semanticBlock, result);
        } else if (source?.QuerySelector("span") != null) {
            ApplyTargetSemanticRuns(textBox.Paragraphs[0], source);
        } else if (semanticBlock != null && semanticBlock.Runs.Count > 0) {
            ApplySemanticRuns(textBox.Paragraphs[0], semanticBlock.Runs);
        }
        if (source != null) ApplyShapeTransforms(source, textBox, budget, result);
        result.TextBoxes++;
        return Math.Max(fallbackTop + 58D, top + height + 10D);
    }

    private static void ApplyTargetSemanticRuns(PptCore.PowerPointParagraph paragraph, IElement source) {
        IElement[] spans = source.QuerySelectorAll("span").ToArray();
        if (spans.Length == 0) return;
        string combined = string.Concat(spans.Select(span => span.TextContent));
        if (!string.Equals(combined, source.TextContent, StringComparison.Ordinal)) return;

        paragraph.Text = combined;
        IReadOnlyList<PptCore.PowerPointTextRun> targetRuns = paragraph.Runs;
        ApplyTargetSemanticRun(targetRuns[0], spans[0]);
        for (int index = 1; index < spans.Length; index++) {
            PptCore.PowerPointTextRun target = paragraph.AddRun(spans[index].TextContent);
            ApplyTargetSemanticRun(target, spans[index]);
        }
    }

    private static void ApplyTargetSemanticRun(PptCore.PowerPointTextRun target, IElement source) {
        IReadOnlyDictionary<string, string> css = ParseTargetInlineStyle(source.GetAttribute("style"));
        target.Text = source.TextContent;
        target.Bold = TryGetTargetCss(css, "font-weight", out string weight)
            && (weight.Equals("bold", StringComparison.OrdinalIgnoreCase)
                || int.TryParse(weight, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numericWeight) && numericWeight >= 600);
        target.Italic = TryGetTargetCss(css, "font-style", out string fontStyle)
            && (fontStyle.Equals("italic", StringComparison.OrdinalIgnoreCase) || fontStyle.Equals("oblique", StringComparison.OrdinalIgnoreCase));
        target.UnderlineStyle = ResolveTargetUnderline(source, css);
        target.StrikeStyle = ResolveTargetStrike(source, css);
        target.BaselinePercent = ResolveTargetBaseline(source, css);
        if (Enum.TryParse(source.GetAttribute("data-officeimo-powerpoint-capitalization"), true,
                out PptCore.PowerPointCapitalization capitalization)) {
            target.Capitalization = capitalization;
        } else if (TryGetTargetCss(css, "font-variant", out string variant)
                   && variant.IndexOf("small-caps", StringComparison.OrdinalIgnoreCase) >= 0) {
            target.Capitalization = PptCore.PowerPointCapitalization.SmallCaps;
        } else if (TryGetTargetCss(css, "text-transform", out string transform)
                   && transform.Equals("uppercase", StringComparison.OrdinalIgnoreCase)) {
            target.Capitalization = PptCore.PowerPointCapitalization.AllCaps;
        }
        if (TryGetTargetCss(css, "font-family", out string family)) target.FontName = NormalizeSemanticFontName(family);
        if (TryParseSemanticPixels(TryGetTargetCss(css, "font-size", out string size) ? size : null, out double pixels)) {
            target.FontSizePoints = Math.Max(1D, pixels * 0.75D);
        }
        if (TryGetTargetCss(css, "color", out string color)) {
            string normalized = NormalizeSemanticColor(color);
            if (normalized.Length > 0) target.Color = normalized;
        }
    }

    private static PptCore.PowerPointUnderlineStyle? ResolveTargetUnderline(IElement source, IReadOnlyDictionary<string, string> css) {
        if (Enum.TryParse(source.GetAttribute("data-officeimo-powerpoint-underline"), true,
                out PptCore.PowerPointUnderlineStyle native)) return native;
        if (!HasTargetDecoration(css, "underline")) return null;
        return TryGetTargetCss(css, "text-decoration-style", out string style) ? style.ToLowerInvariant() switch {
            "double" => PptCore.PowerPointUnderlineStyle.Double,
            "dotted" => PptCore.PowerPointUnderlineStyle.Dotted,
            "dashed" => PptCore.PowerPointUnderlineStyle.Dash,
            "wavy" => PptCore.PowerPointUnderlineStyle.Wavy,
            _ => PptCore.PowerPointUnderlineStyle.Single
        } : PptCore.PowerPointUnderlineStyle.Single;
    }

    private static PptCore.PowerPointStrikeStyle? ResolveTargetStrike(IElement source, IReadOnlyDictionary<string, string> css) {
        if (Enum.TryParse(source.GetAttribute("data-officeimo-powerpoint-strike"), true,
                out PptCore.PowerPointStrikeStyle native)) return native;
        if (!HasTargetDecoration(css, "line-through")) return null;
        return TryGetTargetCss(css, "text-decoration-style", out string style)
               && style.Equals("double", StringComparison.OrdinalIgnoreCase)
            ? PptCore.PowerPointStrikeStyle.Double
            : PptCore.PowerPointStrikeStyle.Single;
    }

    private static double? ResolveTargetBaseline(IElement source, IReadOnlyDictionary<string, string> css) {
        if (double.TryParse(source.GetAttribute("data-officeimo-powerpoint-baseline-percent"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out double native) && native >= -100D && native <= 100D) return native;
        if (!TryGetTargetCss(css, "vertical-align", out string vertical)) return null;
        if (vertical.Equals("super", StringComparison.OrdinalIgnoreCase)) return 30D;
        if (vertical.Equals("sub", StringComparison.OrdinalIgnoreCase)) return -25D;
        return null;
    }

    private static IReadOnlyDictionary<string, string> ParseTargetInlineStyle(string? value) {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string declaration in (value ?? string.Empty).Split(';')) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            string name = declaration.Substring(0, separator).Trim();
            string content = declaration.Substring(separator + 1).Trim();
            if (name.Length > 0 && content.Length > 0) result[name] = content;
        }
        return result;
    }

    private static bool TryGetTargetCss(IReadOnlyDictionary<string, string> css, string name, out string value) =>
        css.TryGetValue(name, out value!);

    private static bool HasTargetDecoration(IReadOnlyDictionary<string, string> css, string value) =>
        (TryGetTargetCss(css, "text-decoration-line", out string lines) || TryGetTargetCss(css, "text-decoration", out lines))
        && lines.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
            .Any(item => item.Equals(value, StringComparison.OrdinalIgnoreCase));

    private static void ApplySemanticList(
        PptCore.PowerPointTextBox textBox,
        HtmlSemanticBlock list,
        HtmlToPowerPointResult result) {
        var items = new List<SemanticListItem>();
        AppendSemanticListItems(list, 0, items, result);
        if (items.Count == 0) return;
        textBox.Text = string.Join("\n", items.Select(item => item.Block.Text));
        IReadOnlyList<PptCore.PowerPointParagraph> paragraphs = textBox.Paragraphs;
        for (int index = 0; index < Math.Min(items.Count, paragraphs.Count); index++) {
            SemanticListItem item = items[index];
            PptCore.PowerPointParagraph paragraph = paragraphs[index];
            if (item.Ordered) {
                if (item.ShouldRestart) {
                    paragraph.SetNumbered(item.Ordinal ?? 1);
                } else {
                    paragraph.SetNumbered(PptCore.PowerPointNumberingScheme.ArabicPeriod);
                }
            } else {
                paragraph.SetBullet();
            }
            paragraph.Level = Math.Min(8, item.Level);
            ApplySemanticRuns(paragraph, item.Block.Runs);
        }
    }

    private static void AppendSemanticListItems(
        HtmlSemanticBlock list,
        int level,
        ICollection<SemanticListItem> result,
        HtmlToPowerPointResult conversionResult) {
        int? previousOrdinal = null;
        foreach (HtmlSemanticBlock item in list.Children) {
            int? ordinal = list.Ordered
                ? NormalizePowerPointListOrdinal(item.ListItem?.Ordinal ?? 1, conversionResult)
                : null;
            bool shouldRestart = list.Ordered && (!previousOrdinal.HasValue
                || item.ListItem?.ExplicitOrdinal.HasValue == true
                || list.List?.IsReversed == true
                || (long)ordinal.GetValueOrDefault() != (long)previousOrdinal.Value + 1L);
            result.Add(new SemanticListItem(item, list.Ordered, ordinal, shouldRestart, level));
            previousOrdinal = ordinal;
            foreach (HtmlSemanticBlock nested in item.Children.Where(child => child.Kind == HtmlSemanticBlockKind.List)) {
                AppendSemanticListItems(nested, level + 1, result, conversionResult);
            }
        }
    }

    private static int NormalizePowerPointListOrdinal(int ordinal, HtmlToPowerPointResult result) {
        const int minimum = 1;
        const int maximum = 32767;
        int normalized = Math.Max(minimum, Math.Min(maximum, ordinal));
        if (normalized != ordinal) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "An HTML list ordinal outside PowerPoint's supported range was clamped to "
                + normalized.ToString(CultureInfo.InvariantCulture) + ".",
                lossKind: OfficeConversionLossKind.Approximation,
                source: "list ordinal",
                detail: "Ordinal=" + ordinal.ToString(CultureInfo.InvariantCulture) + "; Supported=1..32767");
        }
        return normalized;
    }

    private static void ApplySemanticRuns(PptCore.PowerPointParagraph paragraph, IReadOnlyList<HtmlSemanticRun> runs) {
        if (runs.Count == 0) return;
        paragraph.Text = string.Concat(runs.Select(run => run.Text));
        IReadOnlyList<PptCore.PowerPointTextRun> targetRuns = paragraph.Runs;
        PptCore.PowerPointTextRun first = targetRuns[0];
        ApplySemanticRun(first, runs[0]);
        for (int index = 1; index < runs.Count; index++) {
            HtmlSemanticRun source = runs[index];
            PptCore.PowerPointTextRun target = paragraph.AddRun(source.Text);
            ApplySemanticRun(target, source);
        }
    }

    private static void ApplySemanticRun(PptCore.PowerPointTextRun target, HtmlSemanticRun source) {
        target.Text = source.Text;
        target.Bold = source.Bold;
        target.Italic = source.Italic;
        target.UnderlineStyle = ResolvePowerPointUnderline(source);
        target.StrikeStyle = ResolvePowerPointStrike(source);
        target.BaselinePercent = ResolvePowerPointBaseline(source);
        if (source.DataAttributes.TryGetValue("data-officeimo-powerpoint-capitalization", out string? exactCapitalization)
            && Enum.TryParse(exactCapitalization, ignoreCase: true, out PptCore.PowerPointCapitalization capitalization)) {
            target.Capitalization = capitalization;
        } else {
            string fontVariant = source.Style?.GetValue("font-variant") ?? string.Empty;
            string textTransform = source.Style?.GetValue("text-transform") ?? string.Empty;
            if (fontVariant.IndexOf("small-caps", StringComparison.OrdinalIgnoreCase) >= 0) {
                target.Capitalization = PptCore.PowerPointCapitalization.SmallCaps;
            } else if (string.Equals(textTransform.Trim(), "uppercase", StringComparison.OrdinalIgnoreCase)) {
                target.Capitalization = PptCore.PowerPointCapitalization.AllCaps;
            }
        }
        string color = NormalizeSemanticColor(source.Style?.GetValue("color"));
        if (color.Length > 0) target.Color = color;
        string fontName = NormalizeSemanticFontName(source.Style?.GetValue("font-family"));
        if (fontName.Length > 0) target.FontName = fontName;
        if (TryParseSemanticPixels(source.Style?.GetValue("font-size"), out double pixels)) {
            target.FontSizePoints = Math.Max(1D, pixels * 0.75D);
        }
        if (!string.IsNullOrWhiteSpace(source.Hyperlink)
            && Uri.TryCreate(source.Hyperlink, UriKind.Absolute, out Uri? hyperlink)) {
            target.Hyperlink = hyperlink;
        }
    }

    private static PptCore.PowerPointUnderlineStyle? ResolvePowerPointUnderline(HtmlSemanticRun source) {
        if (source.DataAttributes.TryGetValue("data-officeimo-powerpoint-underline", out string? exact)
            && Enum.TryParse(exact, ignoreCase: true, out PptCore.PowerPointUnderlineStyle native)) return native;
        return source.UnderlineStyle switch {
            OfficeTextDecorationStyle.None => null,
            OfficeTextDecorationStyle.Double => PptCore.PowerPointUnderlineStyle.Double,
            OfficeTextDecorationStyle.Dotted => PptCore.PowerPointUnderlineStyle.Dotted,
            OfficeTextDecorationStyle.Dashed => PptCore.PowerPointUnderlineStyle.Dash,
            OfficeTextDecorationStyle.Wavy => PptCore.PowerPointUnderlineStyle.Wavy,
            _ => PptCore.PowerPointUnderlineStyle.Single
        };
    }

    private static PptCore.PowerPointStrikeStyle? ResolvePowerPointStrike(HtmlSemanticRun source) {
        if (source.DataAttributes.TryGetValue("data-officeimo-powerpoint-strike", out string? exact)
            && Enum.TryParse(exact, ignoreCase: true, out PptCore.PowerPointStrikeStyle native)) return native;
        return source.StrikethroughStyle switch {
            OfficeTextDecorationStyle.None => null,
            OfficeTextDecorationStyle.Double => PptCore.PowerPointStrikeStyle.Double,
            _ => PptCore.PowerPointStrikeStyle.Single
        };
    }

    private static double? ResolvePowerPointBaseline(HtmlSemanticRun source) {
        if (source.DataAttributes.TryGetValue("data-officeimo-powerpoint-baseline-percent", out string? exact)
            && double.TryParse(exact, NumberStyles.Float, CultureInfo.InvariantCulture, out double native)
            && native >= -100D && native <= 100D) return native;
        return source.Baseline switch {
            OfficeTextBaseline.Superscript => 30D,
            OfficeTextBaseline.Subscript => -25D,
            _ => (double?)null
        };
    }

    private static string NormalizeSemanticColor(string? value) {
        string color = (value ?? string.Empty).Trim();
        if (color.Length == 7 && color[0] == '#') return color.Substring(1).ToUpperInvariant();
        if (color.Length == 4 && color[0] == '#') {
            return string.Concat(char.ToUpperInvariant(color[1]), char.ToUpperInvariant(color[1]),
                char.ToUpperInvariant(color[2]), char.ToUpperInvariant(color[2]),
                char.ToUpperInvariant(color[3]), char.ToUpperInvariant(color[3]));
        }
        return string.Empty;
    }

    private static string NormalizeSemanticFontName(string? value) =>
        (value ?? string.Empty).Split(',').FirstOrDefault()?.Trim().Trim('\'', '"') ?? string.Empty;

    private static bool TryParseSemanticPixels(string? value, out double pixels) {
        pixels = 0D;
        string text = (value ?? string.Empty).Trim();
        bool points = text.EndsWith("pt", StringComparison.OrdinalIgnoreCase);
        if (!points && !text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) return false;
        if (!double.TryParse(text.Substring(0, text.Length - 2), NumberStyles.Float,
                CultureInfo.InvariantCulture, out pixels) || pixels <= 0D) return false;
        if (points) pixels /= 0.75D;
        return true;
    }

    private sealed class SemanticListItem {
        internal SemanticListItem(HtmlSemanticBlock block, bool ordered, int? ordinal, bool shouldRestart, int level) {
            Block = block;
            Ordered = ordered;
            Ordinal = ordinal;
            ShouldRestart = shouldRestart;
            Level = level;
        }
        internal HtmlSemanticBlock Block { get; }
        internal bool Ordered { get; }
        internal int? Ordinal { get; }
        internal bool ShouldRestart { get; }
        internal int Level { get; }
    }

    private static void ReadSemanticShapeGeometry(
        IElement element,
        double fallbackLeft,
        double fallbackTop,
        double fallbackWidth,
        double fallbackHeight,
        HtmlImportBudget budget,
        HtmlToPowerPointResult result,
        out double left,
        out double top,
        out double width,
        out double height) {
        left = NormalizeGeometry(ReadOptionalDoubleAttribute(element, "data-officeimo-left") ?? fallbackLeft, fallbackLeft, -budget.Limits.MaxAbsoluteGeometry, budget, result, "shape left");
        top = NormalizeGeometry(ReadOptionalDoubleAttribute(element, "data-officeimo-top") ?? fallbackTop, fallbackTop, -budget.Limits.MaxAbsoluteGeometry, budget, result, "shape top");
        width = NormalizeGeometry(ReadOptionalDoubleAttribute(element, "data-officeimo-width") ?? fallbackWidth, fallbackWidth, 1D, budget, result, "shape width");
        height = NormalizeGeometry(ReadOptionalDoubleAttribute(element, "data-officeimo-height") ?? fallbackHeight, fallbackHeight, 1D, budget, result, "shape height");
    }

    private sealed class PowerPointSemanticImportItem {
        internal PowerPointSemanticImportItem(
            IElement element,
            PowerPointSemanticImportKind kind,
            int? layerIndex,
            int fallbackOrder) {
            Element = element;
            Kind = kind;
            LayerIndex = layerIndex;
            FallbackOrder = fallbackOrder;
        }

        internal IElement Element { get; }

        internal PowerPointSemanticImportKind Kind { get; }

        internal int? LayerIndex { get; }

        internal int FallbackOrder { get; }
    }

    private enum PowerPointSemanticImportKind {
        TextBox,
        Table,
        Picture,
        Chart
    }
}
