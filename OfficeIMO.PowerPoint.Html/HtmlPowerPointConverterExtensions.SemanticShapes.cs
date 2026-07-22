using AngleSharp.Dom;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static void ImportSemanticShapes(
        IElement section,
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
        foreach (PowerPointSemanticImportItem item in items
            .OrderBy(item => item.LayerIndex ?? item.FallbackOrder)
            .ThenBy(item => item.FallbackOrder)) {
            switch (item.Kind) {
                case PowerPointSemanticImportKind.TextBox:
                    contentTop = ImportSemanticTextBox(item.Element, slide, contentTop, result, budget);
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
        PptCore.PowerPointSlide slide,
        double fallbackTop,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        string text = PreserveText(paragraph.TextContent);
        return ImportTextBox(paragraph, text, slide, fallbackTop, result, budget, 48D);
    }

    private static double ImportTextBox(
        IElement source,
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
                lossKind: HtmlConversionLossKind.Omission, detail: metadataLimit);
            return fallbackTop;
        }

        if (!budget.TryReserveShape(out string shapeLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "A slide text block was omitted because the shared shape limit was reached.",
                lossKind: HtmlConversionLossKind.Omission, detail: shapeLimit);
            return fallbackTop;
        }

        ReadSemanticShapeGeometry(source, 64D, fallbackTop, 620D, fallbackHeight, budget, result,
            out double left, out double top, out double width, out double height);
        PptCore.PowerPointTextBox textBox = slide.AddTextBoxPoints(text, left, top, width, height);
        if (semanticBlock?.Kind == HtmlSemanticBlockKind.List) {
            ApplySemanticList(textBox, semanticBlock);
        } else if (semanticBlock != null && semanticBlock.Runs.Count > 0) {
            ApplySemanticRuns(textBox.Paragraphs[0], semanticBlock.Runs);
        }
        ApplyShapeTransforms(source, textBox, budget, result);
        result.TextBoxes++;
        return Math.Max(fallbackTop + 58D, top + height + 10D);
    }

    private static void ApplySemanticList(PptCore.PowerPointTextBox textBox, HtmlSemanticBlock list) {
        var items = new List<SemanticListItem>();
        AppendSemanticListItems(list, 0, items);
        if (items.Count == 0) return;
        textBox.Text = string.Join("\n", items.Select(item => item.Block.Text));
        IReadOnlyList<PptCore.PowerPointParagraph> paragraphs = textBox.Paragraphs;
        for (int index = 0; index < Math.Min(items.Count, paragraphs.Count); index++) {
            SemanticListItem item = items[index];
            PptCore.PowerPointParagraph paragraph = paragraphs[index];
            if (item.Ordered) paragraph.SetNumbered(index + 1);
            else paragraph.SetBullet();
            paragraph.Level = Math.Min(8, item.Level);
            ApplySemanticRuns(paragraph, item.Block.Runs);
        }
    }

    private static void AppendSemanticListItems(HtmlSemanticBlock list, int level, ICollection<SemanticListItem> result) {
        foreach (HtmlSemanticBlock item in list.Children) {
            result.Add(new SemanticListItem(item, list.Ordered, level));
            foreach (HtmlSemanticBlock nested in item.Children.Where(child => child.Kind == HtmlSemanticBlockKind.List)) {
                AppendSemanticListItems(nested, level + 1, result);
            }
        }
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
        target.Underline = source.Underline;
        target.Strikethrough = source.Strikethrough;
        string color = NormalizeSemanticColor(source.Style?.GetValue("color"));
        if (color.Length > 0) target.Color = color;
        string fontName = NormalizeSemanticFontName(source.Style?.GetValue("font-family"));
        if (fontName.Length > 0) target.FontName = fontName;
        if (TryParseSemanticPixels(source.Style?.GetValue("font-size"), out double pixels)) {
            target.FontSize = Math.Max(1, (int)Math.Round(pixels * 0.75D));
        }
        if (!string.IsNullOrWhiteSpace(source.Hyperlink)
            && Uri.TryCreate(source.Hyperlink, UriKind.Absolute, out Uri? hyperlink)) {
            target.Hyperlink = hyperlink;
        }
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
        if (!text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) return false;
        return double.TryParse(text.Substring(0, text.Length - 2), NumberStyles.Float,
            CultureInfo.InvariantCulture, out pixels) && pixels > 0D;
    }

    private sealed class SemanticListItem {
        internal SemanticListItem(HtmlSemanticBlock block, bool ordered, int level) {
            Block = block;
            Ordered = ordered;
            Level = level;
        }
        internal HtmlSemanticBlock Block { get; }
        internal bool Ordered { get; }
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
