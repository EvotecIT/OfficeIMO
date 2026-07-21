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
        double fallbackHeight) {
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
        ApplyShapeTransforms(source, textBox, budget, result);
        result.TextBoxes++;
        return Math.Max(fallbackTop + 58D, top + height + 10D);
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
