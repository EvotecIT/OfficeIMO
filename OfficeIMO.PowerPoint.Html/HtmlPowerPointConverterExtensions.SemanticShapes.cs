using AngleSharp.Dom;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static void ImportSemanticShapes(
        IElement section,
        PptCore.PowerPointSlide slide,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result) {
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
                    contentTop = ImportSemanticTextBox(item.Element, slide, contentTop, result);
                    break;
                case PowerPointSemanticImportKind.Table:
                    contentTop = ImportTable(item.Element, slide, contentTop, options, result);
                    break;
                case PowerPointSemanticImportKind.Picture:
                    ImportPicture(item.Element, slide, result, ref pictureTop);
                    break;
                case PowerPointSemanticImportKind.Chart:
                    ImportChart(item.Element, slide, result, ref chartTop);
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
        HtmlToPowerPointResult result) {
        string text = PreserveText(paragraph.TextContent);
        if (text.Length == 0) {
            return fallbackTop;
        }

        ReadSemanticShapeGeometry(paragraph, 64D, fallbackTop, 620D, 48D,
            out double left, out double top, out double width, out double height);
        PptCore.PowerPointTextBox textBox = slide.AddTextBoxPoints(text, left, top, width, height);
        ApplyShapeTransforms(paragraph, textBox);
        result.TextBoxes++;
        return Math.Max(fallbackTop + 58D, top + height + 10D);
    }

    private static void ReadSemanticShapeGeometry(
        IElement element,
        double fallbackLeft,
        double fallbackTop,
        double fallbackWidth,
        double fallbackHeight,
        out double left,
        out double top,
        out double width,
        out double height) {
        left = ReadOptionalDoubleAttribute(element, "data-officeimo-left") ?? fallbackLeft;
        top = ReadOptionalDoubleAttribute(element, "data-officeimo-top") ?? fallbackTop;
        width = Math.Max(1D, ReadOptionalDoubleAttribute(element, "data-officeimo-width") ?? fallbackWidth);
        height = Math.Max(1D, ReadOptionalDoubleAttribute(element, "data-officeimo-height") ?? fallbackHeight);
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
