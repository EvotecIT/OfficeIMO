using System.Globalization;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private static bool HasSourceLinkedNumericAxisReference(OfficeXmlNode plot, OfficeXmlNode axis) {
        uint? id = axis.Child(C, "axId").UInt32Val();
        if (!id.HasValue) return false;
        foreach (OfficeXmlNode layer in plot.Children()) {
            uint?[] ids = ReadImageExportLayerAxisIds(layer);
            if (!ids.Contains(id)) continue;
            bool horizontal = (layer.Is(C, "scatterChart") || layer.Is(C, "bubbleChart")) && ids.FirstOrDefault() == id;
            foreach (OfficeXmlNode series in layer.Children()) {
                if (series.LocalName != "ser") continue;
                if (GetSourceLinkedAxisSource(series, axis, horizontal).Child(C, "numRef").Exists) return true;
            }
        }
        return false;
    }

    // The first three c:axId values of a chart group, or the axis's own id for an axis.
    private static uint?[] ReadImageExportLayerAxisIds(OfficeXmlNode layer) {
        var ids = new List<uint?>(3);
        foreach (OfficeXmlNode item in layer.Children(C, "axId")) {
            if (ids.Count == 3) break;
            ids.Add(item.UInt32Val());
        }
        return ids.ToArray();
    }

    // The series data the axis plots: x values on a scatter or bubble horizontal axis, categories on a category or date
    // axis, and values otherwise.
    private static OfficeXmlNode GetSourceLinkedAxisSource(OfficeXmlNode series, OfficeXmlNode axis, bool horizontal) {
        if (horizontal) return series.Child(C, "xVal");
        if (axis.Is(C, "catAx") || axis.Is(C, "dateAx")) return series.Child(C, "cat");
        OfficeXmlNode values = series.Child(C, "val");
        return values.Exists ? values : series.Child(C, "yVal");
    }

    private string? ResolveSourceLinkedAxisNumberFormat(OfficeXmlNode axis) {
        OfficeXmlNode plot = axis.Parent;
        if (!plot.Is(C, "plotArea")) return null;
        uint? axisId = axis.Child(C, "axId").UInt32Val();
        if (!axisId.HasValue) return null;
        string? format = null;
        long remaining = 100_000;
        int remainingStyleRecords = 100_000;
        var sheetStyles = new Dictionary<ExcelSheet, SourceLinkedSheetStyles>();
        string?[]? numberFormats = ReadBoundedSourceLinkedStyles();
        if (numberFormats == null) return null;
        foreach (OfficeXmlNode layer in plot.Children()) {
            uint?[] axisIds = ReadImageExportLayerAxisIds(layer);
            if (!axisIds.Contains(axisId)) continue;
            bool numericHorizontal = (layer.Is(C, "scatterChart") || layer.Is(C, "bubbleChart")) && axisIds.FirstOrDefault() == axisId;
            foreach (OfficeXmlNode series in layer.Children()) {
                if (series.LocalName != "ser") continue;
                if (--remaining < 0) return null;
                OfficeXmlNode reference = GetSourceLinkedAxisSource(series, axis, numericHorizontal).Child(C, "numRef");
                if (reference.IsNone || !ExcelChartUtils.TryParseSheetQualifiedRange(reference.ChildText(C, "f"), out string sheetName, out string range) ||
                    !ExcelReference.TryParse(range, out ExcelReference? address) || address == null ||
                    address.Start.Row <= 0 || address.Start.Column <= 0 || address.End.Row <= 0 || address.End.Column <= 0 ||
                    address.End.Row > 1048576 || address.End.Column > 16384) return null;
                long count = (long)(address.End.Row - address.Start.Row + 1) * (address.End.Column - address.Start.Column + 1);
                if (count <= 0 || count > remaining) return null;
                remaining -= count;
                ExcelSheet sheet;
                try { sheet = _document[sheetName]; } catch (ArgumentException) { return null; }
                if (!sheetStyles.TryGetValue(sheet, out var styles)) {
                    styles = ReadSourceLinkedSheetStyles(sheet, ref remainingStyleRecords);
                    if (styles == null) return null;
                    sheetStyles.Add(sheet, styles);
                }
                for (int row = address.Start.Row; row <= address.End.Row; row++) {
                    for (int column = address.Start.Column; column <= address.End.Column; column++) {
                        uint styleIndex = styles.Resolve(row, column, ExcelA1.CellReference(row, column));
                        if (styleIndex >= numberFormats.Length) return null;
                        string? cellFormat = numberFormats[styleIndex];
                        if (cellFormat == null) return null;
                        if (format != null && !string.Equals(format, cellFormat, StringComparison.Ordinal)) return null;
                        format = cellFormat;
                    }
                }
            }
        }
        return format;
    }
}
