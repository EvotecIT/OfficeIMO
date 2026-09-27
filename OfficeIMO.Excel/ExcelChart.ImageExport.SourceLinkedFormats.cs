using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Excel.Utilities;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private static bool HasSourceLinkedNumericAxisReference(C.PlotArea plot, OpenXmlCompositeElement axis) {
        uint? id = axis.GetFirstChild<C.AxisId>()?.Val?.Value;
        if (!id.HasValue) return false;
        foreach (var layer in plot.ChildElements.OfType<OpenXmlCompositeElement>()) {
            var ids = layer.Elements<C.AxisId>().Take(3).Select(item => item.Val?.Value).ToArray();
            if (!ids.Contains(id)) continue;
            bool horizontal = (layer is C.ScatterChart || layer is C.BubbleChart) && ids.FirstOrDefault() == id;
            foreach (var series in layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser")) {
                OpenXmlElement? source = horizontal ? series.GetFirstChild<C.XValues>() :
                    axis is C.CategoryAxis || axis is C.DateAxis ? series.GetFirstChild<C.CategoryAxisData>() :
                    (OpenXmlElement?)series.GetFirstChild<C.Values>() ?? series.GetFirstChild<C.YValues>();
                if (source?.GetFirstChild<C.NumberReference>() != null) return true;
            }
        }
        return false;
    }

    private string? ResolveSourceLinkedAxisNumberFormat(OpenXmlCompositeElement axis) {
        if (axis.Parent is not C.PlotArea plot) return null;
        uint? axisId = axis.GetFirstChild<C.AxisId>()?.Val?.Value;
        if (!axisId.HasValue) return null;
        string? format = null;
        long remaining = 100_000;
        int remainingStyleRecords = 100_000;
        var sheetStyles = new Dictionary<ExcelSheet, SourceLinkedSheetStyles>();
        string?[]? numberFormats = ReadBoundedSourceLinkedStyles();
        if (numberFormats == null) return null;
        foreach (var layer in plot.ChildElements.OfType<OpenXmlCompositeElement>()) {
            var axisIds = layer.Elements<C.AxisId>().Take(3).Select(item => item.Val?.Value).ToArray();
            if (!axisIds.Contains(axisId)) continue;
            bool numericHorizontal = (layer is C.ScatterChart || layer is C.BubbleChart) && axisIds.FirstOrDefault() == axisId;
            foreach (var series in layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser")) {
                if (--remaining < 0) return null;
                OpenXmlElement? source = numericHorizontal ? series.GetFirstChild<C.XValues>() :
                    axis is C.CategoryAxis || axis is C.DateAxis ? series.GetFirstChild<C.CategoryAxisData>() :
                    (OpenXmlElement?)series.GetFirstChild<C.Values>() ?? series.GetFirstChild<C.YValues>();
                var reference = source?.GetFirstChild<C.NumberReference>();
                if (reference == null || !ExcelChartUtils.TryParseSheetQualifiedRange(reference.Formula?.Text, out string sheetName, out string range) ||
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
                        uint styleIndex = styles.Resolve(row, column, A1.CellReference(row, column));
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
