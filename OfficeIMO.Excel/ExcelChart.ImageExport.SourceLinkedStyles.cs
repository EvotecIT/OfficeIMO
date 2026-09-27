using System;
using System.Collections.Generic;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private string[]? ReadBoundedSourceLinkedStyles() {
        var stylesheet = _document.WorkbookPartRoot?.WorkbookStylesPart?.Stylesheet;
        if (stylesheet == null) return new[] { "General" };
        int remaining = 100_000;
        var custom = new Dictionary<uint, string>();
        var styles = new List<string>();
        try {
            foreach (var item in stylesheet.NumberingFormats?.Elements<S.NumberingFormat>() ?? Enumerable.Empty<S.NumberingFormat>()) {
                if (--remaining < 0) return null;
                if (item.NumberFormatId?.Value is not uint id || item.FormatCode?.Value is not string code || custom.ContainsKey(id)) return null;
                custom.Add(id, code);
            }
            foreach (var style in stylesheet.CellFormats?.Elements<S.CellFormat>() ?? Enumerable.Empty<S.CellFormat>()) {
                if (--remaining < 0) return null;
                uint id = style.NumberFormatId?.Value ?? 0;
                styles.Add(custom.TryGetValue(id, out string? code) ? code : ExcelBuiltInNumberFormats.GetCode(id) ?? "General");
            }
        } catch (FormatException) { return null; }
        return styles.Count == 0 ? new[] { "General" } : styles.ToArray();
    }
}
