using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private string?[]? ReadBoundedSourceLinkedStyles() {
        var stylesheet = _document.WorkbookPartRoot?.WorkbookStylesPart?.Stylesheet;
        if (stylesheet == null) return new[] { "General" };
        int remaining = 100_000;
        var custom = new Dictionary<uint, string>();
        var styles = new List<string?>();
        var namedStyles = new List<uint>();
        try {
            foreach (var item in stylesheet.NumberingFormats?.Elements<S.NumberingFormat>() ?? Enumerable.Empty<S.NumberingFormat>()) {
                if (--remaining < 0) return null;
                if (!uint.TryParse(item.NumberFormatId?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out uint id) || item.FormatCode?.Value is not string code || custom.ContainsKey(id)) return null;
                custom.Add(id, code);
            }
            foreach (var style in stylesheet.CellStyleFormats?.Elements<S.CellFormat>() ?? Enumerable.Empty<S.CellFormat>()) {
                if (--remaining < 0 || !uint.TryParse(style.NumberFormatId?.InnerText ?? "0", NumberStyles.None, CultureInfo.InvariantCulture, out uint id)) return null;
                namedStyles.Add(id);
            }
            foreach (var style in stylesheet.CellFormats?.Elements<S.CellFormat>() ?? Enumerable.Empty<S.CellFormat>()) {
                if (--remaining < 0) return null;
                uint baseIndex = 0;
                if (style.FormatId != null && !uint.TryParse(style.FormatId.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out baseIndex)) return null;
                if (style.FormatId != null && baseIndex >= namedStyles.Count) return null;
                uint id = style.FormatId != null ? namedStyles[(int)baseIndex] : 0;
                if (style.ApplyNumberFormat?.Value != false && style.NumberFormatId != null &&
                    !uint.TryParse(style.NumberFormatId.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out id)) return null;
                styles.Add(custom.TryGetValue(id, out string? code) ? code : ExcelBuiltInNumberFormats.GetCode(id));
            }
        } catch (FormatException) { return null; }
        return styles.Count == 0 ? new[] { "General" } : styles.ToArray();
    }
}
