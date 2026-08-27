using OfficeIMO.Excel;

namespace OfficeIMO.Markup.Excel;

internal sealed partial class OfficeMarkupExcelExporter {
    private static void ApplyTextFormatting(
        ExcelSheet sheet,
        int row,
        int column,
        IDictionary<string, string> attributes) {
        var fontColor = GetAttribute(attributes,
            "color", "font-color", "fontColor", "text-color", "textColor", "textcolor");
        if (!string.IsNullOrWhiteSpace(fontColor)) {
            sheet.CellFontColor(row, column, fontColor!);
        }

        var fontName = GetAttribute(attributes, "font", "font-name", "fontName", "font-family", "fontFamily");
        if (!string.IsNullOrWhiteSpace(fontName)) {
            sheet.CellFontName(row, column, fontName!);
        }

        var fontSize = GetAttribute(attributes, "font-size", "fontSize", "fontsize", "size");
        if (double.TryParse((fontSize ?? string.Empty).Replace("pt", string.Empty), NumberStyles.Float,
            CultureInfo.InvariantCulture, out double parsedFontSize) && parsedFontSize > 0D) {
            sheet.CellFontSize(row, column, parsedFontSize);
        }

        var bold = GetAttribute(attributes, "bold");
        if (TryParseBoolean(bold, out bool parsedBold)) {
            sheet.CellBold(row, column, parsedBold);
        }

        var italic = GetAttribute(attributes, "italic");
        if (TryParseBoolean(italic, out bool parsedItalic)) {
            sheet.CellItalic(row, column, parsedItalic);
        }

        var underline = GetAttribute(attributes, "underline", "underline-style", "underlineStyle");
        if (TryParseUnderline(underline, out ExcelUnderlineStyle underlineStyle)) {
            sheet.CellUnderline(row, column, underlineStyle);
        }

        var strike = GetAttribute(attributes, "strike", "strikethrough", "strike-through");
        if (TryParseBoolean(strike, out bool parsedStrike)) {
            sheet.CellStrikethrough(row, column, parsedStrike);
        }

        var baseline = GetAttribute(attributes, "baseline", "script");
        if (TryParseTextBaseline(baseline, out ExcelVerticalTextAlignment textBaseline)) {
            sheet.CellVerticalTextAlignment(row, column, textBaseline);
        }

        var textCase = GetAttribute(attributes, "text-case", "textCase", "case");
        if (TryParseTextCase(textCase, out OfficeTextCase parsedTextCase)) {
            sheet.TransformCellTextCase(row, column, parsedTextCase, CultureInfo.InvariantCulture);
        }
    }

    private static bool TryParseBoolean(string? value, out bool result) {
        if (string.IsNullOrWhiteSpace(value)) {
            result = false;
            return false;
        }

        switch (Normalize(value!)) {
            case "true": case "yes": case "on": case "1": result = true; return true;
            case "false": case "no": case "off": case "0": result = false; return true;
            default: result = false; return false;
        }
    }

    private static bool TryParseUnderline(string? value, out ExcelUnderlineStyle style) {
        if (TryParseBoolean(value, out bool enabled)) {
            style = enabled ? ExcelUnderlineStyle.Single : ExcelUnderlineStyle.None;
            return true;
        }

        switch (Normalize(value ?? string.Empty)) {
            case "single": case "solid": style = ExcelUnderlineStyle.Single; return true;
            case "double": style = ExcelUnderlineStyle.Double; return true;
            case "singleaccounting": case "accounting": style = ExcelUnderlineStyle.SingleAccounting; return true;
            case "doubleaccounting": style = ExcelUnderlineStyle.DoubleAccounting; return true;
            case "none": style = ExcelUnderlineStyle.None; return true;
            default: style = ExcelUnderlineStyle.None; return false;
        }
    }

    private static bool TryParseTextBaseline(string? value, out ExcelVerticalTextAlignment alignment) {
        switch (Normalize(value ?? string.Empty)) {
            case "normal": case "baseline": case "none": alignment = ExcelVerticalTextAlignment.Baseline; return true;
            case "sup": case "super": case "superscript": alignment = ExcelVerticalTextAlignment.Superscript; return true;
            case "sub": case "subscript": alignment = ExcelVerticalTextAlignment.Subscript; return true;
            default: alignment = ExcelVerticalTextAlignment.Baseline; return false;
        }
    }

    private static bool TryParseTextCase(string? value, out OfficeTextCase textCase) {
        switch (Normalize(value ?? string.Empty)) {
            case "none": case "preserve": textCase = OfficeTextCase.None; return true;
            case "upper": case "uppercase": textCase = OfficeTextCase.Uppercase; return true;
            case "lower": case "lowercase": textCase = OfficeTextCase.Lowercase; return true;
            case "title": case "titlecase": textCase = OfficeTextCase.TitleCase; return true;
            case "sentence": case "sentencecase": textCase = OfficeTextCase.SentenceCase; return true;
            case "toggle": case "togglecase": textCase = OfficeTextCase.ToggleCase; return true;
            case "capitalize": case "capitalise": textCase = OfficeTextCase.Capitalize; return true;
            default: textCase = OfficeTextCase.None; return false;
        }
    }
}
