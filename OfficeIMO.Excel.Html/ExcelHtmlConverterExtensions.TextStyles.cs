using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    private static bool AppendExcelTextStyleAttributes(
        StringBuilder body,
        ExcelCellStyleSnapshot style,
        bool includeDecorations) {
        var css = new StringBuilder();
        AppendCss(css, "font-weight", style.Bold ? "700" : null);
        AppendCss(css, "font-style", style.Italic ? "italic" : null);
        AppendCss(css, "font-family", style.IsFontFamilyExplicit && !string.IsNullOrWhiteSpace(style.FontName)
            ? OfficeHtmlText.QuoteCssString(style.FontName!) : null);
        AppendCss(css, "font-size", style.FontSize.HasValue
            ? style.FontSize.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt" : null);
        AppendCss(css, "color", !string.IsNullOrWhiteSpace(style.FontColorHex) ? "#" + style.FontColorHex : null);
        ExcelUnderlineStyle underlineStyle = style.Underline
            ? style.UnderlineStyle ?? ExcelUnderlineStyle.Single
            : ExcelUnderlineStyle.None;
        bool splitDecorations = includeDecorations && RequiresIndependentExcelDecorations(underlineStyle, style.Strikethrough);
        if (includeDecorations) {
            AppendExcelDecorations(css, underlineStyle, style.Strikethrough, splitDecorations: splitDecorations);
        }
        AppendStyleAttribute(body, css);
        if (style.UnderlineStyle.HasValue && style.UnderlineStyle.Value != ExcelUnderlineStyle.None) {
            body.Append(" data-officeimo-excel-underline=\"")
                .Append(OfficeHtmlText.EscapeAttribute(style.UnderlineStyle.Value.ToString()))
                .Append('"');
        }
        if (style.VerticalTextAlignment.HasValue) {
            body.Append(" data-officeimo-excel-vertical-align=\"")
                .Append(OfficeHtmlText.EscapeAttribute(style.VerticalTextAlignment.Value.ToString()))
                .Append('"');
        }
        if (splitDecorations) {
            body.Append(" data-officeimo-excel-strikethrough=\"true\" data-officeimo-excel-decoration-split=\"true\"");
        }
        return splitDecorations;
    }

    private static bool AppendExcelTextStyleAttributes(
        StringBuilder body,
        ExcelRichTextRun run,
        ExcelCellStyleSnapshot cellStyle) {
        var css = new StringBuilder();
        AppendCss(css, "font-weight", run.BoldSpecified ? (run.Bold ? "700" : "normal") : null);
        AppendCss(css, "font-style", run.ItalicSpecified ? (run.Italic ? "italic" : "normal") : null);
        AppendCss(css, "font-family", !string.IsNullOrWhiteSpace(run.FontName) ? OfficeHtmlText.QuoteCssString(run.FontName!) : null);
        AppendCss(css, "font-size", run.FontSize.HasValue
            ? run.FontSize.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt" : null);
        string color = (run.FontColor ?? string.Empty).Trim().TrimStart('#');
        if (color.Length == 8) color = color.Substring(2);
        AppendCss(css, "color", color.Length == 6 ? "#" + color : null);
        ExcelUnderlineStyle underlineStyle = run.UnderlineSpecified
            ? run.UnderlineStyle ?? (run.Underline ? ExcelUnderlineStyle.Single : ExcelUnderlineStyle.None)
            : cellStyle.Underline ? cellStyle.UnderlineStyle ?? ExcelUnderlineStyle.Single : ExcelUnderlineStyle.None;
        bool strikethrough = run.StrikethroughSpecified ? run.Strikethrough : cellStyle.Strikethrough;
        bool splitDecorations = RequiresIndependentExcelDecorations(underlineStyle, strikethrough);
        AppendExcelDecorations(css, underlineStyle, strikethrough, emitNone: true, splitDecorations: splitDecorations);
        ExcelVerticalTextAlignment? verticalTextAlignment = run.VerticalTextAlignment ?? cellStyle.VerticalTextAlignment;
        AppendCss(css, "vertical-align", verticalTextAlignment switch {
            ExcelVerticalTextAlignment.Superscript => "super",
            ExcelVerticalTextAlignment.Subscript => "sub",
            ExcelVerticalTextAlignment.Baseline => "baseline",
            _ => null
        });
        AppendStyleAttribute(body, css);
        if (underlineStyle != ExcelUnderlineStyle.None || run.UnderlineSpecified) {
            body.Append(" data-officeimo-excel-underline=\"")
                .Append(OfficeHtmlText.EscapeAttribute(underlineStyle.ToString()))
                .Append('"');
        }
        if (splitDecorations) body.Append(" data-officeimo-excel-strikethrough=\"true\"");
        return splitDecorations;
    }

    private static void AppendExcelDecorations(
        StringBuilder css,
        ExcelUnderlineStyle underlineStyle,
        bool strike,
        bool emitNone = false,
        bool splitDecorations = false) {
        var lines = new List<string>(2);
        if (underlineStyle != ExcelUnderlineStyle.None) lines.Add("underline");
        if (strike && !splitDecorations) lines.Add("line-through");
        if (lines.Count == 0) {
            if (emitNone) AppendCss(css, "text-decoration-line", "none");
            return;
        }
        AppendCss(css, "text-decoration-line", string.Join(" ", lines));
        AppendCss(css, "text-decoration-style", underlineStyle is ExcelUnderlineStyle.Double or ExcelUnderlineStyle.DoubleAccounting
            ? "double" : "solid");
    }

    private static bool RequiresIndependentExcelDecorations(ExcelUnderlineStyle underlineStyle, bool strike) =>
        strike && underlineStyle is ExcelUnderlineStyle.Double or ExcelUnderlineStyle.DoubleAccounting;

    private static void AppendIndependentExcelStrike(StringBuilder body, string text) {
        body.Append("<span style=\"text-decoration-line:line-through;text-decoration-style:solid\">")
            .Append(OfficeHtmlText.Escape(text))
            .Append("</span>");
    }

    private static void AppendCss(StringBuilder css, string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        if (css.Length > 0) css.Append(';');
        css.Append(name).Append(':').Append(value);
    }

    private static void AppendStyleAttribute(StringBuilder body, StringBuilder css) {
        if (css.Length == 0) return;
        body.Append(" style=\"")
            .Append(OfficeHtmlText.EscapeAttribute(css.ToString()))
            .Append('"');
    }

}
