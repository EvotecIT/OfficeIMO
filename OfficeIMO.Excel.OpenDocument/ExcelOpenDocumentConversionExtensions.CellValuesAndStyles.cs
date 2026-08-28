using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private enum ExcelValueProjectionStatus {
        Exact,
        Invalid,
        TimeZoneNormalized
    }

    private static bool SetOdsValue(OdsCell target, object? value) {
        if (value == null) return true;
        if (value is string text) target.SetString(text);
        else if (value is bool boolean) target.SetBoolean(boolean);
        else if (value is decimal decimalValue) target.SetDecimal(decimalValue);
        else if (value is DateTime dateTime) target.SetDate(dateTime);
        else if (value is DateTimeOffset dateTimeOffset) target.SetDateTime(dateTimeOffset);
        else if (value is TimeSpan timeSpan) target.SetDuration(timeSpan);
        else if (IsNumeric(value)) target.SetNumber(Convert.ToDouble(value, CultureInfo.InvariantCulture));
        else { target.SetString(Convert.ToString(value, CultureInfo.InvariantCulture)); return false; }
        return true;
    }

    private static ExcelValueProjectionStatus SetExcelValue(ExcelCell target, OdsCellValue value) {
        try {
            switch (value.Kind) {
                case OdsCellValueKind.Empty: return ExcelValueProjectionStatus.Exact;
                case OdsCellValueKind.String: target.SetValue(value.LexicalValue); return ExcelValueProjectionStatus.Exact;
                case OdsCellValueKind.Number:
                case OdsCellValueKind.Percentage:
                case OdsCellValueKind.Currency: target.SetValue(value.AsDecimal()); return ExcelValueProjectionStatus.Exact;
                case OdsCellValueKind.Boolean: target.SetValue(value.AsBoolean()); return ExcelValueProjectionStatus.Exact;
                case OdsCellValueKind.Date:
                    DateTimeOffset dateTime = value.AsDateTimeOffset();
                    bool hasTimeZone = HasExplicitTimeZone(value.LexicalValue);
                    target.SetValue(hasTimeZone ? dateTime.UtcDateTime : dateTime.DateTime);
                    return hasTimeZone
                        ? ExcelValueProjectionStatus.TimeZoneNormalized
                        : ExcelValueProjectionStatus.Exact;
                case OdsCellValueKind.Time: target.SetValue(value.AsTimeSpan()); return ExcelValueProjectionStatus.Exact;
                default: target.SetValue(value.ToString()); return ExcelValueProjectionStatus.Invalid;
            }
        } catch (FormatException) {
            target.SetValue(value.ToString());
            return ExcelValueProjectionStatus.Invalid;
        } catch (OverflowException) {
            target.SetValue(value.ToString());
            return ExcelValueProjectionStatus.Invalid;
        }
    }

    private static bool HasExplicitTimeZone(string lexical) {
        if (lexical.EndsWith("Z", StringComparison.OrdinalIgnoreCase)) return true;
        int length = lexical.Length;
        return length >= 6 && (lexical[length - 6] == '+' || lexical[length - 6] == '-') &&
            lexical[length - 3] == ':' &&
            char.IsDigit(lexical[length - 5]) && char.IsDigit(lexical[length - 4]) &&
            char.IsDigit(lexical[length - 2]) && char.IsDigit(lexical[length - 1]);
    }

    private static void ApplyExcelStyle(OdsDocument document, OdsCell target, ExcelCellStyleSnapshot style,
        IDictionary<uint, string> dataStyles, ref int unsupported) {
        if (style.Bold) target.Bold = true;
        if (style.Italic) target.Italic = true;
        ExcelUnderlineStyle? underlineStyle = style.UnderlineStyle;
        bool underlined = underlineStyle.HasValue
            ? underlineStyle.Value != ExcelUnderlineStyle.None
            : style.Underline;
        if (underlined) {
            target.Underline = true;
            target.UnderlineStyle = OdfTextDecorationStyle.Solid;
            target.UnderlineType = underlineStyle is ExcelUnderlineStyle.Double or ExcelUnderlineStyle.DoubleAccounting
                ? OdfTextDecorationType.Double
                : OdfTextDecorationType.Single;
            if (underlineStyle is ExcelUnderlineStyle.SingleAccounting or ExcelUnderlineStyle.DoubleAccounting) unsupported++;
        }
        if (style.Strikethrough) {
            target.StrikeThrough = true;
            target.LineThroughStyle = OdfTextDecorationStyle.Solid;
            target.LineThroughType = OdfTextDecorationType.Single;
        }
        target.TextPosition = style.VerticalTextAlignment switch {
            ExcelVerticalTextAlignment.Superscript => OdfTextPosition.Superscript,
            ExcelVerticalTextAlignment.Subscript => OdfTextPosition.Subscript,
            ExcelVerticalTextAlignment.Baseline => OdfTextPosition.Normal,
            _ => null
        };
        if (style.FontSize.HasValue) target.FontSize = OdfLength.Points(style.FontSize.Value);
        if (!string.IsNullOrWhiteSpace(style.FontName)) target.FontFamily = style.FontName;
        if (!string.IsNullOrWhiteSpace(style.FontColorHex)) target.Color = OdfColor.Parse(style.FontColorHex!);
        if (!string.IsNullOrWhiteSpace(style.FillColorHex)) target.BackgroundColor = OdfColor.Parse(style.FillColorHex!);
        if (!string.IsNullOrWhiteSpace(style.NumberFormatCode) && style.NumberFormatCode != "General") {
            if (!dataStyles.TryGetValue(style.StyleIndex, out string? name)) {
                name = "xlData" + style.StyleIndex.ToString(CultureInfo.InvariantCulture);
                SpreadsheetNumberFormatSyntax format = SpreadsheetNumberFormatSyntax.Parse(style.NumberFormatCode!);
                bool timeOnly = IsTimeOnly(format);
                if (style.IsDateLike) {
                    if (timeOnly) document.AddTimeStyle(name); else document.AddDateStyle(name);
                } else if (format.IsPercentage) {
                    document.AddPercentageStyle(name, format.DecimalPlaces, format.UsesGrouping);
                } else if (!string.IsNullOrWhiteSpace(format.CurrencySymbol)) {
                    document.AddCurrencyStyle(name, format.CurrencySymbol!, format.DecimalPlaces, format.UsesGrouping);
                } else {
                    document.AddNumberStyle(name, format.DecimalPlaces, format.UsesGrouping);
                }
                if (HasUnsupportedNumberFormatProjection(style, format, timeOnly)) unsupported++;
                dataStyles.Add(style.StyleIndex, name);
            }
            target.NumberFormatName = name;
        }
        if (HasUnsupportedExcelStyle(style)) unsupported++;
    }

    private static bool IsTimeOnly(SpreadsheetNumberFormatSyntax format) =>
        format.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol &&
                (token.Value.IndexOf("h", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 token.Value.IndexOf("s", StringComparison.OrdinalIgnoreCase) >= 0)) &&
        !format.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol &&
                (token.Value.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 token.Value.IndexOf("d", StringComparison.OrdinalIgnoreCase) >= 0));

    private static bool HasUnsupportedNumberFormatProjection(
        ExcelCellStyleSnapshot style,
        SpreadsheetNumberFormatSyntax format,
        bool timeOnly) {
        if (style.IsDateLike) {
            string projected = timeOnly ? "hh:mm:ss" : "yyyy-mm-dd";
            return !format.IsValid || format.SectionCount > 1 ||
                !string.Equals(style.NumberFormatCode, projected, StringComparison.OrdinalIgnoreCase);
        }
        return !format.IsValid || format.SectionCount > 1 || HasUnsupportedPlaceholderSemantics(format) ||
            format.Tokens.Any(token =>
                token.LocaleCode != null ||
                token.Kind == SpreadsheetNumberFormatTokenKind.BracketedDirective ||
                token.Kind == SpreadsheetNumberFormatTokenKind.ScalingSeparator ||
                token.Kind == SpreadsheetNumberFormatTokenKind.TextPlaceholder ||
                token.Kind == SpreadsheetNumberFormatTokenKind.Literal ||
                token.Kind == SpreadsheetNumberFormatTokenKind.Other);
    }

    private static bool HasUnsupportedPlaceholderSemantics(SpreadsheetNumberFormatSyntax format) {
        bool afterDecimal = false;
        bool hasPlaceholder = false;
        int mandatoryIntegerDigits = 0;
        foreach (SpreadsheetNumberFormatToken token in format.Tokens) {
            if (token.Kind == SpreadsheetNumberFormatTokenKind.SectionSeparator) break;
            if (token.Kind == SpreadsheetNumberFormatTokenKind.DecimalSeparator) {
                afterDecimal = true;
                continue;
            }
            if (token.Kind != SpreadsheetNumberFormatTokenKind.Placeholder) continue;
            hasPlaceholder = true;
            if (token.Text.IndexOf('?') >= 0) return true;
            if (afterDecimal && token.Text.IndexOf('#') >= 0) return true;
            if (!afterDecimal) mandatoryIntegerDigits += token.Text.Count(character => character == '0');
        }
        return hasPlaceholder && mandatoryIntegerDigits != 1;
    }

    private static bool HasUnsupportedExcelStyle(ExcelCellStyleSnapshot style) {
        bool nonSolidPattern = !string.IsNullOrWhiteSpace(style.FillPatternType) &&
            !string.Equals(style.FillPatternType, "none", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(style.FillPatternType, "solid", StringComparison.OrdinalIgnoreCase);
        return style.Border != null ||
            nonSolidPattern || style.FillGradientUnsupported || style.FillGradientStops.Count > 0 ||
            style.TextRotation.HasValue || style.HorizontalAlignment != null || style.VerticalAlignment != null ||
            (style.TextIndent.HasValue && style.TextIndent.Value > 0U) || style.WrapText || style.ShrinkToFit;
    }

    private static int ApplyOdsStyle(
        ExcelCell target,
        OdsCellRun style,
        IReadOnlyDictionary<string, OdsDataStyle> dataStyles,
        out bool unsupportedDataStyleFormat,
        ref int approximatedFontFamilyLists,
        ref int unsupportedFontFamilies,
        CultureInfo textCaseCulture) {
        int unsupported = 0;
        unsupportedDataStyleFormat = false;
        if (style.Bold == true) target.SetBold();
        if (style.Italic == true) target.SetItalic();
        ExcelUnderlineStyle? underline = MapOdfUnderline(
            style.Underline,
            style.UnderlineStyle,
            style.UnderlineType,
            out bool approximatedUnderline);
        if (underline.HasValue) target.SetUnderline(underline.Value);
        if (approximatedUnderline) unsupported++;
        bool hasStrike = style.StrikeThrough == true &&
            style.LineThroughStyle != OdfTextDecorationStyle.None &&
            style.LineThroughType != OdfTextDecorationType.None;
        if (hasStrike) target.SetStrikethrough();
        if (hasStrike && (style.LineThroughStyle is not null and not OdfTextDecorationStyle.Solid ||
                          style.LineThroughType == OdfTextDecorationType.Double)) unsupported++;
        if (style.TextPosition.HasValue) {
            target.SetVerticalTextAlignment(style.TextPosition.Value switch {
                OdfTextPosition.Superscript => ExcelVerticalTextAlignment.Superscript,
                OdfTextPosition.Subscript => ExcelVerticalTextAlignment.Subscript,
                _ => ExcelVerticalTextAlignment.Baseline
            });
        }
        if (style.TextTransform == OdfTextTransform.Uppercase) target.TransformTextCase(OfficeTextCase.Uppercase, textCaseCulture);
        else if (style.TextTransform == OdfTextTransform.Lowercase) target.TransformTextCase(OfficeTextCase.Lowercase, textCaseCulture);
        else if (style.TextTransform == OdfTextTransform.Capitalize) target.TransformTextCase(OfficeTextCase.Capitalize, textCaseCulture);
        if (style.SmallCaps == true) unsupported++;
        if (style.FontSize.HasValue) {
            if (style.FontSize.Value.TryToPoints(out double points)) target.SetFontSize(points);
            else unsupported++;
        }
        string? fontFamily = SelectOdfFontFamily(style.FontFamily,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
        if (fontFamily != null) target.SetFontName(fontFamily);
        if (style.Color.HasValue) target.SetFontColor(style.Color.Value.ToString().TrimStart('#'));
        if (style.BackgroundColor.HasValue) target.SetFillColor(style.BackgroundColor.Value.ToString().TrimStart('#'));
        if (style.NumberFormatName != null && dataStyles.TryGetValue(style.NumberFormatName, out OdsDataStyle? dataStyle)) {
            if (dataStyle.TryGetExcelNumberFormatCode(out string formatCode)) target.SetNumberFormat(formatCode);
            else unsupportedDataStyleFormat = true;
        }
        return unsupported;
    }

    private static ExcelUnderlineStyle? MapOdfUnderline(
        bool? enabled,
        OdfTextDecorationStyle? style,
        OdfTextDecorationType? type,
        out bool approximated) {
        approximated = false;
        if (enabled != true || style == OdfTextDecorationStyle.None || type == OdfTextDecorationType.None) return null;
        if (style is not null and not OdfTextDecorationStyle.Solid) approximated = true;
        return type == OdfTextDecorationType.Double ? ExcelUnderlineStyle.Double : ExcelUnderlineStyle.Single;
    }

    private static string? SelectOdfFontFamily(string? value,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (!OdfFontFamilySyntax.TryParse(value, out OdfFontFamilySyntax? syntax)) {
            unsupportedFontFamilies++;
            return null;
        }
        if (syntax!.HasFallbacks) approximatedFontFamilyLists++;
        return syntax.PrimaryFamily;
    }

    private static bool IsNumeric(object value) {
        TypeCode code = Type.GetTypeCode(value.GetType());
        return code >= TypeCode.SByte && code <= TypeCode.Decimal;
    }

    private static bool HasRichTextFormatting(ExcelRichTextRun run) =>
        run.Bold || run.Italic || run.Underline || run.Strikethrough ||
        (run.UnderlineStyle.HasValue && run.UnderlineStyle.Value != ExcelUnderlineStyle.None) ||
        !string.IsNullOrWhiteSpace(run.FontColor) || !string.IsNullOrWhiteSpace(run.FontName) ||
        run.FontSize.HasValue || run.VerticalTextAlignment.HasValue || run.Outline || run.Shadow ||
        run.Condense || run.Extend || run.FontFamily.HasValue || run.FontCharacterSet.HasValue;
}
