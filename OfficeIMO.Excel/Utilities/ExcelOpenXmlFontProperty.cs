using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

internal static class ExcelOpenXmlFontProperty {
    internal static bool IsEnabled(BooleanPropertyType? property) =>
        property != null && (property.Val?.Value ?? true);

    internal static bool IsUnderlineEnabled(Underline? underline) =>
        underline != null && (underline.Val?.Value ?? UnderlineValues.Single) != UnderlineValues.None;
}
