using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>The worksheet children that the native XLSB writers encode as BIFF12 records.</summary>
    internal static class XlsbWorksheetContent {
        private static readonly HashSet<Type> Supported = new HashSet<Type> {
            typeof(SheetProperties),
            typeof(SheetDimension),
            typeof(SheetViews),
            typeof(SheetFormatProperties),
            typeof(Columns),
            typeof(SheetData),
            typeof(SheetProtection),
            typeof(AutoFilter),
            typeof(MergeCells),
            typeof(ConditionalFormatting),
            typeof(DataValidations),
            typeof(Hyperlinks),
            typeof(PrintOptions),
            typeof(PageMargins),
            typeof(PageSetup),
            typeof(HeaderFooter),
            typeof(RowBreaks),
            typeof(ColumnBreaks),
            typeof(IgnoredErrors),
            typeof(LegacyDrawing),
        };

        // CT_Worksheet allows these children more than once; every other supported child occurs at most once.
        private static readonly HashSet<Type> Repeatable = new HashSet<Type> {
            typeof(ConditionalFormatting),
        };

        internal static bool IsSupported(OpenXmlElement element) => Supported.Contains(element.GetType());

        internal static bool IsRepeatable(OpenXmlElement element) => Repeatable.Contains(element.GetType());

        internal static OpenXmlElement? FindUnsupported(Worksheet? worksheet) =>
            worksheet?.ChildElements.FirstOrDefault(element => !IsSupported(element));
    }
}
