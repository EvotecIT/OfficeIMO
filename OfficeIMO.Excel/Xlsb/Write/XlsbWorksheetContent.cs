using DocumentFormat.OpenXml.Packaging;
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
            typeof(TableParts),
            typeof(DocumentFormat.OpenXml.Spreadsheet.Drawing),
        };

        // CT_Worksheet allows these children more than once; every other supported child occurs at most once.
        private static readonly HashSet<Type> Repeatable = new HashSet<Type> {
            typeof(ConditionalFormatting),
        };

        internal static bool IsSupported(OpenXmlElement element) => Supported.Contains(element.GetType());

        internal static bool IsRepeatable(OpenXmlElement element) => Repeatable.Contains(element.GetType());

        internal static string? FindUnsupported(ExcelSheet sheet) =>
            sheet.WorksheetPart.Worksheet?.ChildElements.FirstOrDefault(element => !IsSupported(element))?.LocalName;

        internal static string? FindUnsupportedPart(ExcelSheet sheet, string externalContent) {
            WorksheetPart part = sheet.WorksheetPart;
            return FindUnsupportedPart(part) ?? (part.ExternalRelationships.Any() ? externalContent : null);
        }

        /// <summary>
        /// Returns a message naming the first worksheet-related part the native writers cannot carry, or null. Comments with
        /// the VML drawing that LegacyDrawing names, table definitions and drawings are the related parts accepted.
        /// </summary>
        internal static string? FindUnsupportedPart(WorksheetPart worksheetPart) {
            string? legacyId = worksheetPart.Worksheet?.GetFirstChild<LegacyDrawing>()?.Id?.Value;
            foreach (IdPartPair pair in worksheetPart.Parts) {
                switch (pair.OpenXmlPart) {
                    case WorksheetCommentsPart:
                    case TableDefinitionPart:
                    case DrawingsPart:
                        continue;
                    case VmlDrawingPart when legacyId != null && pair.RelationshipId == legacyId
                        && worksheetPart.WorksheetCommentsPart != null:
                        continue;
                    case WorksheetThreadedCommentsPart:
                        return "threaded comments; legacy notes (SetComment) are supported";
                    default:
                        return "relationship-backed content";
                }
            }
            return null;
        }
    }
}
