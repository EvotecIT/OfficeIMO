using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel
{
    public partial class ExcelSheet
    {
        /// <summary>
        /// Ensures all worksheet elements are in the correct order according to OpenXML schema.
        /// This must be called before saving to prevent validation errors.
        /// </summary>
        internal void EnsureWorksheetElementOrder()
        {
            var worksheet = _worksheetPart.Worksheet;
            
            // Define the correct order of elements according to OpenXML schema
            var elementOrder = new List<System.Type>
            {
                typeof(SheetProperties),
                typeof(SheetDimension),
                typeof(SheetViews),
                typeof(SheetFormatProperties),
                typeof(Columns),
                typeof(SheetData),
                typeof(SheetCalculationProperties),
                typeof(SheetProtection),
                typeof(ProtectedRanges),
                typeof(Scenarios),
                typeof(AutoFilter),
                typeof(SortState),
                typeof(DataConsolidate),
                typeof(CustomSheetViews),
                typeof(MergeCells),
                typeof(PhoneticProperties),
                typeof(ConditionalFormatting),
                typeof(DataValidations),
                typeof(Hyperlinks),
                typeof(PrintOptions),
                typeof(PageMargins),
                typeof(PageSetup),
                typeof(HeaderFooter),
                typeof(RowBreaks),
                typeof(ColumnBreaks),
                typeof(CustomProperties),
                typeof(CellWatches),
                typeof(IgnoredErrors),
                // SmartTags is deprecated and not in current OpenXML SDK
                typeof(Drawing),
                typeof(LegacyDrawing),
                typeof(LegacyDrawingHeaderFooter),
                typeof(DrawingHeaderFooter),
                typeof(Picture),
                typeof(OleObjects),
                typeof(Controls),
                typeof(WebPublishItems),
                typeof(TableParts),
                typeof(ExtensionList)
            };

            // Get all current children
            var children = worksheet.ChildElements.ToList();
            
            // Remove all children
            worksheet.RemoveAllChildren();
            
            // Re-add children in correct order
            foreach (var elementType in elementOrder)
            {
                var elementsOfType = children.Where(c => c.GetType() == elementType).ToList();
                foreach (var element in elementsOfType)
                {
                    worksheet.AppendChild(element);
                }
            }
            
            // Add any remaining elements that weren't in our list (for forward compatibility)
            var remainingElements = children.Where(c => !elementOrder.Contains(c.GetType())).ToList();
            foreach (var element in remainingElements)
            {
                worksheet.AppendChild(element);
            }
        }
    }
}