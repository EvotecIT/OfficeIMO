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

            // Get all current children once and bucket by type for O(n)
            var children = worksheet.ChildElements.ToList();
            var buckets = new Dictionary<System.Type, List<OpenXmlElement>>();
            foreach (var child in children)
            {
                var t = child.GetType();
                if (!buckets.TryGetValue(t, out var list))
                {
                    list = new List<OpenXmlElement>();
                    buckets[t] = list;
                }
                list.Add(child);
            }

            // Remove all children and append back in schema order
            worksheet.RemoveAllChildren();
            foreach (var elementType in elementOrder)
            {
                if (buckets.TryGetValue(elementType, out var list))
                {
                    foreach (var element in list)
                        worksheet.AppendChild(element);
                    buckets.Remove(elementType);
                }
            }

            // Append any remaining elements in their original discovery order
            foreach (var kv in buckets)
            {
                foreach (var element in kv.Value)
                    worksheet.AppendChild(element);
            }
        }
    }
}
