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

            // Snapshot children once and bucket by type for O(n)
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

            // Fast-path: if already in non-decreasing schema order, skip reordering.
            // Unknown types are treated as coming after all known schema-ordered types.
            var orderIndex = new Dictionary<System.Type, int>(elementOrder.Count);
            for (int i = 0; i < elementOrder.Count; i++)
                orderIndex[elementOrder[i]] = i;

            bool needsReorder = false;
            int last = -1;
            int unknownIndexBase = elementOrder.Count; // unknowns come after knowns
            foreach (var child in children)
            {
                var t = child.GetType();
                int idx = orderIndex.TryGetValue(t, out var val) ? val : unknownIndexBase;
                if (idx < last)
                {
                    needsReorder = true;
                    break;
                }
                last = idx;
            }

            if (!needsReorder)
                return;

            // Remove all children and append back in schema order (O(n)).
            worksheet.RemoveAllChildren();
            var knownTypes = new HashSet<System.Type>(elementOrder);

            foreach (var elementType in elementOrder)
            {
                if (buckets.TryGetValue(elementType, out var list))
                {
                    foreach (var element in list)
                        worksheet.AppendChild(element);
                }
            }

            // Append any remaining (unknown) elements preserving their original discovery order
            foreach (var child in children)
            {
                if (!knownTypes.Contains(child.GetType()))
                {
                    worksheet.AppendChild(child);
                }
            }

            // Persist any structural changes
            worksheet.Save();
        }
    }
}
