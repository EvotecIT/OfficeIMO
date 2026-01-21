using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Ensures all worksheet elements are in the correct order according to OpenXML schema.
        /// This must be called before saving to prevent validation errors.
        /// </summary>
        internal void EnsureWorksheetElementOrder() {
            var worksheet = _worksheetPart.Worksheet;
            const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

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
                typeof(DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges),
                typeof(Scenarios),
                typeof(AutoFilter),
                typeof(SortState),
                typeof(DataConsolidate),
                typeof(CustomSheetViews),
                typeof(MergeCells),
                typeof(PhoneticProperties),
                typeof(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting),
                typeof(DocumentFormat.OpenXml.Spreadsheet.DataValidations),
                typeof(DocumentFormat.OpenXml.Office2010.Excel.SparklineGroups),
                typeof(Hyperlinks),
                typeof(PrintOptions),
                typeof(PageMargins),
                typeof(PageSetup),
                typeof(HeaderFooter),
                typeof(RowBreaks),
                typeof(ColumnBreaks),
                typeof(CustomProperties),
                typeof(CellWatches),
                typeof(DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors),
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
                typeof(DocumentFormat.OpenXml.Spreadsheet.ExtensionList)
            };
            int pivotTablePartsIndex = elementOrder.IndexOf(typeof(TableParts)) + 1;

            // Snapshot children once and bucket by type for O(n)
            var children = worksheet.ChildElements.ToList();
            var buckets = new Dictionary<System.Type, List<OpenXmlElement>>();
            var pivotTableParts = new List<OpenXmlElement>();
            foreach (var child in children) {
                if (IsPivotTableParts(child, SpreadsheetNamespace)) {
                    pivotTableParts.Add(child);
                    continue;
                }
                var t = child.GetType();
                if (!buckets.TryGetValue(t, out var list)) {
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
            foreach (var child in children) {
                int idx;
                if (IsPivotTableParts(child, SpreadsheetNamespace)) {
                    idx = pivotTablePartsIndex;
                } else {
                    var t = child.GetType();
                    idx = orderIndex.TryGetValue(t, out var val) ? val : unknownIndexBase;
                }
                if (idx < last) {
                    needsReorder = true;
                    break;
                }
                last = idx;
            }

            if (!needsReorder)
                return;

            // Remove all children and rebuild once in the correct order.
            // Build a single ordered buffer to minimize per-append overhead for large worksheets.
            worksheet.RemoveAllChildren();

            var ordered = new List<OpenXmlElement>(children.Count);
            var knownTypes = new HashSet<System.Type>(elementOrder);

            // Known types in schema order
            foreach (var elementType in elementOrder) {
                if (buckets.TryGetValue(elementType, out var list)) {
                    ordered.AddRange(list);
                }
                if (elementType == typeof(TableParts) && pivotTableParts.Count > 0) {
                    ordered.AddRange(pivotTableParts);
                }
            }

            // Unknown types in original order
            foreach (var child in children) {
                if (!knownTypes.Contains(child.GetType()) && !IsPivotTableParts(child, SpreadsheetNamespace)) {
                    ordered.Add(child);
                }
            }

            // Single append is measurably faster than per-item appends on large sets
            if (ordered.Count > 0) {
                worksheet.Append(ordered.ToArray());
            }

            // Persist any structural changes
            worksheet.Save();
        }

        private static bool IsPivotTableParts(OpenXmlElement element, string mainNamespace) {
            if (element is not OpenXmlUnknownElement unknown) return false;
            return unknown.LocalName == "pivotTableParts" && unknown.NamespaceUri == mainNamespace;
        }
    }
}
