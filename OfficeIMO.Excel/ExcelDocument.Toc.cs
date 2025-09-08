using System;
using System.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Adds or refreshes a simple Table of Contents sheet with hyperlinks.
        /// </summary>
        /// <param name="sheetName">TOC sheet name.</param>
        /// <param name="placeFirst">Move TOC as the first sheet.</param>
        /// <param name="withHyperlinks">Create internal hyperlinks.</param>
        /// <param name="includeNamedRanges">Also list defined names under each sheet.</param>
        /// <param name="includeHiddenNamedRanges">When listing defined names, include ones marked Hidden.</param>
        /// <param name="rangeNameFilter">Optional filter to include only matching defined names.</param>
        public void AddTableOfContents(string sheetName = "TOC", bool placeFirst = true, bool withHyperlinks = true, bool includeNamedRanges = false, bool includeHiddenNamedRanges = false, Predicate<string>? rangeNameFilter = null)
        {
            // Remove existing TOC sheet if present (recreate clean to avoid leftover elements)
            var existing = this.Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                RemoveWorkSheet(sheetName);
            }

            var toc = this.AddWorkSheet(sheetName);
            int r = 1;
            toc.Cell(r, 1, "Table of Contents"); toc.CellBold(r++, 1, true);

            // Build a lookup of defined names to their metadata (for Hidden flag)
            var dnRoot = _workBookPart.Workbook.DefinedNames;
            var dnMeta = dnRoot?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>()
                                 .Where(d => d.Name != null)
                                 .ToDictionary(d => d.Name!.Value, d => d, StringComparer.OrdinalIgnoreCase)
                         ?? new System.Collections.Generic.Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.DefinedName>(StringComparer.OrdinalIgnoreCase);

            foreach (var sh in this.Sheets)
            {
                if (withHyperlinks)
                    toc.SetInternalLink(r, 1, $"'{sh.Name}'!A1", sh.Name);
                else
                    toc.Cell(r, 1, sh.Name);
                r++;

                if (includeNamedRanges)
                {
                    var names = GetAllNamedRanges(sh);
                    foreach (var kv in names)
                    {
                        var name = kv.Key;
                        if (rangeNameFilter != null && !rangeNameFilter(name)) continue;
                        if (dnMeta.TryGetValue(name, out var dn) && dn.Hidden == true && !includeHiddenNamedRanges) continue;

                        var start = kv.Value.Contains(":") ? kv.Value.Split(':')[0] : kv.Value;
                        var text = $"↳ {name}";
                        if (withHyperlinks)
                            toc.SetInternalLink(r, 1, $"'{sh.Name}'!{start}", text);
                        else
                            toc.Cell(r, 1, text);
                        r++;
                    }
                    r++;
                }
            }
            if (placeFirst) MoveSheetToBeginning(sheetName);
        }

        /// <summary>
        /// Backward-compatible alias for AddTableOfContents.
        /// </summary>
        public void CreateTableOfContents(string sheetName = "TOC") => AddTableOfContents(sheetName, placeFirst: true, withHyperlinks: true);

        private void MoveSheetToBeginning(string sheetName)
        {
            var sheets = _workBookPart.Workbook.Sheets;
            if (sheets == null) return;
            var sheet = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                               .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.Ordinal));
            if (sheet == null) return;
            sheet.Remove();
            sheets.InsertAt(sheet, 0);
            _workBookPart.Workbook.Save();
        }

        /// <summary>
        /// Removes a worksheet by name, deleting its part and entry in the workbook.
        /// </summary>
        public void RemoveWorkSheet(string sheetName)
        {
            var wb = _workBookPart.Workbook;
            var sheets = wb.Sheets;
            if (sheets == null) return;
            var sheet = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                              .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.Ordinal));
            if (sheet == null) return;
            var relId = sheet.Id?.Value;
            sheet.Remove();
            if (!string.IsNullOrEmpty(relId))
            {
                var part = (DocumentFormat.OpenXml.Packaging.WorksheetPart)_workBookPart.GetPartById(relId);
                // Remove table parts to avoid orphan relationships
                foreach (var t in part.TableDefinitionParts.ToList())
                {
                    part.DeletePart(t);
                }
                _workBookPart.DeletePart(part);
            }
            wb.Save();
        }
    }
}
