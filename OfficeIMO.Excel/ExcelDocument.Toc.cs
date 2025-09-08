using System;
using System.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Adds or refreshes a visually styled Table of Contents sheet with hyperlinks.
        /// </summary>
        /// <param name="sheetName">TOC sheet name.</param>
        /// <param name="placeFirst">Move TOC as the first sheet.</param>
        /// <param name="withHyperlinks">Create internal hyperlinks.</param>
        /// <param name="includeNamedRanges">Also list defined names in a dedicated column.</param>
        /// <param name="includeHiddenNamedRanges">When listing defined names, include ones marked Hidden.</param>
        /// <param name="rangeNameFilter">Optional filter to include only matching defined names.</param>
        /// <param name="styled">When true, renders a banner and a formatted table with AutoFilter.</param>
        public void AddTableOfContents(string sheetName = "TOC", bool placeFirst = true, bool withHyperlinks = true, bool includeNamedRanges = false, bool includeHiddenNamedRanges = false, Predicate<string>? rangeNameFilter = null, bool styled = true)
        {
            // Remove existing TOC sheet if present (recreate clean to avoid leftover elements)
            var existing = this.Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                RemoveWorkSheet(sheetName);
            }

            var toc = this.AddWorkSheet(sheetName);
            int r = 1;
            // Banner
            toc.Cell(r, 1, "Workbook Navigation"); toc.CellBold(r, 1, true); toc.CellBackground(r, 1, "#D9E1F2"); r++;
            toc.Cell(r++, 1, $"Generated: {System.DateTime.Now:yyyy-MM-dd HH:mm}");

            // Build a lookup of defined names to their metadata (for Hidden flag)
            var dnRoot = _workBookPart.Workbook.DefinedNames;
            var dnMeta = dnRoot?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>()
                                 .Where(d => d.Name != null)
                                 .ToDictionary(d => d.Name!.Value, d => d, StringComparer.OrdinalIgnoreCase)
                         ?? new System.Collections.Generic.Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.DefinedName>(StringComparer.OrdinalIgnoreCase);

            // Header
            int headerRow = r;
            toc.Cell(headerRow, 1, "Sheet"); toc.CellBold(headerRow, 1, true); toc.CellBackground(headerRow, 1, "#F2F2F2");
            toc.Cell(headerRow, 2, "Details"); toc.CellBold(headerRow, 2, true); toc.CellBackground(headerRow, 2, "#F2F2F2");
            int colRanges = includeNamedRanges ? 3 : 0;
            if (includeNamedRanges) { toc.Cell(headerRow, 3, "Named Ranges"); toc.CellBold(headerRow, 3, true); toc.CellBackground(headerRow, 3, "#F2F2F2"); }
            r++;

            var rowsStart = r;
            foreach (var sh in this.Sheets)
            {
                // Skip the TOC sheet itself; it will be moved to first anyway
                if (string.Equals(sh.Name, sheetName, StringComparison.OrdinalIgnoreCase)) continue;

                if (withHyperlinks) toc.SetInternalLink(r, 1, $"'{sh.Name}'!A1", sh.Name); else toc.Cell(r, 1, sh.Name);

                // Details: Used range and size
                string used = sh.GetUsedRangeA1();
                try
                {
                    var (r1, c1, r2, c2) = OfficeIMO.Excel.Read.A1.ParseRange(used);
                    int rows = System.Math.Max(0, r2 - r1 + 1);
                    int cols = System.Math.Max(0, c2 - c1 + 1);
                    toc.Cell(r, 2, $"Used {used} ({rows}×{cols})");
                }
                catch { toc.Cell(r, 2, $"Used {used}"); }

                if (includeNamedRanges)
                {
                    var names = GetAllNamedRanges(sh);
                    var list = new System.Collections.Generic.List<string>();
                    foreach (var kv in names)
                    {
                        var name = kv.Key;
                        if (rangeNameFilter != null && !rangeNameFilter(name)) continue;
                        if (dnMeta.TryGetValue(name, out var dn) && dn.Hidden == true && !includeHiddenNamedRanges) continue;
                        list.Add(name);
                    }
                    string joined = list.Count == 0 ? "—" : string.Join(", ", list);
                    toc.Cell(r, 3, joined);
                }
                r++;
            }
            int rowsEnd = r - 1;

            if (styled && rowsEnd >= rowsStart)
            {
                string endCol = includeNamedRanges ? "C" : "B";
                string tableRange = $"A{headerRow}:{endCol}{rowsEnd}";
                toc.AddTable(tableRange, hasHeader: true, name: "TOC_Items", style: TableStyle.TableStyleMedium2, includeAutoFilter: true);
                try { toc.Freeze(topRows: headerRow, leftCols: 0); } catch { }
                toc.AutoFitColumns();
            }
            if (placeFirst) MoveSheetToBeginning(sheetName);
        }

        /// <summary>
        /// Backward-compatible alias for AddTableOfContents.
        /// </summary>
        public void CreateTableOfContents(string sheetName = "TOC") => AddTableOfContents(sheetName, placeFirst: true, withHyperlinks: true);

        /// <summary>
        /// Adds a small back link to the TOC on each worksheet at the given cell (default A2).
        /// </summary>
        public void AddBackLinksToToc(string tocSheetName = "TOC", int row = 2, int col = 1, string text = "← TOC")
        {
            foreach (var sh in this.Sheets)
            {
                if (string.Equals(sh.Name, tocSheetName, StringComparison.OrdinalIgnoreCase)) continue;
                sh.SetInternalLink(row, col, $"'{tocSheetName}'!A1", text);
            }
        }

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
