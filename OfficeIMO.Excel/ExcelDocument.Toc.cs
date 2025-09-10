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

            // Build a simple Hidden lookup for defined names only when we plan to list them.
            System.Collections.Generic.Dictionary<string, bool>? dnHiddenLookup = null;
            if (includeNamedRanges)
            {
                dnHiddenLookup = new System.Collections.Generic.Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                var dnRoot = _workBookPart.Workbook.DefinedNames;
                if (dnRoot != null)
                {
                    foreach (var d in dnRoot.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>())
                    {
                        var name = d.Name?.Value;
                        if (string.IsNullOrEmpty(name)) continue;
                        bool hidden = d.Hidden?.Value ?? false;
                        if (dnHiddenLookup.TryGetValue(name!, out var prior)) dnHiddenLookup[name!] = prior || hidden; else dnHiddenLookup[name!] = hidden;
                    }
                }
            }

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
                    var (r1, c1, r2, c2) = A1.ParseRange(used);
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
                        if (string.IsNullOrEmpty(name)) continue;
                        if (rangeNameFilter != null && !rangeNameFilter(name)) continue;
                        if (dnHiddenLookup != null && dnHiddenLookup.TryGetValue(name!, out var isHidden) && isHidden && !includeHiddenNamedRanges) continue;
                        list.Add(name!);
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

            // As a safety net, clean up any broken/duplicate defined names after sheet changes.
            RepairDefinedNames(save: true);
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
            var wb = _workBookPart.Workbook;
            var sheets = wb.Sheets;
            if (sheets == null) return;

            var all = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList();
            var target = all.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.Ordinal));
            if (target == null) return;

            int oldIdx = all.IndexOf(target);
            if (oldIdx <= 0)
            {
                // Already first or not found
                return;
            }

            // Reorder sheet nodes
            target.Remove();
            sheets.InsertAt(target, 0);

            // Adjust LocalSheetId for all defined names to reflect new sheet positions
            var definedNames = wb.DefinedNames;
            if (definedNames != null)
            {
                foreach (var dn in definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>())
                {
                    if (dn.LocalSheetId == null) continue;
                    uint v = dn.LocalSheetId.Value;
                    if (v == (uint)oldIdx) dn.LocalSheetId = 0u; // moved sheet
                    else if (v < (uint)oldIdx) dn.LocalSheetId = v + 1; // sheets that were before moved one shift right
                    // v > oldIdx unchanged
                }
            }

            wb.Save();
        }

        /// <summary>
        /// Removes a worksheet by name, deleting its part and entry in the workbook.
        /// </summary>
        public void RemoveWorkSheet(string sheetName)
        {
            var wb = _workBookPart.Workbook;
            var sheets = wb.Sheets;
            if (sheets == null) return;
            var all = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList();
            var sheet = all.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.Ordinal));
            if (sheet == null) return;

            int removedIdx = all.IndexOf(sheet);
            var relId = sheet.Id?.Value;
            sheet.Remove();

            // Clean up defined names scoped to the removed sheet, and reindex others after the removal
            var definedNames = wb.DefinedNames;
            if (definedNames != null)
            {
                foreach (var dn in definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().ToList())
                {
                    if (dn.LocalSheetId == null) continue;
                    uint v = dn.LocalSheetId.Value;
                    if (v == (uint)removedIdx)
                    {
                        // Remove names that belonged to the deleted sheet to avoid Excel repair
                        dn.Remove();
                    }
                    else if (v > (uint)removedIdx)
                    {
                        // Shift indices down so names remain attached to the same logical sheet
                        dn.LocalSheetId = v - 1;
                    }
                }
                if (!definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Any())
                {
                    wb.DefinedNames = null;
                }
            }

            if (!string.IsNullOrEmpty(relId))
            {
                var part = (DocumentFormat.OpenXml.Packaging.WorksheetPart)_workBookPart.GetPartById(relId!);
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
