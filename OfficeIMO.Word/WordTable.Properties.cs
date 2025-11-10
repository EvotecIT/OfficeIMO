using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a table in a Word document and exposes various
    /// properties controlling its appearance and behavior.
    /// </summary>
    public partial class WordTable {
        /// <summary>
        /// Rebuilds the table grid (w:tblGrid) using DXA widths derived from
        /// current column width values/types and the table preferred width.
        /// Many consumers (Word Online, Google Docs) rely primarily on tblGrid
        /// when laying out columns and ignore cell tcW percentages. Keeping
        /// tblGrid in sync avoids the observed 50/50 column issue.
        /// </summary>
        private void RefreshTblGridFromColumnWidths() {
            if (_suppressGridRefresh) return;
            _suppressGridRefresh = true;
            try {
                // We need both the number of columns and the configured column widths
                if (Rows.Count == 0) return;
                int columnCount = _table.GetFirstChild<TableGrid>()?.OfType<GridColumn>().Count()
                                   ?? Rows.Max(r => r.Cells.Count);
                if (columnCount <= 0) columnCount = Rows[0].CellsCount;

                // If nothing to base on, do nothing
                var colWidths = GetBestAvailableColumnWidths(out var detectedType, columnCount);
                if (colWidths == null || colWidths.Count == 0) return;
                // If we managed to detect a concrete type for widths, prefer it when ColumnWidthType is missing
                if (ColumnWidthType == null && detectedType != null) {
                    ColumnWidthType = detectedType;
                }

                // Ensure list size matches column count by trimming or padding evenly
                if (colWidths.Count > columnCount) {
                    colWidths = colWidths.Take(columnCount).ToList();
                } else if (colWidths.Count < columnCount) {
                    // Pad missing columns with an even share of remaining width
                    int missing = columnCount - colWidths.Count;
                    int add = 0;
                    if (ColumnWidthType == TableWidthUnitValues.Pct) {
                        int used = colWidths.Sum();
                        add = Math.Max(0, (5000 - used) / Math.Max(1, missing));
                    } else {
                        // When widths are DXA, split remaining table width if we can estimate it,
                        // otherwise reuse last value (keeps behaviour stable)
                        add = colWidths.Count > 0 ? colWidths[colWidths.Count - 1] : 2400;
                    }
                    for (int i = 0; i < missing; i++) colWidths.Add(add);
                }

                // Compute the target table width in DXA
                int tableWidthDxa = EstimateTableWidthInDxa();

                // Convert each column width to DXA
                List<int> gridDxa = new List<int>(columnCount);
                if (ColumnWidthType == TableWidthUnitValues.Pct) {
                    // values are stored in 1/50 %, so 5000 == 100%
                    int totalPct = Math.Max(1, colWidths.Sum());
                    // Scale to the table width so that the sum matches the table width
                    int allocated = 0;
                    for (int i = 0; i < columnCount; i++) {
                        int dxa = (int)Math.Round((double)tableWidthDxa * colWidths[i] / totalPct);
                        // Accumulate and fix rounding on the last column
                        if (i == columnCount - 1) dxa = Math.Max(0, tableWidthDxa - allocated);
                        allocated += dxa;
                        gridDxa.Add(Math.Max(1, dxa));
                    }
                } else if (ColumnWidthType == TableWidthUnitValues.Dxa) {
                    // Normalize DXA widths so the sum matches the table container width.
                    // This ensures online viewers neither clip nor shrink the table.
                    int sum = Math.Max(1, colWidths.Sum());
                    int target = Math.Max(1, tableWidthDxa);
                    int allocated = 0;
                    for (int i = 0; i < columnCount; i++) {
                        int dxa = (int)Math.Round((double)target * colWidths[i] / sum);
                        if (i == columnCount - 1) dxa = Math.Max(0, target - allocated);
                        allocated += dxa;
                        gridDxa.Add(Math.Max(1, dxa));
                    }
                } else {
                    // Auto/other: distribute evenly within the estimated table width
                    int baseWidth = columnCount == 0 ? 0 : tableWidthDxa / columnCount;
                    int remainder = tableWidthDxa - baseWidth * columnCount;
                    for (int i = 0; i < columnCount; i++) gridDxa.Add(baseWidth + (i == columnCount - 1 ? remainder : 0));
                }

                // Write/replace tblGrid with computed DXA widths
                TableGrid? tableGrid = _table.GetFirstChild<TableGrid>();
                if (tableGrid == null) {
                    _table.InsertAfter(new TableGrid(), _tableProperties);
                    tableGrid = _table.GetFirstChild<TableGrid>();
                }
                if (tableGrid == null) return; // safety

                tableGrid.RemoveAllChildren();
                foreach (var dxa in gridDxa) {
                    tableGrid.Append(new GridColumn { Width = dxa.ToString() });
                }
            } catch {
                // Never throw from a setter; layout will still be valid in Word desktop.
            }
            finally { _suppressGridRefresh = false; }
        }

        /// <summary>
        /// Attempts to derive a complete set of column widths and their unit by scanning rows.
        /// Prefer any row that has widths for all columns; fall back to the first row with any widths.
        /// </summary>
        private List<int> GetBestAvailableColumnWidths(out TableWidthUnitValues? detectedType, int expectedColumns) {
            detectedType = ColumnWidthType; // start with the table-level hint

            if (Rows.Count == 0) return new List<int>();
            int cols = expectedColumns > 0 ? expectedColumns : Rows.Max(r => r.Cells.Count);
            List<int>? candidate = null;
            TableWidthUnitValues? candidateType = null;

            foreach (var row in Rows) {
                var w = new List<int>();
                TableWidthUnitValues? typeForRow = null;
                bool allPresent = true;
                for (int i = 0; i < System.Math.Min(cols, row.Cells.Count); i++) {
                    var cell = row.Cells[i];
                    var wv = cell.Width;
                    if (wv == null) { allPresent = false; w.Add(0); continue; }
                    w.Add(wv.Value);
                    // Remember the first non-null type encountered
                    typeForRow ??= cell.WidthType;
                }
                while (w.Count < cols) { w.Add(0); allPresent = false; }
                if (w.Any(x => x != 0)) {
                    if (allPresent) {
                        detectedType = typeForRow ?? detectedType;
                        return w;
                    }
                    if (candidate == null) { candidate = w; candidateType = typeForRow; }
                }
            }

            if (candidate != null) {
                detectedType ??= candidateType;
                // Replace zeros with an even share
                int missing = candidate.Count(x => x == 0);
                if (missing > 0) {
                    if ((detectedType ?? ColumnWidthType) == TableWidthUnitValues.Pct) {
                        int used = candidate.Where(x => x > 0).Sum();
                        int add = Math.Max(0, (5000 - used) / missing);
                        for (int i = 0; i < candidate.Count; i++) if (candidate[i] == 0) candidate[i] = add;
                    } else {
                        int even = EstimateTableWidthInDxa() / Math.Max(1, candidate.Count);
                        for (int i = 0; i < candidate.Count; i++) if (candidate[i] == 0) candidate[i] = even;
                    }
                }
                return candidate;
            }

            // No widths anywhere – distribute evenly
            int evenDxa = EstimateTableWidthInDxa() / Math.Max(1, cols);
            detectedType = TableWidthUnitValues.Dxa;
            return Enumerable.Repeat(evenDxa, cols).ToList();
        }

        /// <summary>
        /// Exposed for internal callers in this assembly that change table structure
        /// (e.g., InsertColumn) and need to update the tblGrid.
        /// </summary>
        internal void RefreshGrid() => RefreshTblGridFromColumnWidths();

        /// <summary>
        /// Estimates the effective table width in DXA (twips) based on
        /// Table.Width/WidthType and the section page width/margins.
        /// </summary>
        private int EstimateTableWidthInDxa() {
            // For nested tables, use the containing cell's width as the reference
            if (IsNestedTable) {
                int container = EstimateContainingCellContentWidthInDxa();
                if (this.WidthType == TableWidthUnitValues.Dxa && (this.Width ?? 0) > 0) {
                    return Math.Min(this.Width!.Value, container);
                }
                if (this.WidthType == TableWidthUnitValues.Pct && (this.Width ?? 0) > 0) {
                    int desired = (int)Math.Round((double)container * this.Width!.Value / 5000);
                    return Math.Min(desired, container);
                }
                // Auto or unspecified => fit to container
                return container;
            }

            // Non-nested: default to page content area as reference
            int contentWidth = EstimateContentAreaWidthInDxa();

            if (this.WidthType == TableWidthUnitValues.Dxa && (this.Width ?? 0) > 0) {
                return this.Width!.Value;
            }
            if (this.WidthType == TableWidthUnitValues.Pct && (this.Width ?? 0) > 0) {
                // Width is in 1/50 %, 5000 == 100%
                return (int)Math.Round((double)contentWidth * this.Width!.Value / 5000);
            }
            // Auto or unspecified
            return contentWidth;
        }

        /// <summary>
        /// Returns the estimated text area width (page width minus left/right margins) in DXA.
        /// Uses the first section when the owning section can't be easily resolved.
        /// </summary>
        private int EstimateContentAreaWidthInDxa() {
            try {
                var section = _document.Sections.Count > 0 ? _document.Sections[0] : null;
                if (section != null) {
                    var page = section.PageSettings;
                    var width = (int)(page.Width?.Value ?? WordPageSizes.A4.Width!.Value);
                    var left = (int)(section.Margins.Left?.Value ?? 1440U);
                    var right = (int)(section.Margins.Right?.Value ?? 1440U);
                    int content = Math.Max(0, width - left - right);
                    return content > 0 ? content : 9000; // fallback ~6.25"
                }
            } catch { /* ignore */ }
            // Sensible default if anything fails
            return 9000; // ~6.25 inches
        }

        /// <summary>
        /// Estimates the available content width of the containing table cell (for nested tables).
        /// Falls back to page content width when structure cannot be determined.
        /// </summary>
        private int EstimateContainingCellContentWidthInDxa() {
            try {
                // We expect the table parent to be a TableCell when nested.
                if (_table.Parent is DocumentFormat.OpenXml.Wordprocessing.TableCell cell) {
                    // Parent row and table
                    var row = cell.Parent as DocumentFormat.OpenXml.Wordprocessing.TableRow;
                    var parentTable = row?.Parent as DocumentFormat.OpenXml.Wordprocessing.Table;
                    if (row != null && parentTable != null) {
                        // Determine the starting grid index of this cell by iterating row cells and
                        // accumulating gridSpan for cells before our target.
                        int gridIndex = 0;
                        foreach (var c in row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>()) {
                            if (object.ReferenceEquals(c, cell)) break;
                            int span = (int)(c.TableCellProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.GridSpan>()?.Val?.Value ?? 1);
                            gridIndex += Math.Max(1, span);
                        }

                        int spanThis = (int)(cell.TableCellProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.GridSpan>()?.Val?.Value ?? 1);
                        spanThis = Math.Max(1, spanThis);

                        var grid = parentTable.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TableGrid>();
                        if (grid != null) {
                            var cols = grid.Elements<DocumentFormat.OpenXml.Wordprocessing.GridColumn>().ToList();
                            int sum = 0;
                            for (int i = 0; i < spanThis && (gridIndex + i) < cols.Count; i++) {
                                if (int.TryParse(cols[gridIndex + i].Width?.Value ?? "0", out int w)) sum += w;
                            }
                            if (sum > 0) {
                                // Subtract cell left/right margins and borders to get usable inner width
                                int leftMargin = 0, rightMargin = 0;
                                int leftBorder = 0, rightBorder = 0;

                                // Margins: prefer explicit cell margins; fall back to table defaults; else Word default 108 twips
                                var cellMar = cell.TableCellProperties?.TableCellMargin;
                                if (cellMar?.LeftMargin?.Width?.Value != null) {
                                    int.TryParse(cellMar.LeftMargin.Width.Value, out leftMargin);
                                }
                                if (cellMar?.RightMargin?.Width?.Value != null) {
                                    int.TryParse(cellMar.RightMargin.Width.Value, out rightMargin);
                                }

                                if (leftMargin == 0 || rightMargin == 0) {
                                    var ptProps = parentTable.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TableProperties>();
                                    // TableCellMarginDefault stores left/right in twips (DXA) as Int16
                                    var dflt = ptProps?.TableCellMarginDefault;
                                    if (leftMargin == 0 && dflt?.TableCellLeftMargin?.Width != null) leftMargin = dflt.TableCellLeftMargin.Width.Value;
                                    if (rightMargin == 0 && dflt?.TableCellRightMargin?.Width != null) rightMargin = dflt.TableCellRightMargin.Width.Value;
                                }

                                if (leftMargin == 0) leftMargin = 108; // Word default ~0.075"
                                if (rightMargin == 0) rightMargin = 108;

                                // Borders: check explicit cell borders first; else assume style default ~10 twips per side (size=4 → 0.5pt)
                                var cellBorders = cell.TableCellProperties?.TableCellBorders;
                                if (cellBorders?.LeftBorder?.Size != null) leftBorder = SizeUnitsToTwips(cellBorders.LeftBorder.Size.Value);
                                if (cellBorders?.RightBorder?.Size != null) rightBorder = SizeUnitsToTwips(cellBorders.RightBorder.Size.Value);
                                if (leftBorder == 0) leftBorder = 10;
                                if (rightBorder == 0) rightBorder = 10;

                                int usable = Math.Max(1, sum - leftMargin - rightMargin - leftBorder - rightBorder);
                                return usable;
                            }
                        }

                        // Fallback to parent table estimated width when grid is unavailable
                        var parent = new WordTable(_document, parentTable, initializeChildren: false);
                        return Math.Max(1, parent.EstimateTableWidthInDxa());
                    }
                }
            } catch { /* ignore */ }
            return EstimateContentAreaWidthInDxa();
        }

        private static int SizeUnitsToTwips(UInt32Value sizeUnits) {
            // Border size is in eighths of a point. 1pt = 20 twips → 20/8 = 2.5 twips per unit.
            // Round up to avoid fractional loss.
            try { return (int)Math.Ceiling(sizeUnits.Value * 2.5); } catch { return 0; }
        }
        /// <summary>
        /// Gets or sets a Title/Caption to a Table
        /// </summary>
        public string? Title {
            get {
                if (_tableProperties != null && _tableProperties.TableCaption != null)
                    return _tableProperties.TableCaption.Val;

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableCaption == null) _tableProperties.TableCaption = new TableCaption();
                if (value != null)
                    _tableProperties.TableCaption.Val = value;
                else
                    _tableProperties.TableCaption.Remove();
            }
        }

        /// <summary>
        /// Gets or sets Description for a Table
        /// </summary>
        public string? Description {
            get {
                if (_tableProperties != null && _tableProperties.TableDescription != null)
                    return _tableProperties.TableDescription.Val;

                return null;
            }
            set {
                CheckTableProperties();
                if (_tableProperties!.TableDescription == null)
                    _tableProperties.TableDescription = new TableDescription();
                if (value != null)
                    _tableProperties.TableDescription.Val = value;
                else
                    _tableProperties.TableDescription.Remove();
            }
        }

        /// <summary>
        /// Allow table to overlap or not
        /// </summary>
        public bool AllowOverlap {
            get {
                if (Position.TableOverlap == TableOverlapValues.Overlap) return true;
                return false;
            }
            set => Position.TableOverlap = value ? TableOverlapValues.Overlap : TableOverlapValues.Never;
        }

        /// <summary>
        /// Gets or sets the effective layout mode of the table using WordTableLayoutType enum.
        /// Setting FixedWidth via this property defaults to 100% width.
        /// Use SetFixedWidth(percentage) for specific percentages.
        /// </summary>
        public WordTableLayoutType LayoutMode {
            get => GetCurrentLayoutType();
            set {
                switch (value) {
                    case WordTableLayoutType.AutoFitToContents:
                        AutoFitToContents();
                        break;
                    case WordTableLayoutType.AutoFitToWindow:
                        AutoFitToWindow();
                        break;
                    case WordTableLayoutType.FixedWidth:
                        // Default to 100% when setting via this property
                        SetFixedWidth(100);
                        break;
                }
            }
        }

        /// <summary>
        /// Gets or sets the AutoFit behavior for the table. Alias for <see cref="LayoutMode"/>.
        /// </summary>
        public WordTableLayoutType AutoFit {
            get => LayoutMode;
            set => LayoutMode = value;
        }

        /// <summary>
        /// Allow text to wrap around table.
        /// </summary>
        public bool AllowTextWrap {
            get {
                if (Position.VerticalAnchor == VerticalAnchorValues.Text) return true;

                return false;
            }
            set {
                if (value)
                    Position.VerticalAnchor = VerticalAnchorValues.Text;
                else
                    Position.VerticalAnchor = null;
            }
        }

        /// <summary>
        /// Gets or sets whether text wraps within all cells of the table.
        /// </summary>
        public bool WrapText {
            get {
                return Rows.SelectMany(row => row.Cells).All(cell => cell.WrapText);
            }
            set {
                foreach (var row in Rows) {
                    foreach (var cell in row.Cells) {
                        cell.WrapText = value;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets whether text is compressed to fit within all cells of the table.
        /// </summary>
        public bool FitText {
            get {
                return Rows.SelectMany(row => row.Cells).All(cell => cell.FitText);
            }
            set {
                foreach (var row in Rows) {
                    foreach (var cell in row.Cells) {
                        cell.FitText = value;
                    }
                }
            }
        }

        /// <summary>
        /// Sets or gets grid columns width (not really doing anything as far as I can see)
        /// </summary>
        public List<int> GridColumnWidth {
            get {
                var listReturn = new List<int>();
                TableGrid? tableGrid = _table.GetFirstChild<TableGrid>();
                if (tableGrid != null) {
                    var list = tableGrid.OfType<GridColumn>();
                    foreach (var column in list) {
                        if (column.Width != null && column.Width.Value != null) {
                            listReturn.Add(int.Parse(column.Width.Value));
                        }
                    }
                }
                return listReturn;
            }
            set {
                TableGrid? tableGrid = _table.GetFirstChild<TableGrid>();
                if (tableGrid != null) {
                    tableGrid.RemoveAllChildren();
                } else {
                    _table.InsertAfter(new TableGrid(), _tableProperties);
                    tableGrid = _table.GetFirstChild<TableGrid>();
                }
                if (tableGrid != null) {
                    foreach (var columnWidth in value) {
                        tableGrid.Append(new GridColumn { Width = columnWidth.ToString() });
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets column width for a whole table simplifying setup of column width
        /// Please note that it assumes first row has the same width as the rest of rows
        /// which may give false positives if there are multiple values set differently.
        /// </summary>
        public List<int> ColumnWidth {
            get {
                var listReturn = new List<int>();
                // we assume the first row has the same widths as all rows, which may or may not be true
                for (int cellIndex = 0; cellIndex < this.Rows[0].CellsCount; cellIndex++) {
                    var width = this.Rows[0].Cells[cellIndex].Width;
                    if (width.HasValue) {
                        listReturn.Add(width.Value);
                    }
                }
                return listReturn;
            }
            set {
                for (int cellIndex = 0; cellIndex < value.Count; cellIndex++) {
                    foreach (var row in this.Rows) {
                        row.Cells[cellIndex].Width = value[cellIndex];
                    }
                }
                // Keep tblGrid in sync for non-desktop renderers
                if (!_suppressGridRefresh) RefreshTblGridFromColumnWidths();
            }
        }

        /// <summary>
        /// Gets or sets the column width type for a whole table simplifying setup of column width
        /// </summary>
        public TableWidthUnitValues? ColumnWidthType {
            get {
                var listReturn = new List<TableWidthUnitValues?>();
                // we assume the first row has the same widths as all rows, which may or may not be true
                for (int cellIndex = 0; cellIndex < this.Rows[0].CellsCount; cellIndex++) {
                    listReturn.Add(this.Rows[0].Cells[cellIndex].WidthType);
                }
                // we assume all cells have the same width type, which may or may not be true
                return listReturn[0];
            }
            set {
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        cell.WidthType = value;
                    }
                }
                // Update tblGrid to reflect width type changes (Pct -> DXA conversion)
                if (!_suppressGridRefresh) RefreshTblGridFromColumnWidths();
            }
        }

        /// <summary>
        /// Get or set row heights for the table
        /// </summary>
        public List<int> RowHeight {
            get {
                var listReturn = new List<int>();
                for (int rowIndex = 0; rowIndex < this.Rows.Count; rowIndex++) {
                    listReturn.Add(this.Rows[rowIndex].Height ?? 0);
                }
                return listReturn;
            }
            set {
                for (int rowIndex = 0; rowIndex < value.Count; rowIndex++) {
                    this.Rows[rowIndex].Height = value[rowIndex];
                }
            }

        }

        /// <summary>
        /// Get all WordTableCells in a table. A short way to loop thru all cells
        /// </summary>
        public List<WordTableCell> Cells {
            get {
                var listReturn = new List<WordTableCell>();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        listReturn.Add(cell);
                    }
                }
                return listReturn;
            }
        }


        /// <summary>
        /// Gets information whether the Table has other nested tables in at least one of the TableCells
        /// </summary>
        public bool HasNestedTables {
            get {
                foreach (var cell in this.Cells) {
                    if (cell.HasNestedTables) {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Get all nested tables in the table
        /// </summary>
        public List<WordTable> NestedTables {
            get {
                var listReturn = new List<WordTable>();
                foreach (var cell in this.Cells) {
                    listReturn.AddRange(cell.NestedTables);
                }
                return listReturn;
            }
        }

        /// <summary>
        /// Gets information whether the table is nested table (within TableCell)
        /// </summary>
        public bool IsNestedTable {
            get {
                var openXmlElement = this._table.Parent;
                if (openXmlElement != null) {
                    var typeOfParent = openXmlElement.GetType();
                    if (typeOfParent.FullName == "DocumentFormat.OpenXml.Wordprocessing.TableCell") {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Gets nested table parent table if table is nested table
        /// </summary>
        public WordTable? ParentTable {
            get {
                if (IsNestedTable) {
                    if (this._table.Parent?.Parent?.Parent is Table table) {
                        return new WordTable(this._document, table);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets all structured document tags contained in the table.
        /// </summary>
        public List<WordStructuredDocumentTag> StructuredDocumentTags {
            get {
                List<WordStructuredDocumentTag> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsStructuredDocumentTag).ToList();
                        foreach (var paragraph in paragraphs) {
                            var structuredDocumentTag = paragraph.StructuredDocumentTag;
                            if (structuredDocumentTag != null) {
                                list.Add(structuredDocumentTag);
                            }
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets all checkbox content controls contained in the table.
        /// </summary>
        public List<WordCheckBox> CheckBoxes {
            get {
                List<WordCheckBox> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsCheckBox).ToList();
                        foreach (var paragraph in paragraphs) {
                            var checkBox = paragraph.CheckBox;
                            if (checkBox != null) {
                                list.Add(checkBox);
                            }
                        }
                    }
                }
                return list;
            }
        }
        /// <summary>
        /// Gets all date picker content controls contained in the table.
        /// </summary>
        public List<WordDatePicker> DatePickers {
            get {
                List<WordDatePicker> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsDatePicker).ToList();
                        foreach (var paragraph in paragraphs) {
                            var datePicker = paragraph.DatePicker;
                            if (datePicker != null) {
                                list.Add(datePicker);
                            }
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets all dropdown list content controls contained in the table.
        /// </summary>
        public List<WordDropDownList> DropDownLists {
            get {
                List<WordDropDownList> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsDropDownList).ToList();
                        foreach (var paragraph in paragraphs) {
                            var dropDownList = paragraph.DropDownList;
                            if (dropDownList != null) {
                                list.Add(dropDownList);
                            }
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets all repeating section content controls contained in the table.
        /// </summary>
        public List<WordRepeatingSection> RepeatingSections {
            get {
                List<WordRepeatingSection> list = new();
                foreach (var row in this.Rows) {
                    foreach (var cell in row.Cells) {
                        var paragraphs = cell.Paragraphs.Where(p => p.IsRepeatingSection).ToList();
                        foreach (var paragraph in paragraphs) {
                            var repeatingSection = paragraph.RepeatingSection;
                            if (repeatingSection != null) {
                                list.Add(repeatingSection);
                            }
                        }
                    }
                }
                return list;
            }
        }
    }
}
