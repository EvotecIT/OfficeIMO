using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // Represents a table in a Word document and exposes various
    // properties controlling its appearance and behavior.
    public partial class WordTable {
        /// <summary>
        /// Exposed for internal callers in this assembly that change table structure
        /// (e.g., InsertColumn) and need to update the tblGrid.
        /// </summary>
        internal void RefreshGrid() => RefreshTblGridFromColumnWidths();

        /// <summary>
        /// Estimates the effective table width in DXA (twips) based on
        /// Table.Width and the section page width/margins.
        /// </summary>
        internal int EstimateTableWidthInDxa() {
            // For nested tables, use the containing cell's width as the reference
            if (GetWidthHostCell() != null) {
                int container = EstimateContainingCellContentWidthInDxa();
                if (WordTableWidthNative.ResolveType(_tableProperties?.TableWidth) == TableWidthUnitValues.Dxa && TableWidthNativeValue > 0) {
                    return Math.Min(TableWidthNativeValue, container);
                }
                if (WordTableWidthNative.ResolveType(_tableProperties?.TableWidth) == TableWidthUnitValues.Pct && TableWidthFraction > 0) {
                    int desired = (int)Math.Round((double)container * TableWidthFraction);
                    return Math.Min(desired, container);
                }
                // Auto or unspecified => fit to container
                return container;
            }

            // A loaded shared header/footer may have several different layout widths.
            // Its existing grid is already the resolved table width, including tblW percent.
            if (WordTableWidthNative.ResolveType(_tableProperties?.TableWidth) != TableWidthUnitValues.Dxa && GetUnscopedStoryGridWidth() is int storyWidth)
                return storyWidth;

            // Non-nested: default to page content area as reference
            int contentWidth = EstimateContentAreaWidthInDxa();

            if (WordTableWidthNative.ResolveType(_tableProperties?.TableWidth) == TableWidthUnitValues.Dxa && TableWidthNativeValue > 0) {
                return TableWidthNativeValue;
            }
            if (WordTableWidthNative.ResolveType(_tableProperties?.TableWidth) == TableWidthUnitValues.Pct && TableWidthFraction > 0) {
                // Width is in 1/50 %, 5000 == 100%
                return (int)Math.Round((double)contentWidth * TableWidthFraction);
            }
            // Auto or unspecified
            return contentWidth;
        }

        /// <summary>
        /// Estimates the width available to this table in its page or containing
        /// table cell, before the table's own authored width is applied.
        /// </summary>
        internal int EstimateAvailableContainerWidthInDxa() =>
            GetWidthHostCell() != null
                ? EstimateContainingCellContentWidthInDxa()
                : EstimateContentAreaWidthInDxa();

        /// <summary>
        /// Returns the estimated text area width (page width minus left/right margins) in DXA.
        /// Uses the section that contains the table when it can be resolved.
        /// </summary>
        private int EstimateContentAreaWidthInDxa() {
            try {
                var section = ResolveOwningSection();
                if (section != null) {
                    var page = section.PageSettings;
                    var width = (int)(checked((uint?)(page.Width?.ToTwips())) ?? checked((uint)(WordPageSizes.A4.Width.ToTwips())));
                    var left = (int)checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips()));
                    var right = (int)checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips()));
                    int content = Math.Max(0, width - left - right);
                    return content > 0 ? content : 9000; // fallback ~6.25"
                }
            } catch { /* ignore */ }
            // Sensible default if anything fails
            return 9000; // ~6.25 inches
        }

        /// <summary>
        /// Estimates the usable content width of a specific table cell in DXA (twips).
        /// </summary>
        internal static int? EstimateCellContentWidthInDxa(
            WordDocument document,
            DocumentFormat.OpenXml.Wordprocessing.TableCell cell) {
            try {
                var row = cell.Parent as DocumentFormat.OpenXml.Wordprocessing.TableRow;
                var parentTable = row?.Parent as DocumentFormat.OpenXml.Wordprocessing.Table;
                if (row != null && parentTable != null &&
                    WordTableCellSpanResolver.TryGetLogicalColumn(cell, out int gridIndex)) {
                    int spanThis = WordTableCellSpanResolver.GetColumnSpan(cell);
                    return CreateCellContentWidthEstimatorInDxa(document, parentTable)(
                        cell,
                        gridIndex,
                        spanThis);
                }
            } catch { /* ignore */ }
            return null;
        }

        /// <summary>
        /// Creates a table-scoped cell-width estimator that materializes grid widths and the
        /// fallback table width once, then resolves cells by their known grid position.
        /// </summary>
        internal static Func<DocumentFormat.OpenXml.Wordprocessing.TableCell, int, int, int?>
            CreateCellContentWidthEstimatorInDxa(
                WordDocument document,
                DocumentFormat.OpenXml.Wordprocessing.Table table) {
            int[] gridWidths = table
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TableGrid>()?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.GridColumn>()
                .Select(column => int.TryParse(column.Width?.Value ?? "0", out int width)
                    ? width
                    : 0)
                .ToArray() ?? Array.Empty<int>();
            int? fallbackWidth = null;
            try {
                var parent = new WordTable(document, table, initializeChildren: false);
                fallbackWidth = Math.Max(1, parent.EstimateTableWidthInDxa());
            } catch { /* ignore */ }

            var tableProperties = table
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TableProperties>();
            var defaultMargins = tableProperties?.TableCellMarginDefault;
            return (cell, gridIndex, gridSpan) => {
                try {
                    int sum = 0;
                    int normalizedSpan = Math.Max(Math.Max(1, gridSpan), WordTableCellSpanResolver.GetColumnSpan(cell));
                    for (int index = 0;
                         index < normalizedSpan && gridIndex + index < gridWidths.Length;
                         index++) {
                        sum += gridWidths[gridIndex + index];
                    }
                    if (sum <= 0) {
                        return fallbackWidth;
                    }

                    int leftMargin = 0, rightMargin = 0;
                    int leftBorder = 0, rightBorder = 0;
                    var cellMargins = cell.TableCellProperties?.TableCellMargin;
                    if (cellMargins?.LeftMargin?.Width?.Value != null) {
                        int.TryParse(cellMargins.LeftMargin.Width.Value, out leftMargin);
                    }
                    if (cellMargins?.RightMargin?.Width?.Value != null) {
                        int.TryParse(cellMargins.RightMargin.Width.Value, out rightMargin);
                    }
                    if (leftMargin == 0 && defaultMargins?.TableCellLeftMargin?.Width != null) {
                        leftMargin = defaultMargins.TableCellLeftMargin.Width.Value;
                    }
                    if (rightMargin == 0 && defaultMargins?.TableCellRightMargin?.Width != null) {
                        rightMargin = defaultMargins.TableCellRightMargin.Width.Value;
                    }
                    if (leftMargin == 0) leftMargin = 108;
                    if (rightMargin == 0) rightMargin = 108;

                    var cellBorders = cell.TableCellProperties?.TableCellBorders;
                    if (cellBorders?.LeftBorder?.Size != null) {
                        leftBorder = SizeUnitsToTwips(cellBorders.LeftBorder.Size.Value);
                    }
                    if (cellBorders?.RightBorder?.Size != null) {
                        rightBorder = SizeUnitsToTwips(cellBorders.RightBorder.Size.Value);
                    }
                    if (leftBorder == 0) leftBorder = 10;
                    if (rightBorder == 0) rightBorder = 10;

                    return Math.Max(
                        1,
                        sum - leftMargin - rightMargin - leftBorder - rightBorder);
                } catch {
                    return fallbackWidth;
                }
            };
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
                if (Position.TableOverlap == WordTableOverlap.Overlap) return true;
                return false;
            }
            set => Position.TableOverlap = value ? WordTableOverlap.Overlap : WordTableOverlap.Never;
        }

        /// <summary>
        /// Gets or sets the layout algorithm used by Word for the table.
        /// Changing the algorithm preserves the table and cell preferred widths.
        /// Use <see cref="AutoFitToWindow"/> or <see cref="SetFixedWidth(int)"/> when the width should also change.
        /// </summary>
        public WordTableLayoutMode LayoutMode {
            get => GetCurrentLayoutMode();
            set {
                switch (value) {
                    case WordTableLayoutMode.AutoFit:
                        SetLayoutAlgorithm(TableLayoutValues.Autofit);
                        break;
                    case WordTableLayoutMode.Fixed:
                        SetLayoutAlgorithm(TableLayoutValues.Fixed);
                        break;
                }
            }
        }

        /// <summary>
        /// Allow text to wrap around table.
        /// </summary>
        public bool AllowTextWrap {
            get {
                if (Position.VerticalAnchor == WordTableVerticalAnchor.Text) return true;

                return false;
            }
            set {
                if (value)
                    Position.VerticalAnchor = WordTableVerticalAnchor.Text;
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

        /// <summary>Gets or sets physical grid geometry independently of preferred widths.</summary>
        public List<global::OfficeIMO.OfficeLength> GridColumnWidth {
            get => _table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
                .Where(column => column.Width?.Value != null)
                .Select(column => global::OfficeIMO.OfficeLength.Emus(checked((long)int.Parse(column.Width!.Value!, System.Globalization.CultureInfo.InvariantCulture) * 635)))
                .ToList() ?? new List<global::OfficeIMO.OfficeLength>();
            set {
                if (value == null) throw new ArgumentNullException(nameof(value));
                // Validate the complete update before replacing the native grid.
                var columns = value.Select(width => new GridColumn {
                    Width = WordTableWidth.Length(width).ToNativeValue().ToString(System.Globalization.CultureInfo.InvariantCulture)
                }).ToArray();
                TableGrid? grid = _table.GetFirstChild<TableGrid>();
                if (grid == null) {
                    CheckTableProperties();
                    grid = new TableGrid();
                    _table.InsertAfter(grid, _tableProperties);
                }
                grid.RemoveAllChildren<GridColumn>();
                grid.Append(columns);
            }
        }

        /// <summary>Gets first-row positional cell preferences, including absent slots, or applies preferences to each row.</summary>
        /// <remarks>Each item retains its own interpretation. Rows with fewer cells keep their existing shape;
        /// this collection addresses physical cells, not merged logical grid slots.</remarks>
        public List<WordTableWidth?> ColumnWidth {
            get => Rows.Count == 0 ? new List<WordTableWidth?>() : Rows[0].Cells.Select(cell => cell.Width).ToList();
            set {
                if (value == null) throw new ArgumentNullException(nameof(value));
                if (Rows.Any(row => row.Cells.Count < value.Count))
                    throw new ArgumentException("Column preferences exceed a row's physical cell count.", nameof(value));
                bool suppressed = _suppressGridRefresh;
                _suppressGridRefresh = true;
                try {
                    foreach (var row in Rows)
                        for (int i = 0; i < value.Count; i++) row.Cells[i].Width = value[i];
                } finally { _suppressGridRefresh = suppressed; }
                if (!suppressed) RefreshTblGridFromColumnWidths();
            }
        }

        /// <summary>Gets declared row heights, including absent constraints, or applies exact heights in row order.</summary>
        /// <remarks>Null removes a row's height constraint. All values are checked before any row is changed.</remarks>
        public List<OfficeLength?> RowHeight {
            get => Rows.Select(row => row.Height).ToList();
            set {
                if (value == null) throw new ArgumentNullException(nameof(value));
                if (value.Count > Rows.Count)
                    throw new ArgumentException("Row heights exceed the table's row count.", nameof(value));
                foreach (OfficeLength? height in value) {
                    if (height < OfficeLength.Zero)
                        throw new ArgumentOutOfRangeException(nameof(value), "Row height cannot be negative.");
                    if (height.HasValue) _ = checked((int)height.Value.ToTwips());
                }
                for (int index = 0; index < value.Count; index++) Rows[index].Height = value[index];
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

                        foreach (var structuredDocumentTag in cell.Elements.OfType<WordStructuredDocumentTag>()) {
                            list.Add(structuredDocumentTag);
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
