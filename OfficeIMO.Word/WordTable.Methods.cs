using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        /// <summary>
        /// Add comment to a Table
        /// </summary>
        /// <param name="author">Provide an author of the comment</param>
        /// <param name="initials">Provide initials of an author</param>
        /// <param name="comment">Provide comment to insert</param>
        public void AddComment(string author, string initials, string comment) {
            WordComment wordComment = WordComment.Create(_document, author, initials, comment);
            InsertComment(wordComment,
                this.FirstRow.FirstCell.Paragraphs[0]._paragraph,
                this.LastRow.LastCell.Paragraphs[0]._paragraph,
                this.LastRow.LastCell.Paragraphs[0]._paragraph);
        }

        internal void InsertComment(WordComment wordComment, OpenXmlElement rangeStart, OpenXmlElement rangeEnd, OpenXmlElement reference) {
            // Specify the text range for the Comment.
            // Insert the new CommentRangeStart before the first run of paragraph.
            rangeStart.InsertBefore(new CommentRangeStart() { Id = wordComment.Id }, rangeStart.GetFirstChild<OpenXmlElement>());

            // Insert the new CommentRangeEnd after last run of paragraph.
            var cmtEnd = rangeEnd.InsertAfter(new CommentRangeEnd() { Id = wordComment.Id }, rangeEnd.Elements().Last());

            // Compose a run with CommentReference and insert it.
            reference.InsertAfter(new Run(new CommentReference() { Id = wordComment.Id }), cmtEnd);
        }

        /// <summary>
        /// Distribute columns evenly by setting their size to the same value
        /// based on the current table width settings (or 100% Pct if not set).
        /// </summary>
        public void DistributeColumnsEvenly() {
            if (Rows.Count == 0 || Rows[0].Cells.Count == 0) return;
            int columnCount = Rows[0].Cells.Count;
            if (columnCount == 0) return;

            CheckTableProperties();

            // Determine the target total width and type to distribute
            TableWidthUnitValues targetType = this.WidthType ?? TableWidthUnitValues.Pct; // Default to Pct if not set
            int targetTotalWidth;

            if (targetType == TableWidthUnitValues.Pct) {
                targetTotalWidth = this.Width ?? 5000; // Default to 100% (5000) if width not set
            } else if (targetType == TableWidthUnitValues.Dxa) {
                targetTotalWidth = this.Width ?? 0;
                // If Dxa width is 0 or not set, it's ambiguous. Default to distributing 100% Pct.
                if (targetTotalWidth <= 0) {
                    targetType = TableWidthUnitValues.Pct;
                    targetTotalWidth = 5000;
                }
            } else { // Auto or unspecified - default to distributing 100% Pct
                targetType = TableWidthUnitValues.Pct;
                targetTotalWidth = 5000;
            }

            // Ensure Table Width properties reflect the distribution target
            // Setting these ensures the table container matches the distributed columns
            this.WidthType = targetType;
            this.Width = targetTotalWidth;

            // Calculate width per column, handling potential rounding for the last column
            int baseColumnWidth = targetTotalWidth / columnCount;
            int remainder = targetTotalWidth % columnCount;

            List<int> newColumnWidths = new List<int>();
            for (int i = 0; i < columnCount; i++) {
                int currentWidth = baseColumnWidth;
                if (i == columnCount - 1) { // Add remainder to the last column
                    currentWidth += remainder;
                }
                newColumnWidths.Add(currentWidth);
            }

            // Apply the new widths to all cells in all rows
            foreach (var row in this.Rows) {
                // Ensure row has the expected number of cells for safety
                if (row.Cells.Count == columnCount) {
                    for (int j = 0; j < columnCount; j++) {
                        var cell = row.Cells[j];
                        // Ensure TableCellProperties and TableCellWidth exist
                        var tcPr = cell._tableCellProperties ?? new TableCellProperties();
                        if (cell._tableCellProperties == null) cell._tableCell.InsertAt(tcPr, 0);

                        var tcW = tcPr.Elements<TableCellWidth>().FirstOrDefault() ?? new TableCellWidth();
                        if (!tcPr.Elements<TableCellWidth>().Any()) tcPr.Append(tcW);

                        // Set the calculated type and width for the cell
                        tcW.Type = targetType;
                        tcW.Width = newColumnWidths[j].ToString();
                    }
                }
            }

            // Update the TableGrid for consistency
            TableGrid tableGrid = _table.GetFirstChild<TableGrid>();
            if (tableGrid != null) {
                tableGrid.RemoveAllChildren<GridColumn>();
                foreach (int width in newColumnWidths) {
                    tableGrid.Append(new GridColumn() { Width = width.ToString() });
                }
            }
        }

        public WordTable SetStyleId(string styleId) {
            //Todo Check the styleId exist
            if (!string.IsNullOrEmpty(styleId)) {
                if (_tableProperties?.TableStyle == null) {
                    _tableProperties.TableStyle = new TableStyle() { Val = styleId };
                } else {
                    _tableProperties.TableStyle.Val = styleId;
                }
            }
            return this;
        }

        /// <summary>
        /// Copy existing WordTableRow and inserts it as a last row in a Table
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public WordTableRow CopyRow(WordTableRow row) {
            // Ensure the table and row are not null
            if (_table == null || row == null) {
                throw new InvalidOperationException("The table doesn't exists or rows doesn't exists");
            }

            // Get the last row in the table
            var lastRow = _table.Elements<TableRow>().LastOrDefault();
            if (lastRow == null) {
                throw new InvalidOperationException("The table does not contain any rows.");
            }

            // Clone the row to avoid the "part of a tree" error
            var clonedRow = (TableRow)row._tableRow.CloneNode(true);

            // Insert the new row after the last row
            var insertedRow = lastRow.InsertAfterSelf(clonedRow);

            return new WordTableRow(this, insertedRow, _document);
        }

        /// <summary>
        /// Sets the table layout with proper AutoFit options
        /// </summary>
        /// <param name="layoutType">Type of layout to apply</param>
        /// <param name="percentage">Optional percentage for fixed width (0-100)</param>
        public void SetTableLayout(WordTableLayoutType layoutType, int? percentage = null) {
            CheckTableProperties();

            // Apply the appropriate settings based on the layout type
            switch (layoutType) {
                case WordTableLayoutType.FixedWidth:
                    // Set OpenXML layout type to Fixed
                    if (_tableProperties.TableLayout == null) {
                        _tableProperties.TableLayout = new TableLayout();
                    }
                    _tableProperties.TableLayout.Type = TableLayoutValues.Fixed;

                    if (percentage.HasValue) {
                        // For fixed width, set the width to the specified percentage
                        this.WidthType = TableWidthUnitValues.Pct;
                        this.Width = percentage.Value * 50; // Convert percentage to Word's internal units (50 = 1%)
                    } else {
                        // Default to 100% if no percentage specified
                        this.WidthType = TableWidthUnitValues.Pct;
                        this.Width = 5000; // 100% width
                    }
                    break;

                case WordTableLayoutType.AutoFitToContents:
                    // Set OpenXML layout type to Autofit
                    if (_tableProperties.TableLayout == null) {
                        _tableProperties.TableLayout = new TableLayout();
                    }
                    _tableProperties.TableLayout.Type = TableLayoutValues.Autofit;

                    // For AutoFit to Contents
                    this.WidthType = TableWidthUnitValues.Auto;
                    this.Width = 0;
                    break;

                case WordTableLayoutType.AutoFitToWindow:
                    // Set OpenXML layout type to Fixed
                    if (_tableProperties.TableLayout == null) {
                        _tableProperties.TableLayout = new TableLayout();
                    }
                    _tableProperties.TableLayout.Type = TableLayoutValues.Fixed;

                    // For AutoFit to Window
                    this.WidthType = TableWidthUnitValues.Pct;
                    this.Width = 5000; // 100% width
                    break;
            }
        }

        /// <summary>
        /// Sets the table to AutoFit to Contents
        /// </summary>
        public void AutoFitToContents() {
            CheckTableProperties();

            // 1. Set Table Layout to Autofit
            if (_tableProperties.TableLayout == null) {
                _tableProperties.TableLayout = new TableLayout();
            }
            _tableProperties.TableLayout.Type = TableLayoutValues.Autofit;

            // 2. Set Table Width to Auto / 0
            if (_tableProperties.TableWidth == null) {
                _tableProperties.TableWidth = new TableWidth();
            }
            _tableProperties.TableWidth.Type = TableWidthUnitValues.Auto;
            _tableProperties.TableWidth.Width = "0";

            // 3. Clear Cell Widths (Set to Auto)
            // This is crucial for contents to determine width
            foreach (var row in Rows) {
                foreach (var cell in row.Cells) {
                    var tcPr = cell._tableCellProperties;
                    // Ensure TableCellProperties exists
                    if (tcPr == null) {
                        tcPr = new TableCellProperties();
                        cell._tableCell.InsertAt(tcPr, 0); // Insert if doesn't exist
                    } else {
                        // Clear existing width if present
                        var existingWidth = tcPr.Elements<TableCellWidth>().FirstOrDefault();
                        if (existingWidth != null) {
                            existingWidth.Remove();
                        }
                    }
                    // Setting tcW with type=auto and w=0 might be equivalent to removing it
                    // Depending on Word's interpretation. Removing is often cleaner.
                }
            }
            // Remove the complex/inaccurate content estimation for now
            // AdjustColumnWidthsBasedOnContent();
        }

        /// <summary>
        /// Analyzes the content of each cell and adjusts column widths accordingly
        /// </summary>
        private void AdjustColumnWidthsBasedOnContent() {
            if (Rows.Count == 0) return;

            int columnCount = Rows[0].Cells.Count;
            List<int> maxContentWidths = new List<int>(new int[columnCount]);

            // Calculate the maximum content width for each column
            foreach (var row in Rows) {
                for (int i = 0; i < Math.Min(row.Cells.Count, columnCount); i++) {
                    var cell = row.Cells[i];
                    int contentWidth = EstimateContentWidth(cell);
                    maxContentWidths[i] = Math.Max(maxContentWidths[i], contentWidth);
                }
            }

            // Apply calculated widths to columns
            ApplyCalculatedWidths(maxContentWidths);
        }

        /// <summary>
        /// Estimates the width of content in a cell
        /// </summary>
        /// <param name="cell">The table cell to analyze</param>
        /// <returns>Estimated width in DXA units</returns>
        private int EstimateContentWidth(WordTableCell cell) {
            int maxWidth = 0;

            foreach (var paragraph in cell.Paragraphs) {
                // Calculate the width based on text length
                int textWidth = CalculateTextWidth(paragraph);
                maxWidth = Math.Max(maxWidth, textWidth);
            }

            // Add some padding (minimum width of 1000 DXA units, ~0.7 inches)
            return Math.Max(1000, maxWidth);
        }

        /// <summary>
        /// Calculates approximate text width based on content
        /// </summary>
        /// <param name="paragraph">The paragraph to analyze</param>
        /// <returns>Estimated width in DXA units</returns>
        private int CalculateTextWidth(WordParagraph paragraph) {
            // Simple estimation based on character count
            // Average character is roughly 100 DXA units (~0.07 inches)
            // This is a rough approximation - precise measurement would require font metrics
            string text = paragraph.Text;
            if (string.IsNullOrEmpty(text)) return 0;

            // Start with a base width
            int width = 0;

            // Add width for each character (using average character width)
            width += text.Length * 100;

            // Add extra for formatting
            if (paragraph.Bold == true) width += (int)(width * 0.1); // Bold text is wider
            if (paragraph.Italic == true) width += (int)(width * 0.05); // Italic text is slightly wider

            return width;
        }

        /// <summary>
        /// Applies the calculated column widths to the table
        /// </summary>
        /// <param name="columnWidths">List of column widths in DXA units</param>
        private void ApplyCalculatedWidths(List<int> columnWidths) {
            // Set column widths
            this.ColumnWidth = columnWidths;
            this.ColumnWidthType = TableWidthUnitValues.Dxa;

            // Ensure the overall table width is set appropriately
            if (columnWidths.Sum() > 0) {
                this.Width = columnWidths.Sum();
                this.WidthType = TableWidthUnitValues.Dxa;
            }
        }

        /// <summary>
        /// Sets the table to AutoFit to Window (100% width)
        /// </summary>
        public void AutoFitToWindow() {
            CheckTableProperties();

            // 1. Set Table Layout to Fixed (or ensure it exists)
            if (_tableProperties.TableLayout == null) {
                _tableProperties.TableLayout = new TableLayout();
            }
            _tableProperties.TableLayout.Type = TableLayoutValues.Fixed;

            // 2. Set Table Width to 100% Pct
            if (_tableProperties.TableWidth == null) {
                _tableProperties.TableWidth = new TableWidth();
            }
            _tableProperties.TableWidth.Type = TableWidthUnitValues.Pct;
            _tableProperties.TableWidth.Width = "5000"; // 5000 = 100%

            // 3. Remove Table Indentation (if exists)
            _tableProperties.TableIndentation?.Remove();

            // 4. Distribute columns evenly (as percentage)
            if (Rows.Count > 0 && Rows[0].Cells.Count > 0) {
                int columnCount = Rows[0].Cells.Count;
                // Calculate width per column, handling potential rounding for the last column
                int baseColumnWidthPct = 5000 / columnCount;
                int remainder = 5000 % columnCount;

                for (int i = 0; i < Rows.Count; i++) {
                    var row = Rows[i];
                    // Ensure row has the expected number of cells for safety
                    if (row.Cells.Count == columnCount) {
                        for (int j = 0; j < columnCount; j++) {
                            var cell = row.Cells[j];
                            var tcPr = cell._tableCellProperties;
                            if (tcPr == null) {
                                tcPr = new TableCellProperties();
                                cell._tableCell.InsertAt(tcPr, 0);
                            }

                            // Ensure TableCellWidth exists
                            var tcW = tcPr.Elements<TableCellWidth>().FirstOrDefault();
                            if (tcW == null) {
                                tcW = new TableCellWidth();
                                tcPr.Append(tcW);
                            }

                            // Set type and width
                            tcW.Type = TableWidthUnitValues.Pct;
                            // Add remainder to the last column to ensure total is 5000
                            tcW.Width = (j == columnCount - 1) ? (baseColumnWidthPct + remainder).ToString() : baseColumnWidthPct.ToString();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Sets the table to Fixed Width with the specified percentage
        /// </summary>
        /// <param name="percentage">Width percentage (0-100)</param>
        public void SetFixedWidth(int percentage) {
            if (percentage < 0) percentage = 0;
            if (percentage > 100) percentage = 100;

            // Set table layout type to Fixed
            CheckTableProperties();
            if (_tableProperties.TableLayout == null) {
                _tableProperties.TableLayout = new TableLayout();
            }
            _tableProperties.TableLayout.Type = TableLayoutValues.Fixed;

            // Set table width
            if (_tableProperties.TableWidth == null) {
                _tableProperties.TableWidth = new TableWidth();
            }
            _tableProperties.TableWidth.Type = TableWidthUnitValues.Pct;
            _tableProperties.TableWidth.Width = (percentage * 50).ToString(); // Convert percentage to Word units (50 = 1%)

            // Set fixed column widths proportionally
            if (Rows.Count > 0) {
                int columnCount = Rows[0].Cells.Count;
                int columnWidth = percentage * 50 / columnCount;

                foreach (var row in Rows) {
                    foreach (var cell in row.Cells) {
                        var tcPr = cell._tableCellProperties;
                        if (tcPr == null) {
                            tcPr = new TableCellProperties();
                            cell._tableCellProperties = tcPr;
                        }

                        if (tcPr.TableCellWidth == null) {
                            tcPr.TableCellWidth = new TableCellWidth();
                        }
                        tcPr.TableCellWidth.Type = TableWidthUnitValues.Pct;
                        tcPr.TableCellWidth.Width = columnWidth.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// Gets the current table layout mode based on its properties
        /// </summary>
        /// <returns>The current WordTableLayoutType</returns>
        public WordTableLayoutType GetCurrentLayoutType() {
            // Get properties defensively
            TableLayoutValues? layoutType = null;
            TableWidthUnitValues? widthType = null;
            string widthValue = null;

            if (_tableProperties != null) {
                if (_tableProperties.TableLayout != null && _tableProperties.TableLayout.Type != null) {
                    layoutType = _tableProperties.TableLayout.Type.Value;
                }
                if (_tableProperties.TableWidth != null) {
                    if (_tableProperties.TableWidth.Type != null) {
                        widthType = _tableProperties.TableWidth.Type.Value;
                    }
                    widthValue = _tableProperties.TableWidth.Width;
                }
            }

            // Debugging line (optional, remove in production)
            // Console.WriteLine($"DEBUG: Layout={layoutType}, WidthType={widthType}, WidthValue={widthValue}");

            // --- Decision Logic ---

            // 1. Explicit Autofit Layout = AutoFitToContents (Highest priority)
            if (layoutType.HasValue && layoutType.Value == TableLayoutValues.Autofit) {
                return WordTableLayoutType.AutoFitToContents;
            }

            // 2. Width Type Percentage = AutoFitToWindow or FixedWidth
            if (widthType.HasValue && widthType.Value == TableWidthUnitValues.Pct) {
                if (widthValue == "5000") {
                    return WordTableLayoutType.AutoFitToWindow;
                } else {
                    return WordTableLayoutType.FixedWidth;
                }
            }

            // 3. Width Type DXA = FixedWidth
            if (widthType.HasValue && widthType.Value == TableWidthUnitValues.Dxa) {
                return WordTableLayoutType.FixedWidth;
            }

            // 4. Width Type Auto or No Width Spec -> Defaults to AutoFitToWindow visually in Word
            // (Unless LayoutType was explicitly Autofit, which is handled in #1)
            if ((widthType.HasValue && widthType.Value == TableWidthUnitValues.Auto) || !widthType.HasValue) {
                return WordTableLayoutType.AutoFitToWindow;
            }

            // Final fallback - should technically not be reached if logic covers all OpenXML states
            return WordTableLayoutType.AutoFitToWindow;
        }

        /// <summary>
        /// Sets the table width to a percentage of the window width
        /// </summary>
        /// <param name="percentage">Width percentage (0-100)</param>
        public void SetWidthPercentage(int percentage) {
            if (percentage < 0) percentage = 0;
            if (percentage > 100) percentage = 100;
            this.WidthType = TableWidthUnitValues.Pct;
            this.Width = percentage * 50; // Convert percentage to Word's internal units (50 = 1%)
        }
    }
}
