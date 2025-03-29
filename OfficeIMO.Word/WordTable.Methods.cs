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
        /// </summary>
        public void DistributeColumnsEvenly() {
            this.Width = 0;
            this.WidthType = TableWidthUnitValues.Auto;

            var columnWidth = this.ColumnWidth;
            if (columnWidth.Count == 0) {
                return;
            }
            // check if column width entries are the same
            var firstWidth = columnWidth[0];
            var allSame = columnWidth.All(w => w == firstWidth);
            if (allSame) {
                return;
            }

            // set all column widths to the same value
            var sum = columnWidth.Sum(); // sum of all columns
            var count = columnWidth.Count(); // count of all columns

            if (ColumnWidthType == TableWidthUnitValues.Pct) {
                // 100% = 5000
                var currentPercent = (int)(sum / 5000.0 * 100);
                var newPercent = (int)(100.0 / count);
                var diff = currentPercent - newPercent;
                var newWidth = (int)(newPercent * 5000.0 / 100);
                for (int i = 0; i < columnWidth.Count; i++) {
                    columnWidth[i] = newWidth;
                }
                // add the difference to the last column
                columnWidth[columnWidth.Count - 1] += (int)(diff * 5000.0 / 100);
                this.ColumnWidth = columnWidth;


            } else if (ColumnWidthType == TableWidthUnitValues.Dxa) {
                var totalWidth = columnWidth.Sum();
                var newWidth = totalWidth / columnWidth.Count;
                for (int i = 0; i < columnWidth.Count; i++) {
                    columnWidth[i] = newWidth;
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
            SetTableLayout(WordTableLayoutType.AutoFitToContents);
        }

        /// <summary>
        /// Sets the table to AutoFit to Window (100% width)
        /// </summary>
        public void AutoFitToWindow() {
            SetTableLayout(WordTableLayoutType.AutoFitToWindow);
        }

        /// <summary>
        /// Sets the table to Fixed Width with specified percentage
        /// </summary>
        /// <param name="percentage">Width percentage (0-100)</param>
        public void SetFixedWidth(int percentage) {
            if (percentage < 0) percentage = 0;
            if (percentage > 100) percentage = 100;
            SetTableLayout(WordTableLayoutType.FixedWidth, percentage);
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
