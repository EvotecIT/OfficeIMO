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


            //// set all column widths to the same value
            //var sum = columnWidth.Sum();
            //var count = columnWidth.Count();
            //var newWidth = sum / count;

            //for (int i = 0; i < columnWidth.Count; i++) {
            //    columnWidth[i] = newWidth;
            //}
            //ColumnWidth = columnWidth;

            //// set table width to the sum of column widths
            //this.Width = sum;
            //this.WidthType = this.ColumnWidthType;
        }

        /// <summary>
        /// Set width of the table to given percentage 
        /// </summary>
        /// <param name="percentage"></param>
        public void SetWidthPercentage(int percentage) {
            if (percentage > 100) {
                percentage = 100;
            } else if (percentage < 0) {
                percentage = 0;
            }
            this.Width = percentage * 50;
            this.WidthType = TableWidthUnitValues.Pct;
        }
    }
}
