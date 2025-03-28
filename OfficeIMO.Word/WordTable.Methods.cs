using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        /// <summary>
        /// Add comment to a Table
        /// </summary>
        /// <param name="author"></param>
        /// <param name="initials"></param>
        /// <param name="comment"></param>
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
    }
}
