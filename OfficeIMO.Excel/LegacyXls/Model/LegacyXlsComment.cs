namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a legacy XLS cell comment parsed from NOTE/OBJ/TXO records.
    /// </summary>
    public sealed class LegacyXlsComment {
        /// <summary>
        /// Creates parsed legacy XLS comment metadata.
        /// </summary>
        public LegacyXlsComment(
            int row,
            int column,
            string text,
            string author,
            ushort objectId,
            bool visible,
            IReadOnlyList<LegacyXlsCommentFormattingRun>? formattingRuns = null) {
            Row = row;
            Column = column;
            Text = text;
            Author = author;
            ObjectId = objectId;
            Visible = visible;
            FormattingRuns = formattingRuns ?? Array.Empty<LegacyXlsCommentFormattingRun>();
        }

        /// <summary>Gets the 1-based row containing the comment.</summary>
        public int Row { get; }

        /// <summary>Gets the 1-based column containing the comment.</summary>
        public int Column { get; }

        /// <summary>Gets the plain comment text.</summary>
        public string Text { get; }

        /// <summary>Gets the comment author.</summary>
        public string Author { get; }

        /// <summary>Gets the BIFF object id associated with the comment.</summary>
        public ushort ObjectId { get; }

        /// <summary>Gets whether the source comment was configured as always visible.</summary>
        public bool Visible { get; }

        /// <summary>Gets rich text run boundaries and font indexes from the source TxO records.</summary>
        public IReadOnlyList<LegacyXlsCommentFormattingRun> FormattingRuns { get; }
    }
}
