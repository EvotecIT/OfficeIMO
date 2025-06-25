using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Adds revision-related utilities.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Inserts revision text marked as added.
        /// </summary>
        /// <param name="text">Text to insert.</param>
        /// <param name="author">Revision author.</param>
        /// <param name="date">Revision date. Uses current date when null.</param>
        /// <returns>The current <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddInsertedText(string text, string author, DateTime? date = null) {
            VerifyRun();
            date ??= DateTime.Now;
            var run = new Run();
            run.RsidRunAddition = WordHeadersAndFooters.GenerateRsid();
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var ins = new InsertedRun() { Author = author, Date = date.Value, Id = WordHeadersAndFooters.GenerateRevisionId() };
            ins.Append(run);
            _paragraph.Append(ins);
            return this;
        }

        /// <summary>
        /// Inserts revision text marked as deleted.
        /// </summary>
        /// <param name="text">Text to delete.</param>
        /// <param name="author">Revision author.</param>
        /// <param name="date">Revision date. Uses current date when null.</param>
        /// <returns>The current <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddDeletedText(string text, string author, DateTime? date = null) {
            VerifyRun();
            date ??= DateTime.Now;
            var run = new Run();
            run.RsidRunDeletion = WordHeadersAndFooters.GenerateRsid();
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var del = new DeletedRun() { Author = author, Date = date.Value, Id = WordHeadersAndFooters.GenerateRevisionId() };
            del.Append(run);
            _paragraph.Append(del);
            return this;
        }
    }
}
