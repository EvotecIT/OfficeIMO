using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
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
