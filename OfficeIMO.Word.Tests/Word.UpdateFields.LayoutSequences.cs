using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        // Word 16 updates these fields in document order to 1, 2, 3, 4, 4, 5, (hidden 6), 7 (checked with Fields.Update):
        // the reset takes the page number of page 3, and the repeat, next and hidden fields continue from it. After the
        // next Heading 1, \s 1 starts again at 1.
        [Fact]
        public void Test_LayoutPaginationKeepsSequenceFieldsInDocumentOrder() {
            using WordDocument document = WordDocument.Create();
            AddSequenceParagraph(document, "A", " SEQ Figure ");
            AddSequenceParagraph(document, "B", " SEQ Figure ");
            document.AddPageBreak();
            document.AddPageBreak();
            AddSequenceParagraph(document, "C", " SEQ Figure \\r ", nestedPage: true);
            AddSequenceParagraph(document, "D", " SEQ Figure ");
            AddSequenceParagraph(document, "E", " SEQ Figure \\c ");
            AddSequenceParagraph(document, "F", " SEQ Figure \\n ");
            AddSequenceParagraph(document, "G", " SEQ Figure \\h ");
            AddSequenceParagraph(document, "H", " SEQ Figure ");
            document.AddParagraph("Next chapter").Style = WordParagraphStyles.Heading1;
            AddSequenceParagraph(document, "I", " SEQ Figure \\s 1 ");
            AddSequenceParagraph(document, "J", " SEQ Figure ");

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions().WithPdfLayoutPagination());

            string[] sequence = report.Results.Where(result => result.FieldType == WordFieldType.Seq)
                .OrderBy(result => result.Index).Select(result => result.ResultText!).ToArray();
            Assert.Equal(new[] { "1", "2", "3", "4", "4", "5" }, sequence.Take(6));
            Assert.Equal(new[] { "7", "1", "2" }, sequence.Skip(sequence.Length - 3));
        }

        private static void AddSequenceParagraph(WordDocument document, string label, string instruction, bool nestedPage = false) {
            WordParagraph paragraph = document.AddParagraph(label + " ");
            Paragraph xml = paragraph._paragraph;
            xml.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            xml.Append(new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }));
            if (nestedPage) {
                xml.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                xml.Append(new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }));
                xml.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                xml.Append(new Run(new Text("1")));
                xml.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            }
            xml.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            xml.Append(new Run(new Text("0")));
            xml.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }
    }
}
