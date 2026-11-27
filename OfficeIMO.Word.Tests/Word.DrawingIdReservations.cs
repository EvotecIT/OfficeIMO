using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Drawing ids reserved after the package was handed out: parts nobody loaded are skipped once they're known to hold
    /// no ids, and ids added to them through the package afterwards are still never reused.
    /// </summary>
    public class WordDrawingIdReservations {
        [Fact]
        public void IdsAddedToASkippedPartAfterwardsAreNotReused() {
            using WordDocument document = WordDocument.Create();
            WordprocessingDocument package = document.OpenXmlDocument;
            FootnotesPart footnotes = package.MainDocumentPart!.FootnotesPart!;
            document.AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(20));
            Assert.False(footnotes.IsRootElementLoaded);

            footnotes.Footnotes!.AppendChild(new Footnote(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                new DW.Inline(new DW.DocProperties { Id = 500U, Name = "Added through the package" }))))) { Id = 5 });
            document.AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(20));

            uint[] ids = package.MainDocumentPart.Document!.Descendants<DW.DocProperties>().Select(properties => properties.Id!.Value).ToArray();
            Assert.Equal(2, ids.Length);
            Assert.True(ids[1] > 500U, $"The second shape got id {ids[1]}, which the footnote already uses or precedes.");
        }

#if NET
        [Fact]
        [Trait("Category", "Performance")]
        public void ALargePartWithoutIdsIsReadOnceAcrossReservations() {
            using WordDocument document = WordDocument.Create();
            WordprocessingDocument package = document.OpenXmlDocument;
            WordprocessingCommentsPart comments = package.MainDocumentPart!.WordprocessingCommentsPart
                ?? package.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            var markup = new StringBuilder("<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
            for (int index = 0; markup.Length < 16 * 1024 * 1024; index++) {
                markup.Append("<w:comment w:id=\"").Append(index).Append("\" w:author=\"A\"><w:p><w:r><w:t>Comment text without drawings</w:t></w:r></w:p></w:comment>");
            }
            markup.Append("</w:comments>");
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(markup.ToString()))) comments.FeedData(stream);
            markup.Clear();

            document.AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(20));
            long before = GC.GetAllocatedBytesForCurrentThread();
            for (int index = 0; index < 3; index++) {
                document.AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(20));
            }
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

            Assert.False(comments.IsRootElementLoaded);
            Assert.True(allocated < 4 * 1024 * 1024, $"Three more reservations allocated {allocated:N0} bytes.");
        }
#endif
    }
}
