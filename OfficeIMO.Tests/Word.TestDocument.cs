using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_OpeningWordCreatedInOffice365() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "TestDocument365.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 6,"Paragraphs not matching. Current: " + document.Paragraphs.Count);

                // There is only one PageBreak in this document.
                Assert.True(document.PageBreaks.Count() == 1,"Page breaks not matching. Current: " + document.PageBreaks.Count);


                Assert.True(document.Sections[0].Paragraphs.Count == 6, "Paragraphs not matching. Current: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Sections not matching. Current: " + document.Sections.Count);

                Assert.True(document.PageOrientation == PageOrientationValues.Portrait, "Page orientation. Current: " + document.PageOrientation);
                //Assert.True(document.Settings == PageOrientationValues.Portrait, "Page orientation. Current: " + document.PageOrientation);


                // TODO add revisions to accept and check those really got accepted
                document.AcceptRevisions("Przemysław Kłys");

                document.AcceptRevisions();
            }
        }
    }
}
