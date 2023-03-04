using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithTabStops() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithTabStops.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("\tFirst Line");

                Assert.True(document.Paragraphs.Count == 1);
                Assert.True(paragraph.TabStops.Count == 0);

                var tab1 = paragraph.AddTabStop(1440);

                Assert.True(paragraph.TabStops.Count == 1);

                var tab2 = paragraph.AddTabStop(1440);
                tab2.Alignment = TabStopValues.Left;
                tab2.Leader = TabStopLeaderCharValues.Hyphen;
                tab2.Position = 1440;

                paragraph.AddText("\tMore text");

                Assert.True(paragraph.TabStops.Count == 2);

                var paragraph1 = document.AddParagraph("\tNext Line");

                var tab3 = paragraph1.AddTabStop(5000);
                tab3.Leader = TabStopLeaderCharValues.Hyphen;

                var tab4 = paragraph1.AddTabStop(1440 * 2);
                paragraph1.AddText("\tEven more text");


                var tab5 = paragraph1.AddTabStop(1440 * 3, TabStopValues.Decimal, TabStopLeaderCharValues.MiddleDot);
                paragraph1.AddText("\tLast more text");

                Assert.True(paragraph.TabStops.Count == 2);
                Assert.True(paragraph1.TabStops.Count == 3);

                Assert.True(document.Paragraphs.Count == 5);
                // First paragraph with 2 runs, having 2 tab stops
                Assert.True(document.Paragraphs[0].TabStops.Count == 2);
                Assert.True(document.Paragraphs[1].TabStops.Count == 2);
                // Actual new paragraph with 3 runs, having 3 tab stops
                Assert.True(document.Paragraphs[2].TabStops.Count == 3);
                Assert.True(document.Paragraphs[3].TabStops.Count == 3);
                Assert.True(document.Paragraphs[4].TabStops.Count == 3);

                // two WordParagraphs, share same Paragraph, and same ParagraphProperties, so the tab stops are shared
                Assert.True(document.Paragraphs[0].TabStops[0].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[0].TabStops[0].Leader == TabStopLeaderCharValues.None);
                Assert.True(document.Paragraphs[0].TabStops[0].Position == 1440);

                Assert.True(document.Paragraphs[0].TabStops[1].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[0].TabStops[1].Leader == TabStopLeaderCharValues.Hyphen);
                Assert.True(document.Paragraphs[0].TabStops[1].Position == 1440);

                // three WordParagraphs, share same Paragraph, and same ParagraphProperties, so the tab stops are shared
                Assert.True(document.Paragraphs[2].TabStops[0].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[2].TabStops[0].Leader == TabStopLeaderCharValues.Hyphen);
                Assert.True(document.Paragraphs[2].TabStops[0].Position == 5000);

                Assert.True(document.Paragraphs[3].TabStops[1].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[3].TabStops[1].Leader == TabStopLeaderCharValues.None);
                Assert.True(document.Paragraphs[3].TabStops[1].Position == 2880);

                Assert.True(document.Paragraphs[4].TabStops[2].Alignment == TabStopValues.Decimal);
                Assert.True(document.Paragraphs[4].TabStops[2].Leader == TabStopLeaderCharValues.MiddleDot);
                Assert.True(document.Paragraphs[4].TabStops[2].Position == 1440 * 3);

                Assert.True(document.Sections[0].Paragraphs[0].TabStops.Count == 2);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTabStops.docx"))) {
                Assert.True(document.Paragraphs.Count == 5);
                // First paragraph with 2 runs, having 2 tab stops
                Assert.True(document.Paragraphs[0].TabStops.Count == 2);
                Assert.True(document.Paragraphs[1].TabStops.Count == 2);
                // Actual new paragraph with 3 runs, having 3 tab stops
                Assert.True(document.Paragraphs[2].TabStops.Count == 3);
                Assert.True(document.Paragraphs[3].TabStops.Count == 3);
                Assert.True(document.Paragraphs[4].TabStops.Count == 3);

                // two WordParagraphs, share same Paragraph, and same ParagraphProperties, so the tab stops are shared
                Assert.True(document.Paragraphs[0].TabStops[0].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[0].TabStops[0].Leader == TabStopLeaderCharValues.None);
                Assert.True(document.Paragraphs[0].TabStops[0].Position == 1440);

                Assert.True(document.Paragraphs[0].TabStops[1].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[0].TabStops[1].Leader == TabStopLeaderCharValues.Hyphen);
                Assert.True(document.Paragraphs[0].TabStops[1].Position == 1440);

                // three WordParagraphs, share same Paragraph, and same ParagraphProperties, so the tab stops are shared
                Assert.True(document.Paragraphs[2].TabStops[0].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[2].TabStops[0].Leader == TabStopLeaderCharValues.Hyphen);
                Assert.True(document.Paragraphs[2].TabStops[0].Position == 5000);

                Assert.True(document.Paragraphs[3].TabStops[1].Alignment == TabStopValues.Left);
                Assert.True(document.Paragraphs[3].TabStops[1].Leader == TabStopLeaderCharValues.None);
                Assert.True(document.Paragraphs[3].TabStops[1].Position == 2880);

                Assert.True(document.Paragraphs[4].TabStops[2].Alignment == TabStopValues.Decimal);
                Assert.True(document.Paragraphs[4].TabStops[2].Leader == TabStopLeaderCharValues.MiddleDot);
                Assert.True(document.Paragraphs[4].TabStops[2].Position == 1440 * 3);

                Assert.True(document.Sections[0].Paragraphs[0].TabStops.Count == 2);
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTabStops.docx"))) {

            }
        }
    }
}
