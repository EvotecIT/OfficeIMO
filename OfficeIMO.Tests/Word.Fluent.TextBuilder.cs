using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentTextBuilderFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTextBuilder.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Hello")
                        .Text(" World", t => t.BoldOn().ItalicOn().Color("#ff0000"))
                        .Text("!", t => t.BoldOn()))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs[0];
                var runs = paragraph.GetRuns().ToList();
                Assert.Equal(3, runs.Count);
                Assert.Equal("Hello", runs[0].Text);
                Assert.Equal(" World", runs[1].Text);
                Assert.True(runs[1].Bold);
                Assert.True(runs[1].Italic);
                Assert.Equal("ff0000", runs[1].ColorHex);
                Assert.Equal("!", runs[2].Text);
                Assert.True(runs[2].Bold);
            }
        }

        [Fact]
        public void Test_FluentTextBuilderAdditionalFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTextBuilderAdvanced.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p
                        .Text("Underline", t => t.Underline(UnderlineValues.Double))
                        .Text(" Strike", t => t.Strike())
                        .Text(" DoubleStrike", t => t.DoubleStrike())
                        .Text(" FontSize", t => t.FontSize(20))
                        .Text(" FontFamily", t => t.FontFamily("Arial"))
                        .Text(" Highlight", t => t.Highlight(HighlightColorValues.Yellow))
                        .Text(" Sub", t => t.SubScript())
                        .Text(" Super", t => t.SuperScript())
                        .Text(" Caps", t => t.CapsStyle(CapsStyle.Caps)))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var runs = document.Paragraphs[0].GetRuns().ToList();
                Assert.Equal(9, runs.Count);
                Assert.Equal(UnderlineValues.Double, runs[0].Underline);
                Assert.True(runs[1].Strike);
                Assert.True(runs[2].DoubleStrike);
                Assert.Equal(20, runs[3].FontSize);
                Assert.Equal("Arial", runs[4].FontFamily);
                Assert.Equal(HighlightColorValues.Yellow, runs[5].Highlight);
                Assert.Equal(VerticalPositionValues.Subscript, runs[6].VerticalTextAlignment);
                Assert.Equal(VerticalPositionValues.Superscript, runs[7].VerticalTextAlignment);
                Assert.Equal(CapsStyle.Caps, runs[8].CapsStyle);
            }
        }
    }
}