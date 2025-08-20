using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private static void RemoveCustomStyle(string styleId) {
            var field = typeof(WordParagraphStyle).GetField("_customStyles", BindingFlags.NonPublic | BindingFlags.Static);
            var dict = (IDictionary<string, Style>)field!.GetValue(null);
            dict.Remove(styleId);
        }

        [Fact]
        public void Test_FluentParagraphBuilderStylePersistence() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentParagraphBuilder.docx");
            string customStyleId = "MyStyle";
            var style = WordParagraphStyle.CreateFontStyle(customStyleId, "Arial");
            WordParagraphStyle.RegisterCustomStyle(customStyleId, style);

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Heading").Style(WordParagraphStyles.Heading1))
                    .Paragraph(p => p.Text("Custom style").Style(customStyleId))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(WordParagraphStyles.Heading1, document.Paragraphs[0].Style);
                Assert.Equal(customStyleId, document.Paragraphs[1].StyleId);
            }

            RemoveCustomStyle(customStyleId);
        }

        [Fact]
        public void Test_FluentParagraphBuilderNewMethods() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentParagraphBuilderNewMethods.docx");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Before").Tab().Text("After"))
                    .Paragraph(p => p.Link("https://example.com", "Example", true))
                    .Paragraph(p => p.Text("Line1").Break().Text("Line2"))
                    .Paragraph(p => p.Align(HorizontalAlignment.Right).Text("Aligned"))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[1].IsTab);
                Assert.True(document.Paragraphs[3].IsHyperLink);
                Assert.Equal("https://example.com/", document.Paragraphs[3].Hyperlink?.Uri?.ToString());
                Assert.True(document.Paragraphs[5].IsBreak);
                Assert.Equal(JustificationValues.Right, document.Paragraphs.Last().ParagraphAlignment);
            }
        }
    }
}
