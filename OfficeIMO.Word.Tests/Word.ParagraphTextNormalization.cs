using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private const char NonTextBreakPlaceholder = '\u2028';

        [Theory]
        [InlineData("Line1\rLine2", "Line1\nLine2")]
        [InlineData("Line1\r\nLine2\rLine3", "Line1\nLine2\nLine3")]
        [InlineData("\nStartsWith", "\nStartsWith")]
        [InlineData("EndsWith\r\n", "EndsWith\n")]
        [InlineData("Line1\n\nLine2", "Line1\n\nLine2")]
        public void ParagraphText_NormalizesNewLines(string input, string expected) {
            string filePath = Path.Combine(_directoryWithFiles, $"ParagraphTextNormalization_{Guid.NewGuid():N}.docx");

            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.Text = input;

                Assert.Equal(expected, paragraph.Text);
                Assert.Equal(expected, document.Paragraphs.Last().Text);

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(expected, document.Paragraphs.Last().Text);
            }
        }

        [Theory]
        [InlineData("Page")]
        [InlineData("Column")]
        public void ParagraphText_PreservesNonTextWrappingBreaksDuringReplacement(string breakTypeName) {
            BreakValues breakType = breakTypeName switch {
                "Page" => BreakValues.Page,
                "Column" => BreakValues.Column,
                _ => throw new ArgumentOutOfRangeException(nameof(breakTypeName), breakTypeName, "Unsupported break type")
            };
            string filePath = Path.Combine(_directoryWithFiles, $"ParagraphTextBreaks_{breakTypeName}_{Guid.NewGuid():N}.docx");

            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.Text = "Before";
                document.Save(false);
            }

            using (var wordprocessing = WordprocessingDocument.Open(filePath, true)) {
                var body = wordprocessing.MainDocumentPart!.Document.Body!;
                var run = body.Elements<Paragraph>().First().Elements<Run>().First();

                run.AppendChild(new Break { Type = breakType });
                run.AppendChild(new Text("After") { Space = SpaceProcessingModeValues.Preserve });

                wordprocessing.MainDocumentPart.Document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs.First();

                if (breakType == BreakValues.Page) {
                    Assert.NotNull(paragraph.PageBreak);
                } else {
                    Assert.NotNull(paragraph.Break);
                }

                Assert.Contains(NonTextBreakPlaceholder.ToString(), paragraph.Text);
                paragraph.Text = paragraph.Text.Replace("Before", "Updated");

                if (breakType == BreakValues.Page) {
                    Assert.NotNull(paragraph.PageBreak);
                } else {
                    Assert.NotNull(paragraph.Break);
                }

                Assert.Contains(NonTextBreakPlaceholder.ToString(), paragraph.Text);

                document.Save(false);
            }

            using (var wordprocessing = WordprocessingDocument.Open(filePath, false)) {
                var body = wordprocessing.MainDocumentPart!.Document.Body!;
                var run = body.Elements<Paragraph>().First().Elements<Run>().First();
                var elements = run.ChildElements.ToList();

                Assert.Equal(3, elements.Count);
                var firstText = Assert.IsType<Text>(elements[0]);
                var breakNode = Assert.IsType<Break>(elements[1]);
                var secondText = Assert.IsType<Text>(elements[2]);

                Assert.Equal(breakType, breakNode.Type?.Value);
                Assert.Equal("Updated", firstText.Text);
                Assert.Equal("After", secondText.Text);
            }
        }
    }
}
