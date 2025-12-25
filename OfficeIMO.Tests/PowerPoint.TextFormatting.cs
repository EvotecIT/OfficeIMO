using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTextFormatting {
        [Fact]
        public void CanApplyFormattingToTextBoxAndBullets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Hello");
                box.Bold = true;
                box.Italic = true;
                box.FontSize = 24;
                box.FontName = "Arial";
                box.Color = "FF0000";
                box.AddBullet("Bullet1");
                box.AddBullet("Bullet2");
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                Shape shape = slidePart.Slide.Descendants<Shape>().First();
                var paragraphs = shape.TextBody!.Elements<A.Paragraph>().ToList();
                    foreach (var paragraph in paragraphs) {
                        A.Run run = paragraph.GetFirstChild<A.Run>()!;
                        A.RunProperties rp = run.RunProperties!;
                        Assert.True(rp.Bold?.Value ?? false);
                        Assert.True(rp.Italic?.Value ?? false);
                        Assert.Equal(2400, rp.FontSize!.Value);
                        Assert.Equal("Arial", rp.GetFirstChild<A.LatinFont>()?.Typeface);
                        Assert.Equal("FF0000", rp.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val);
                    }
            }

            File.Delete(filePath);
        }

        [Fact]
        public void TextBoxTextReturnsEmptyWhenParagraphsAreMissing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("Initial");
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                Shape shape = slidePart.Slide.Descendants<Shape>().First();
                shape.TextBody!.RemoveAllChildren<A.Paragraph>();
                slidePart.Slide.Save();
                document.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides.First();
                PowerPointTextBox textBox = slide.TextBoxes.First();
                Assert.Equal(string.Empty, textBox.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void TextBoxTextHandlesSingleParagraph() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox textBox = slide.AddTextBox("Initial");

                textBox.Text = "Updated text";
                Assert.Equal("Updated text", textBox.Text);

                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                Shape shape = slidePart.Slide.Descendants<Shape>().First();
                var paragraphs = shape.TextBody!.Elements<A.Paragraph>().ToList();
                Assert.Single(paragraphs);
                Assert.Equal("Updated text", paragraphs[0].InnerText);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void TextBoxTextHandlesMultipleParagraphs() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string[] lines = { "First line", "Second line", "Third line" };
            string text = string.Join(Environment.NewLine, lines);

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox textBox = slide.AddTextBox(string.Empty);

                textBox.Text = text;
                Assert.Equal(text, textBox.Text);

                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                Shape shape = slidePart.Slide.Descendants<Shape>().First();
                var paragraphs = shape.TextBody!.Elements<A.Paragraph>().ToList();
                Assert.Equal(lines.Length, paragraphs.Count);
                for (int i = 0; i < lines.Length; i++) {
                    Assert.Equal(lines[i], paragraphs[i].InnerText);
                }
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanApplyTextAndParagraphStyles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Highlights");
                box.AddBullets(new[] { "First", "Second" });
                box.ApplyTextStyle(new PowerPointTextStyle(fontSize: 20, bold: true, color: "336699"));
                box.ApplyParagraphStyle(new PowerPointParagraphStyle(
                    lineSpacingMultiplier: 1.2,
                    spaceAfterPoints: 4,
                    leftMarginPoints: 18,
                    indentPoints: -18));
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                Shape shape = slidePart.Slide.Descendants<Shape>().First();
                A.Paragraph paragraph = shape.TextBody!.Elements<A.Paragraph>().First();
                A.Run run = paragraph.GetFirstChild<A.Run>()!;
                A.RunProperties rp = run.RunProperties!;

                Assert.Equal(2000, rp.FontSize!.Value);
                Assert.True(rp.Bold?.Value ?? false);
                Assert.Equal("336699", rp.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val);

                A.ParagraphProperties pp = paragraph.ParagraphProperties!;
                Assert.Equal(1800, pp.LeftMargin!.Value);
                Assert.Equal(-1800, pp.Indent!.Value);
                Assert.Equal(120000, pp.LineSpacing!.SpacingPercent!.Val!.Value);
                Assert.Equal(400, pp.SpaceAfter!.SpacingPoints!.Val!.Value);
            }

            File.Delete(filePath);
        }
    }
}
