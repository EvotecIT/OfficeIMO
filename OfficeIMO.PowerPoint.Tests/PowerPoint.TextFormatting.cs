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
        public void DrawingFontSizeAuthoringEnforcesSchemaBoundsAcrossTextSurfaces() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox box = slide.AddTextBox("Text");
            PowerPointTextRun run = box.Paragraphs.Single().Runs.Single();
            PowerPointTableCell cell = slide.AddTable(1,1).GetCell(0, 0);

            run.FontSize = global::OfficeIMO.OfficeLength.Points(1D);
            Assert.Equal(1D, run.FontSize?.ToPoints());
            run.FontSize = global::OfficeIMO.OfficeLength.Points(4000D);
            Assert.Equal(4000D, run.FontSize?.ToPoints());

            Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSize = global::OfficeIMO.OfficeLength.Points(0.5D));
            Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSize = global::OfficeIMO.OfficeLength.Points(4000.01D));
            Assert.Throws<ArgumentOutOfRangeException>(() => box.FontSize = global::OfficeIMO.OfficeLength.Points(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => box.FontSize = global::OfficeIMO.OfficeLength.Points(4001));
            Assert.Throws<ArgumentOutOfRangeException>(() => cell.FontSize = global::OfficeIMO.OfficeLength.Points(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => cell.FontSize = global::OfficeIMO.OfficeLength.Points(4001));
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(4000), run.FontSize);
            run.FontSize = global::OfficeIMO.OfficeLength.Points(14.505);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(14.51), run.FontSize);
            Assert.Equal(1451, run.RunProperties!.FontSize!.Value);
            run.FontSize = null;
            Assert.Null(run.FontSize);
            Assert.Null(run.RunProperties.FontSize);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void CanApplyFormattingToTextBoxAndBullets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Hello");
                box.Bold = true;
                box.Italic = true;
                box.FontSize = global::OfficeIMO.OfficeLength.Points(24);
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

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
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

            using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                PowerPointSlide slide = presentation.Slides.First();
                PowerPointTextBox textBox = slide.TextBoxes.First();
                Assert.Equal(string.Empty, textBox.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void TextBoxTextHandlesSingleParagraph() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
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

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
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

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Highlights");
                box.AddBullets(new[] { "First", "Second" });
                box.ApplyTextStyle(new PowerPointTextStyle(fontSize: global::OfficeIMO.OfficeLength.Points(20), bold: true, color: "336699"));
                box.ApplyParagraphStyle(new PowerPointParagraphStyle(
                    lineSpacingMultiplier: 1.2,
                    spaceAfter: global::OfficeIMO.OfficeLength.Points(4),
                    leftMargin: global::OfficeIMO.OfficeLength.Points(18),
                    indent: global::OfficeIMO.OfficeLength.Points(-18)));
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
                Assert.Equal(global::OfficeIMO.OfficeLength.Points(18).ToEmus(), pp.LeftMargin!.Value);
                Assert.Equal(global::OfficeIMO.OfficeLength.Points(-18).ToEmus(), pp.Indent!.Value);
                Assert.Equal(120000, pp.LineSpacing!.SpacingPercent!.Val!.Value);
                Assert.Equal(400, pp.SpaceAfter!.SpacingPoints!.Val!.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void ClearPreservesValidTextBodyStructure() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Initial");

                box.Clear();
                presentation.Save();

                Assert.Empty(presentation.ValidateDocument());
                Assert.Equal(string.Empty, box.Text);
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                Shape shape = document.PresentationPart!.SlideParts.First().Slide.Descendants<Shape>().First();
                Assert.Single(shape.TextBody!.Elements<A.Paragraph>());
            }

            File.Delete(filePath);
        }

        [Fact]
        public void SetParagraphs_ReplacesContentWithoutLeadingBlankParagraph() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Initial");

                box.SetParagraphs(new[] { "First", "Second" });

                Assert.Equal("First" + Environment.NewLine + "Second", box.Text);
                Assert.Equal(2, box.Paragraphs.Count);
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                Shape shape = document.PresentationPart!.SlideParts.First().Slide.Descendants<Shape>().First();
                var paragraphs = shape.TextBody!.Elements<A.Paragraph>().ToList();
                Assert.Equal(2, paragraphs.Count);
                Assert.Equal("First", paragraphs[0].InnerText);
                Assert.Equal("Second", paragraphs[1].InnerText);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void SetBulletsAndNumberedList_DoNotLeaveLeadingBlankParagraph() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox bullets = slide.AddTextBox("Initial bullets");
                PowerPointTextBox numbered = slide.AddTextBox("Initial numbered");

                bullets.SetBullets(new[] { "Bullet A", "Bullet B" });
                numbered.SetNumberedList(new[] { "One", "Two" });

                Assert.Equal(2, bullets.Paragraphs.Count);
                Assert.Equal(2, numbered.Paragraphs.Count);
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                var shapes = document.PresentationPart!.SlideParts.First().Slide.Descendants<Shape>().ToList();

                var bulletParagraphs = shapes[0].TextBody!.Elements<A.Paragraph>().ToList();
                Assert.Equal(2, bulletParagraphs.Count);
                Assert.Equal("Bullet A", bulletParagraphs[0].InnerText);
                Assert.Equal("Bullet B", bulletParagraphs[1].InnerText);

                var numberedParagraphs = shapes[1].TextBody!.Elements<A.Paragraph>().ToList();
                Assert.Equal(2, numberedParagraphs.Count);
                Assert.Equal("One", numberedParagraphs[0].InnerText);
                Assert.Equal("Two", numberedParagraphs[1].InnerText);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void BulletSizingAndFontStayInSchemaOrder() {
            string filePath = CreateTempFilePath(".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Bullets");
                box.SetBullets(new[] { "First", "Second" }, configure: paragraph => {
                    paragraph.SetBulletSizePercent(72)
                        .SetBulletFont("Aptos")
                        .SetHanging(global::OfficeIMO.OfficeLength.Points(16))
                        .SetFontSize(global::OfficeIMO.OfficeLength.Points(18));
                });

                presentation.Save();

                Assert.Empty(presentation.ValidateDocument());
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                Shape shape = document.PresentationPart!.SlideParts.First().Slide.Descendants<Shape>().First();
                A.ParagraphProperties props = shape.TextBody!.Elements<A.Paragraph>().First().ParagraphProperties!;
                var children = props.ChildElements.ToList();

                int sizeIndex = IndexOf<A.BulletSizePercentage>(children);
                int fontIndex = IndexOf<A.BulletFont>(children);
                int bulletIndex = IndexOf<A.CharacterBullet>(children);

                Assert.NotEqual(-1, sizeIndex);
                Assert.NotEqual(-1, fontIndex);
                Assert.NotEqual(-1, bulletIndex);
                Assert.True(sizeIndex < fontIndex);
                Assert.True(fontIndex < bulletIndex);
            }

            File.Delete(filePath);

            static int IndexOf<T>(IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> children)
                where T : DocumentFormat.OpenXml.OpenXmlElement {
                return children.ToList().FindIndex(child => child is T);
            }
        }

        [Fact]
        public void TextGeometryRetainsNativeUnitsAndAtomicHangingIndent() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox box = slide.AddTextBox("Geometry");
            PowerPointParagraph paragraph = box.Paragraphs.Single();
            PowerPointTableCell cell = slide.AddTable(1, 1).GetCell(0, 0);
            paragraph.Indent = global::OfficeIMO.OfficeLength.Emus(-12701);
            paragraph.LeftMargin = global::OfficeIMO.OfficeLength.Emus(25401);
            paragraph.SpaceBefore = global::OfficeIMO.OfficeLength.Emus(63);
            paragraph.SpaceAfter = global::OfficeIMO.OfficeLength.Emus(64);
            Assert.Equal(-12701, paragraph.Paragraph.ParagraphProperties!.Indent!.Value);
            Assert.Equal(25401, paragraph.Paragraph.ParagraphProperties.LeftMargin!.Value);
            Assert.Equal(0, paragraph.Paragraph.ParagraphProperties.SpaceBefore!.SpacingPoints!.Val!.Value);
            Assert.Equal(1, paragraph.Paragraph.ParagraphProperties.SpaceAfter!.SpacingPoints!.Val!.Value);
            paragraph.LineSpacingMultiplier = 1.25;
            Assert.Null(paragraph.LineSpacing);
            paragraph.LineSpacing = global::OfficeIMO.OfficeLength.Points(18);
            Assert.Null(paragraph.LineSpacingMultiplier);
            paragraph.BulletSize = global::OfficeIMO.OfficeLength.Points(11.5);
            Assert.Equal(1150, paragraph.Paragraph.ParagraphProperties.GetFirstChild<A.BulletSizePoints>()!.Val!.Value);
            paragraph.BulletSizePercent = 72;
            Assert.Null(paragraph.BulletSize);
            Assert.Equal(72, paragraph.BulletSizePercent);
            paragraph.BulletSize = null;
            Assert.Null(paragraph.BulletSizePercent);
            string beforeFailure = paragraph.Paragraph.OuterXml;
            Assert.Throws<OverflowException>(() => paragraph.SetHanging(global::OfficeIMO.OfficeLength.Emus(int.MinValue)));
            Assert.Equal(beforeFailure, paragraph.Paragraph.OuterXml);
            box.FontSize = global::OfficeIMO.OfficeLength.Inches(0.25);
            box.TextMarginLeft = global::OfficeIMO.OfficeLength.Emus(91441);
            cell.PaddingTop = global::OfficeIMO.OfficeLength.Emus(12701);
            using var output = new MemoryStream();
            presentation.Save(output);
            output.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(output);
            PowerPointTextBox restored = reopened.Slides[0].TextBoxes.Single();
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(18), restored.FontSize);
            Assert.Equal(global::OfficeIMO.OfficeLength.Emus(91441), restored.TextMarginLeft);
            Assert.Equal(global::OfficeIMO.OfficeLength.Emus(-12701), restored.Paragraphs[0].Indent);
            Assert.Equal(global::OfficeIMO.OfficeLength.Emus(127), restored.Paragraphs[0].SpaceAfter);
            Assert.Equal(global::OfficeIMO.OfficeLength.Emus(12701), reopened.Slides[0].Tables.Single().GetCell(0, 0).PaddingTop);
            restored.Paragraphs[0].Indent = null;
            restored.TextMarginLeft = null;
            Assert.Null(restored.Paragraphs[0].Indent);
            Assert.Null(restored.TextMarginLeft);
        }

        private static string CreateTempFilePath(string extension) {
            string path = Path.GetTempFileName();
            File.Delete(path);
            return Path.ChangeExtension(path, extension);
        }
    }
}
