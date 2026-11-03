using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public class PowerPointTextLineBreakTests {
        private static readonly string[] BreakForms = { "\n", "\r\n", "\r", "\v" };

        private static readonly IReadOnlyDictionary<string, Func<PowerPointSlide, string, Func<string>>> Writers =
            new Dictionary<string, Func<PowerPointSlide, string, Func<string>>> {
                ["cell text"] = (slide, text) => {
                    PowerPointTableCell cell = Cell(slide);
                    cell.Text = text;
                    return () => cell.Text;
                },
                ["cell run"] = (slide, text) => {
                    PowerPointTextRun run = Cell(slide).AddText(text);
                    return () => run.Text;
                },
                ["cell paragraph"] = (slide, text) => {
                    PowerPointTableCell cell = Cell(slide);
                    cell.SetParagraphs(new[] { text });
                    return () => cell.Text;
                },
                ["text box"] = (slide, text) => {
                    PowerPointTextBox box = slide.AddTextBox(text);
                    return () => box.Text;
                },
                ["title"] = (slide, text) => {
                    PowerPointTextBox title = slide.AddTitle(text);
                    return () => title.Text;
                },
                ["auto shape"] = (slide, text) => {
                    PowerPointTextBox shape = slide.AddTextShape(OfficePresetShapeType.Rectangle, text);
                    return () => shape.Text;
                },
                ["added paragraph"] = (slide, text) => {
                    PowerPointParagraph paragraph = slide.AddTextBox("first").AddParagraph(text);
                    return () => paragraph.Text;
                },
                ["paragraph text"] = (slide, text) => {
                    PowerPointParagraph paragraph = slide.AddTextBox("first").Paragraphs[0];
                    paragraph.Text = text;
                    return () => paragraph.Text;
                },
                ["existing run text"] = (slide, text) => {
                    PowerPointTextRun run = slide.AddTextBox("first").Paragraphs[0].Runs[0];
                    run.Text = text;
                    return () => run.Text;
                },
                ["replaced text"] = (slide, text) => {
                    PowerPointTextBox box = slide.AddTextBox("token");
                    box.ReplaceAll("token", text);
                    return () => box.Text;
                },
                ["notes paragraph"] = (slide, text) => {
                    slide.Notes.SetParagraphs(new[] { text });
                    return () => slide.Notes.Text;
                }
            };

        public static IEnumerable<object[]> WriterAndBreak() =>
            from writer in Writers.Keys
            from lineBreak in BreakForms
            select new object[] { writer, lineBreak };

        [Theory]
        [MemberData(nameof(WriterAndBreak))]
        public void NativeText_SavesEveryBreakFormAsBreakBetweenRuns(string writer, string lineBreak) {
            using var stream = new MemoryStream();
            string readBack;
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                readBack = Writers[writer](presentation.AddSlide(), "One" + lineBreak + "Two")();
                presentation.Save();
            }

            A.Paragraph paragraph = Assert.Single(SavedParagraphs(stream), item => item.InnerText.Contains("One"));
            Assert.Equal(new[] { "One", "Two" }, paragraph.Elements<A.Run>().Select(run => run.Text!.Text).Where(line => line.Length > 0));
            Assert.Single(paragraph.Elements<A.Break>());
            Assert.Equal("One" + Environment.NewLine + "Two", readBack);
        }

        [Fact]
        public void CellFormatting_ReachesEveryLineAndSurvivesReassignment() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointTableCell cell = Cell(presentation.AddSlide());
                cell.Text = "One\vTwo";
                cell.FontSize = OfficeIMO.OfficeLength.Points(21);
                cell.Bold = true;
                cell.Italic = true;
                cell.FontFamily = "Arial";
                cell.Color = OfficeColor.Parse("112233");
                cell.Text = "Three\nFour\r\nFive";
                presentation.Save();
            }

            A.Paragraph paragraph = Assert.Single(SavedParagraphs(stream), item => item.InnerText.Contains("Three"));
            Assert.Equal(3, paragraph.Elements<A.Run>().Count());
            Assert.Equal(2, paragraph.Elements<A.Break>().Count());
            A.RunProperties[] properties = paragraph.Elements<A.Run>().Select(run => run.RunProperties!)
                .Concat(paragraph.Elements<A.Break>().Select(lineBreak => lineBreak.RunProperties!)).ToArray();
            Assert.All(properties, item => {
                Assert.Equal(2100, item.FontSize?.Value);
                Assert.True(item.Bold?.Value);
                Assert.True(item.Italic?.Value);
                Assert.Equal("Arial", item.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                Assert.Equal("112233", item.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
            });
        }

        [Fact]
        public void TextBoxFormatting_ReachesLineBreaks() {
            using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
            PowerPointTextBox box = presentation.AddSlide().AddTextBox("One\nTwo");

            box.FontSize = OfficeIMO.OfficeLength.Points(19);

            A.Paragraph paragraph = Assert.Single(box.Paragraphs).Paragraph;
            Assert.All(paragraph.Elements<A.Run>(), run => Assert.Equal(1900, run.RunProperties?.FontSize?.Value));
            Assert.Equal(1900, Assert.Single(paragraph.Elements<A.Break>()).RunProperties?.FontSize?.Value);
        }

        [Fact]
        public void RunWrapper_FormatsAllLinesItWrote() {
            using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
            PowerPointParagraph paragraph = presentation.AddSlide().AddTextBox("Lead ").Paragraphs[0];

            PowerPointTextRun run = paragraph.AddText("One\nTwo").SetBold();
            run.Color = OfficeColor.Parse("445566");

            Assert.False(paragraph.Runs[0].Bold);
            Assert.All(paragraph.Runs.Skip(1), line => {
                Assert.True(line.Bold);
                Assert.Equal(OfficeColor.Parse("445566"), line.Color);
            });
            A.Break lineBreak = Assert.Single(paragraph.Paragraph.Elements<A.Break>());
            Assert.True(lineBreak.RunProperties?.Bold?.Value);
        }

        [Fact]
        public void ParagraphFormatting_CoversMultiLineTextButNotEarlierExplicitBreakContent() {
            using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
            PowerPointTextBox box = presentation.AddSlide().AddTextBox("One\nTwo");
            box.ApplyTextStyle(new PowerPointTextStyle(bold: true));
            Assert.All(box.Paragraphs[0].Runs, run => Assert.True(run.Bold));

            PowerPointParagraph chained = box.AddParagraph(string.Empty);
            chained.AddText("Three");
            chained.AddLineBreak();
            chained.AddText("Four\nFive").SetItalic();
            Assert.False(chained.Runs.First(run => run.Text == "Three").Italic);
            Assert.All(chained.Runs.Where(run => run.Text is "Four" or "Five"), run => Assert.True(run.Italic));
        }

        [Fact]
        public void Hyperlink_StaysOnRunsAndLeavesBreaksPlain() {
            using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
            PowerPointParagraph paragraph = presentation.AddSlide().AddTextBox("x").Paragraphs[0];

            paragraph.AddText("One\nTwo").SetHyperlink("https://example.com/", "tip");

            Assert.All(paragraph.Paragraph.Elements<A.Run>().Skip(1), run =>
                Assert.NotNull(run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>()));
            Assert.Empty(paragraph.Paragraph.Elements<A.Break>().SelectMany(item => item.Descendants<A.HyperlinkType>()));
        }

        [Theory]
        [InlineData("\n")]
        [InlineData("\r\n")]
        [InlineData("\r")]
        public void TextBoxAndNotesText_KeepHardBreaksAsParagraphsAndSoftBreaksInside(string hardBreak) {
            string text = "One\vTwo" + hardBreak + "Three";
            using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox box = slide.AddTextBox("x");
            box.Text = text;
            slide.Notes.Text = text;

            string expected = string.Join(Environment.NewLine, "One", "Two", "Three");
            Assert.Equal(expected, box.Text);
            Assert.Equal(expected, slide.Notes.Text);
            Assert.Equal(2, box.Paragraphs.Count);
            Assert.Equal(2, slide.Notes.Paragraphs.Count);
            Assert.Single(box.Paragraphs[0].Paragraph.Elements<A.Break>());
            Assert.Single(slide.Notes.Paragraphs[0].Paragraph.Elements<A.Break>());
        }

        [Fact]
        public void SingleLineText_StaysOneRun() {
            using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTableCell cell = Cell(slide);
            cell.Text = "Only";

            A.Paragraph cellParagraph = cell.Paragraphs[0].Paragraph;
            A.Paragraph boxParagraph = slide.AddTextBox("Only").Paragraphs[0].Paragraph;
            Assert.Equal("Only", Assert.Single(cellParagraph.Elements<A.Run>()).Text?.Text);
            Assert.Equal("Only", Assert.Single(boxParagraph.Elements<A.Run>()).Text?.Text);
            Assert.Empty(cellParagraph.Elements<A.Break>());
            Assert.Empty(boxParagraph.Elements<A.Break>());
        }

        [Fact]
        public void InvalidXmlCharacters_AreDroppedSoTheDeckSaves() {
            using var stream = new MemoryStream();
            const string text = "A\u0000B\u0001C￿D\uD800E😀";
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                Cell(slide).Text = text;
                slide.AddTextBox(text);
                slide.Notes.Text = text;
                slide.AddTextBox("x").Paragraphs[0].AddField(text, "datetime1");
                presentation.Save();
            }

            A.Text[] written = SavedParagraphs(stream).SelectMany(item => item.Descendants<A.Text>())
                .Where(item => item.Text.Contains("ABCDE")).ToArray();
            Assert.Equal(4, written.Length);
            Assert.All(written, item => Assert.Equal("ABCDE😀", item.Text));
        }

        [Fact]
        public void FieldText_FlattensBreaksBecauseAFieldHoldsOneLine() {
            using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
            PowerPointParagraph paragraph = presentation.AddSlide().AddTextBox("x").Paragraphs[0];
            paragraph.AddField("One\vTwo", "datetime1");

            PowerPointTextRun field = paragraph.InlineNodes.Single(node => node.Kind == PowerPointParagraphInlineKind.Field).Run!;
            Assert.Equal("One Two", field.Text);
            field.Text = "Three\nFour";
            Assert.Equal("Three Four", field.Text);
        }

        [Fact]
        public void ChartTitle_WritesBreaksAndStylesEveryLine() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartType.ColumnClustered,
                    new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("S", new[] { 1D }) }));
                chart.SetTitle("One\vTwo");
                chart.SetTitleTextStyle(bold: true);
                presentation.Save();
            }

            stream.Position = 0;
            using PresentationDocument document = PresentationDocument.Open(stream, false);
            C.Title title = document.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!.Descendants<C.Title>().First();
            Assert.Equal(new[] { "One", "Two" }, title.Descendants<A.Run>().Select(run => run.Text!.Text));
            Assert.All(title.Descendants<A.Run>(), run => Assert.True(run.RunProperties?.Bold?.Value));
            Assert.Single(title.Descendants<A.Break>());
        }

        private static PowerPointTableCell Cell(PowerPointSlide slide) =>
            slide.AddTable(1, 1, OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20),
                OfficeIMO.OfficeLength.Points(200), OfficeIMO.OfficeLength.Points(100)).GetCell(0, 0);

        private static List<A.Paragraph> SavedParagraphs(MemoryStream stream) {
            stream.Position = 0;
            using PresentationDocument document = PresentationDocument.Open(stream, false);
            SlidePart slide = document.PresentationPart!.SlideParts.Single();
            IEnumerable<OpenXmlElement> roots = new OpenXmlElement[] { slide.Slide! };
            if (slide.NotesSlidePart != null) roots = roots.Append(slide.NotesSlidePart.NotesSlide!);
            return roots.SelectMany(root => root.Descendants<A.Paragraph>())
                .Select(paragraph => (A.Paragraph)paragraph.CloneNode(true)).ToList();
        }
    }
}
