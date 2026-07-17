using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointExtractionTests {
        [Fact]
        public void ExtractMarkdownChunks_IncludesTablesAndAllNotesText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTitle("Quarterly Review");
                    slide.AddTextBox("Agenda");

                    PowerPointTable table = slide.AddTable(2, 2);
                    table.GetRow(0).GetCell(0).Text = "Region";
                    table.GetRow(0).GetCell(1).Text = "Revenue";
                    table.GetRow(1).GetCell(0).Text = "North";
                    table.GetRow(1).GetCell(1).Text = "120";

                    slide.Notes.Text = "Presenter reminder";
                    presentation.Save();
                }

                AppendNotesParagraph(filePath, "Second note paragraph");

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointExtractChunk chunk = presentation.ExtractMarkdownChunks(
                        sourcePath: filePath).Single();

                    Assert.Contains("## Slide 1", chunk.Markdown);
                    Assert.Contains("Agenda", chunk.Markdown);
                    Assert.Contains("### Table 1", chunk.Markdown);
                    Assert.Contains("| Region | Revenue |", chunk.Markdown);
                    Assert.Contains("| North | 120 |", chunk.Markdown);
                    Assert.Contains("### Notes", chunk.Markdown);
                    Assert.Contains("Presenter reminder", chunk.Markdown);
                    Assert.Contains("Second note paragraph", chunk.Markdown);
                    Assert.Null(chunk.Warnings);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ExtractMarkdownChunks_RespectsOptionsAndAddsTruncationWarnings() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTextBox(new string('A', 400));

                    PowerPointTable table = slide.AddTable(2, 1);
                    table.GetRow(0).GetCell(0).Text = "Header";
                    table.GetRow(1).GetCell(0).Text = "Value";

                    slide.Notes.Text = "Private note";
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointExtractChunk chunk = presentation.ExtractMarkdownChunks(
                        new PowerPointExtractionExtensions.PowerPointExtractOptions {
                            IncludeNotes = false,
                            IncludeTables = false
                        },
                        new PowerPointExtractChunkingOptions { MaxChars = 256 },
                        sourcePath: filePath).Single();

                    Assert.DoesNotContain("### Table", chunk.Markdown);
                    Assert.DoesNotContain("### Notes", chunk.Markdown);
                    Assert.Contains("<!-- truncated -->", chunk.Markdown);
                    Assert.NotNull(chunk.Warnings);
                    Assert.Contains("Markdown truncated to MaxChars.", chunk.Warnings!);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ExtractMarkdownChunks_PreservesParagraphSemanticsLinksAndShapeOrder() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTitle("Extraction contract");

            PowerPointTextBox list = slide.AddTextBox(string.Empty);
            list.SetBullets(new[] { "Parent bullet" });
            list.AddBullets(new[] { "Nested bullet" }, level: 1);
            list.AddNumberedList(new[] { "Third item", "Fourth item" },
                A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenR,
                startAt: 3);
            PowerPointParagraph closing = list.AddParagraph("Closing paragraph");
            closing.ClearBullet();

            PowerPointTextBox prose = slide.AddTextBox(string.Empty);
            prose.SetParagraphs(new[] { "First paragraph", "Second paragraph" });

            PowerPointTextBox wideNumber = slide.AddTextBox(string.Empty);
            wideNumber.SetNumberedList(new[] { "Hundredth item" },
                A.TextAutoNumberSchemeValues.ArabicPeriod, startAt: 100);
            wideNumber.AddBullets(new[] { "Arrow child" }, level: 1,
                bulletChar: '\u2192');

            PowerPointTable table = slide.AddTable(2, 2);
            table.GetCell(0, 0).Text = "Region";
            table.GetCell(0, 1).Text = "Revenue";
            table.GetCell(1, 0).Text = "North";
            table.GetCell(1, 1).Text = "120";

            PowerPointTextBox afterTable = slide.AddTextBox("Read ");
            afterTable.Paragraphs[0].AddRun("the guide")
                .SetHyperlink("https://example.test/guide");
            afterTable.Paragraphs[0].AddRun(" and ");
            afterTable.Paragraphs[0].AddRun("roadmap")
                .SetHyperlink("Quarter 1.pptx");
            slide.Notes.Text = "Speaker reminder";

            string markdown = presentation.ExtractMarkdownChunks().Single().Markdown;

            Assert.DoesNotContain("\r", markdown, StringComparison.Ordinal);
            Assert.Contains("### Extraction contract", markdown, StringComparison.Ordinal);
            Assert.Contains("- • Parent bullet", markdown, StringComparison.Ordinal);
            Assert.Contains("    - • Nested bullet", markdown, StringComparison.Ordinal);
            Assert.Contains("3. c) Third item", markdown, StringComparison.Ordinal);
            Assert.Contains("4. d) Fourth item", markdown, StringComparison.Ordinal);
            Assert.Contains("4. d) Fourth item\n\nClosing paragraph", markdown, StringComparison.Ordinal);
            Assert.Contains("First paragraph\n\nSecond paragraph", markdown, StringComparison.Ordinal);
            Assert.Contains("100. Hundredth item\n     - → Arrow child",
                markdown, StringComparison.Ordinal);
            Assert.Contains("| Region | Revenue |", markdown, StringComparison.Ordinal);
            Assert.Contains("Read [the guide](<https://example.test/guide>)", markdown, StringComparison.Ordinal);
            Assert.Contains("[roadmap](<Quarter%201.pptx>)", markdown, StringComparison.Ordinal);
            Assert.Contains("### Notes", markdown, StringComparison.Ordinal);
            Assert.Contains("Speaker reminder", markdown, StringComparison.Ordinal);

            int listPosition = markdown.IndexOf("- • Parent bullet", StringComparison.Ordinal);
            int tablePosition = markdown.IndexOf("### Table 1", StringComparison.Ordinal);
            int afterTablePosition = markdown.IndexOf("Read [the guide]", StringComparison.Ordinal);
            Assert.True(listPosition >= 0 && listPosition < tablePosition,
                "The list textbox should be emitted before the table.");
            Assert.True(tablePosition < afterTablePosition,
                "The table should be emitted before the following textbox.");
        }

        [Fact]
        public void ExtractMarkdownChunks_RestartsNumberingAtListBoundaries() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox text = slide.AddTextBox(string.Empty);
            text.SetNumberedList(new[] { "Before one", "Before two" });
            text.AddParagraph("Plain boundary").ClearBullet();
            PowerPointParagraph restarted = text.AddParagraph("After one");
            restarted.SetNumbered(A.TextAutoNumberSchemeValues.ArabicPeriod);
            PowerPointParagraph parentOne = text.AddParagraph("Parent one");
            parentOne.SetNumbered(A.TextAutoNumberSchemeValues.ArabicPeriod);
            PowerPointParagraph childOne = text.AddParagraph("Child one");
            childOne.SetNumbered(A.TextAutoNumberSchemeValues.ArabicPeriod);
            childOne.Level = 1;
            PowerPointParagraph parentTwo = text.AddParagraph("Parent two");
            parentTwo.SetNumbered(A.TextAutoNumberSchemeValues.ArabicPeriod);
            PowerPointParagraph childTwo = text.AddParagraph("Child restart");
            childTwo.SetNumbered(A.TextAutoNumberSchemeValues.ArabicPeriod);
            childTwo.Level = 1;

            string markdown = presentation.ExtractMarkdownChunks()
                .Single().Markdown;

            Assert.Contains("2. Before two\n\nPlain boundary\n\n1. After one",
                markdown, StringComparison.Ordinal);
            Assert.Contains("2. Parent one\n    1. Child one\n3. Parent two\n    1. Child restart",
                markdown, StringComparison.Ordinal);
        }

        private static void AppendNotesParagraph(string filePath, string text) {
            using PresentationDocument document = PresentationDocument.Open(filePath, true);
            NotesSlidePart notesPart = document.PresentationPart!.SlideParts.First().NotesSlidePart!;
            Shape notesShape = notesPart.NotesSlide!.CommonSlideData!.ShapeTree!.Elements<Shape>().First();
            var textBody = notesShape.TextBody!;
            textBody.Append(new A.Paragraph(
                new A.Run(new A.Text(text)),
                new A.EndParagraphRunProperties { Language = "en-US" }));
            notesPart.NotesSlide.Save();
        }
    }
}
