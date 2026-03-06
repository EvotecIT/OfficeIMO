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

                using (PowerPointPresentation presentation = PowerPointPresentation.OpenRead(filePath)) {
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

                using (PowerPointPresentation presentation = PowerPointPresentation.OpenRead(filePath)) {
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
