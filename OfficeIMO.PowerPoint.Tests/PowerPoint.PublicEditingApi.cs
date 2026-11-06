using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPublicEditingApi {
        [Fact]
        public void CreateEditSaveAndReopenUseConcretePresentationObjects() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide first = presentation.AddSlide();
                    first.AddTitle("Canonical title");
                    PowerPointTextBox bullets = first.AddTextBox(string.Empty);
                    bullets.AddBullet("One");
                    bullets.AddBullet("Two");
                    first.AddTable(2,2);
                    first.Notes.Text = "Notes text";

                    PowerPointSlide second = presentation.AddSlide();
                    second.AddTitle("Second slide");
                    presentation.MoveSlide(1, 0);
                    presentation.Save();
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    Assert.Equal(2, presentation.Slides.Count);
                    Assert.Equal("Second slide", presentation.Slides[0].TextBoxes.First().Text);
                    Assert.Equal("Notes text", presentation.Slides[1].Notes.Text);
                    Assert.Single(presentation.Slides[1].Tables);
                    presentation.Slides[1].ReplaceAll("Canonical", "Edited");
                    presentation.Save();
                }

                using PowerPointDocument reopened = PowerPointDocument.Load(filePath,
                    new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
                Assert.Contains(reopened.Slides[1].TextBoxes, textBox => textBox.Text == "Edited title");
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void ReplaceAllIgnoresCaseByDefaultAndHonoursMatchCaseAcrossTextboxesAndTables() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox box = slide.AddTextBox("Draft and DRAFT");
            PowerPointTableCell cell = slide.AddTable(1, 1).GetCell(0, 0);
            cell.Text = "draft";

            Assert.Equal(1, presentation.ReplaceAll("draft", "Final", new PowerPointTextSearchOptions { MatchCase = true }));
            Assert.Equal("Draft and DRAFT", box.Text);
            Assert.Equal("Final", cell.Text);

            Assert.Equal(2, presentation.ReplaceAll("draft", "Final"));
            Assert.Equal("Final and Final", box.Text);
        }

        [Fact]
        public void ReplaceAllSavesSplitRunMatchesWithFormattingAndInlineBoundaries() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box = slide.AddTextBox("Hello ");
                PowerPointParagraph paragraph = box.Paragraphs[0];
                paragraph.AddFormattedText("{{na", bold: true);
                paragraph.AddFormattedText("me}} tail", italic: true);

                PowerPointParagraph withBreak = box.AddParagraph("{{na");
                withBreak.AddLineBreak();
                withBreak.AddText("me}}");

                PowerPointParagraph withField = box.AddParagraph("{{na");
                withField.AddField("1", "slidenum");
                withField.AddText("me}}");

                PowerPointParagraph withinWord = box.AddParagraph("prefix{{na");
                withinWord.AddText("me}}");

                PowerPointTableCell cell = slide.AddTable(1, 1).GetCell(0, 0);
                cell.Text = "{{NA";
                cell.AddText("ME}}");

                Assert.Equal(2, presentation.ReplaceAll("{{name}}", "Alice",
                    new PowerPointTextSearchOptions { WholeWords = true }));
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            PowerPointSlide edited = Assert.Single(reopened.Slides);
            PowerPointTextBox editedBox = Assert.Single(edited.TextBoxes);
            Assert.Collection(editedBox.Paragraphs[0].Runs,
                run => Assert.Equal("Hello ", run.Text),
                run => {
                    Assert.Equal("Alice", run.Text);
                    Assert.True(run.Bold);
                    Assert.False(run.Italic);
                },
                run => {
                    Assert.Equal(" tail", run.Text);
                    Assert.True(run.Italic);
                    Assert.False(run.Bold);
                });
            Assert.Equal("{{na" + Environment.NewLine + "me}}", editedBox.Paragraphs[1].Text);
            Assert.Equal(PowerPointParagraphInlineKind.LineBreak, editedBox.Paragraphs[1].InlineNodes[1].Kind);
            Assert.Equal("{{na1me}}", editedBox.Paragraphs[2].Text);
            Assert.Equal("slidenum", editedBox.Paragraphs[2].InlineNodes[1].FieldType);
            Assert.Equal("prefix{{name}}", editedBox.Paragraphs[3].Text);
            Assert.Equal("Alice", Assert.Single(edited.Tables).GetCell(0, 0).Text);
        }
    }
}
