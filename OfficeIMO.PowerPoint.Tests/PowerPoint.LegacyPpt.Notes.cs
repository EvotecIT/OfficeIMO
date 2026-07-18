using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptNotesTests {
        private static string FixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "BasicPowerPoint.ppt");

        [Fact]
        public void BinaryImport_ResolvesNotesIdThroughNotesDirectory() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);
            LegacyPptSlide[] slidesWithNotes = legacy.Slides
                .Where(slide => slide.NotesId != 0)
                .ToArray();

            Assert.NotEmpty(slidesWithNotes);
            Assert.All(slidesWithNotes, slide => {
                LegacyPptNotesPage page = Assert.IsType<LegacyPptNotesPage>(slide.NotesPage);
                Assert.Equal(slide.NotesId, page.NotesId);
                Assert.Equal(slide.SlideId, page.SlideId);
                Assert.NotEqual(0U, page.PersistId);
                Assert.NotEmpty(page.Shapes);
                Assert.Equal(page.Text, slide.NotesText);
            });
            LegacyPptImportReport report = legacy.CreateImportReport();
            Assert.Equal(slidesWithNotes.Length, report.NotesSlideCount);
            Assert.Equal(slidesWithNotes.Sum(slide => slide.NotesPage!.Shapes.Count),
                report.NotesPageShapeCount);
        }

        [Fact]
        public void NativeWriter_AuthorsAndProjectsEditableSpeakerNotes() {
            const string expected = "First note line\nSecond note line";
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.Notes.Text = expected;
                slide.SlidePart.NotesSlidePart!.NotesSlide!
                    .ShowMasterShapes = false;

                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                Assert.DoesNotContain(source.AnalyzeLegacyPptWrite().Findings,
                    finding => finding.Code == "PPT-WRITE-NOTES");
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptSlide binarySlide = Assert.Single(legacy.Slides);
            LegacyPptNotesPage page = Assert.IsType<LegacyPptNotesPage>(
                binarySlide.NotesPage);
            Assert.Equal(256U, binarySlide.NotesId);
            Assert.Equal(15U, page.PersistId);
            Assert.Equal(binarySlide.SlideId, page.SlideId);
            Assert.Equal(expected, Normalize(page.Text));
            Assert.False(page.FollowsMasterObjects);
            Assert.Contains(page.Shapes,
                shape => shape.PlaceholderKind == LegacyPptPlaceholderKind.NotesBody);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            Assert.Equal(expected, Normalize(projected.Slides[0].Notes.Text));
            Assert.False(projected.Slides[0].SlidePart.NotesSlidePart!
                .NotesSlide!.ShowMasterShapes!.Value);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void ImportedBinaryNotesEdit_AppendsPreservingIncrementalRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.Notes.Text = "Original speaker note";
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                imported.Slides[0].Notes.Text =
                    "Updated speaker note with a longer body";
                imported.Slides[0].SlidePart.NotesSlidePart!.NotesSlide!
                    .ShowMasterShapes = false;
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Equal("Updated speaker note with a longer body",
                saved.Slides[0].NotesText);
            Assert.False(saved.Slides[0].NotesPage!.FollowsMasterObjects);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes,
                writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.False(reopened.Slides[0].SlidePart.NotesSlidePart!
                .NotesSlide!.ShowMasterShapes!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static string Normalize(string value) => (value ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
    }
}
