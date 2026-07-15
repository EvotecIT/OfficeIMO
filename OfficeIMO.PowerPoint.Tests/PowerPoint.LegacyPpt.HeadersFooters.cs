using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptHeaderFooterTests {
        private static string FixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "BasicPowerPoint.ppt");

        [Fact]
        public void BinaryCorpus_DecodesAndProjectsHeaderFooterScopes() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);
            LegacyPptHeaderFooterSettings slideDefaults = Assert.IsType<LegacyPptHeaderFooterSettings>(
                legacy.SlideHeaderFooterDefaults);
            LegacyPptHeaderFooterSettings notesDefaults = Assert.IsType<LegacyPptHeaderFooterSettings>(
                legacy.NotesHeaderFooterDefaults);
            Assert.InRange(slideDefaults.DateTimeFormatId, (short)0, (short)13);
            Assert.InRange(notesDefaults.DateTimeFormatId, (short)0, (short)13);
            Assert.True(legacy.CreateImportReport().HeaderFooterScopeCount >= 2);

            LegacyPptMaster mainMaster = legacy.Masters.First(master => master.IsMainMaster);
            LegacyPptHeaderFooterSettings effective = mainMaster.HeaderFooter
                ?? slideDefaults;
            using PowerPointPresentation projected = PowerPointPresentation.Load(FixturePath);
            SlideMaster projectedMaster = projected.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.First().SlideMaster!;
            HeaderFooter masterHeaderFooter = Assert.IsType<HeaderFooter>(
                projectedMaster.GetFirstChild<HeaderFooter>());
            Assert.Equal(effective.ShowDate, masterHeaderFooter.DateTime?.Value);
            Assert.Equal(effective.ShowFooter, masterHeaderFooter.Footer?.Value);
            Assert.Equal(effective.ShowSlideNumber,
                masterHeaderFooter.SlideNumber?.Value);

            NotesMaster notesMaster = projected.OpenXmlDocument.PresentationPart!
                .NotesMasterPart!.NotesMaster!;
            HeaderFooter notesHeaderFooter = Assert.IsType<HeaderFooter>(
                notesMaster.GetFirstChild<HeaderFooter>());
            Assert.Equal(notesDefaults.ShowHeader, notesHeaderFooter.Header?.Value);
            Assert.Equal(notesDefaults.ShowFooter, notesHeaderFooter.Footer?.Value);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_AuthorsPerSlideFooterDateAndNumberSettings() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide();
                source.EnsureLayoutFooterPlaceholderTextBox(text: "Confidential");
                source.EnsureLayoutDateTimePlaceholderTextBox(text: "15 July 2026");
                source.EnsureLayoutSlideNumberPlaceholderTextBox(text: "1");

                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptHeaderFooterSettings settings = Assert.IsType<LegacyPptHeaderFooterSettings>(
                Assert.Single(legacy.Slides).HeaderFooter);
            Assert.True(settings.ShowDate);
            Assert.True(settings.UseUserDate);
            Assert.False(settings.UseAutomaticDateTime);
            Assert.True(settings.ShowFooter);
            Assert.True(settings.ShowSlideNumber);
            Assert.Equal("15 July 2026", settings.UserDateText);
            Assert.Equal("Confidential", settings.FooterText);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            SlideLayout layout = projected.Slides[0].SlidePart.SlideLayoutPart!.SlideLayout!;
            HeaderFooter headerFooter = Assert.IsType<HeaderFooter>(
                layout.GetFirstChild<HeaderFooter>());
            Assert.True(headerFooter.DateTime?.Value);
            Assert.True(headerFooter.Footer?.Value);
            Assert.True(headerFooter.SlideNumber?.Value);
            Assert.Contains(layout.CommonSlideData!.ShapeTree!.Elements<Shape>(), shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape?.Type?.Value == PlaceholderValues.Footer
                && string.Concat(shape.TextBody!.Descendants<A.Text>()
                    .Select(text => text.Text)) == "Confidential");
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_AuthorsDocumentAndNotesHeaderFooterDefaults() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.Notes.Text = "Speaker note";
                SlideMaster slideMaster = source.OpenXmlDocument.PresentationPart!
                    .SlideMasterParts.First().SlideMaster!;
                slideMaster.Append(new HeaderFooter {
                    DateTime = true,
                    Footer = true,
                    Header = false,
                    SlideNumber = true
                });
                NotesMaster notesMaster = source.OpenXmlDocument.PresentationPart!
                    .NotesMasterPart!.NotesMaster!;
                notesMaster.Append(new HeaderFooter {
                    DateTime = true,
                    Footer = true,
                    Header = true,
                    SlideNumber = true
                });

                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptHeaderFooterSettings slideDefaults = Assert.IsType<LegacyPptHeaderFooterSettings>(
                legacy.SlideHeaderFooterDefaults);
            Assert.True(slideDefaults.ShowDate);
            Assert.True(slideDefaults.ShowFooter);
            Assert.True(slideDefaults.ShowSlideNumber);
            Assert.False(slideDefaults.ShowHeader);
            LegacyPptHeaderFooterSettings notesDefaults = Assert.IsType<LegacyPptHeaderFooterSettings>(
                legacy.NotesHeaderFooterDefaults);
            Assert.True(notesDefaults.ShowDate);
            Assert.True(notesDefaults.ShowFooter);
            Assert.True(notesDefaults.ShowHeader);
            Assert.True(notesDefaults.ShowSlideNumber);
        }

        [Fact]
        public void ImportedPerSlideHeaderFooterEdit_AppendsPreservingRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide();
                source.EnsureLayoutFooterPlaceholderTextBox(text: "Original footer");
                source.EnsureLayoutSlideNumberPlaceholderTextBox(text: "1");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                SlideLayout layout = imported.Slides[0].SlidePart.SlideLayoutPart!.SlideLayout!;
                HeaderFooter headerFooter = layout.GetFirstChild<HeaderFooter>()!;
                headerFooter.Footer = true;
                headerFooter.SlideNumber = false;
                Shape footer = layout.CommonSlideData!.ShapeTree!.Elements<Shape>()
                    .First(shape => shape.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                        .Type?.Value == PlaceholderValues.Footer);
                A.Text text = footer.TextBody!.Descendants<A.Text>().First();
                text.Text = "Updated footer with more text";

                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptHeaderFooterSettings settings = Assert.IsType<LegacyPptHeaderFooterSettings>(
                Assert.Single(saved.Slides).HeaderFooter);
            Assert.True(settings.ShowFooter);
            Assert.False(settings.ShowSlideNumber);
            Assert.Equal("Updated footer with more text", settings.FooterText);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }
    }
}
