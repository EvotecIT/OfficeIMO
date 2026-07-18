using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointMediaCleanupTests {
        [Fact]
        public void RemovingSoundedSlideReleasesPackageMediaData() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.AddSlide().AddTitle("Retained");
            PowerPointSlide removed = presentation.AddSlide();
            PowerPointAutoShape shape = removed.AddRectangle(
                100000, 100000, 1000000, 500000);
            using (var actionSound = new MemoryStream(CreateWave(),
                       writable: false)) {
                shape.SetClickSound(actionSound, "Action");
            }
            removed.Transition = SlideTransition.Fade;
            using (var transitionSound = new MemoryStream(CreateWave(),
                       writable: false)) {
                removed.SetTransitionSound(transitionSound, "Transition");
            }
            removed.AddClassicAnimation(shape,
                PowerPointClassicAnimationEffect.Fade);
            using (var animationSound = new MemoryStream(CreateWave(),
                       writable: false)) {
                removed.SetClassicAnimationSound(shape, animationSound,
                    "Animation");
            }
            Assert.Equal(3, presentation.OpenXmlDocument.DataParts.Count());

            presentation.RemoveSlide(1);

            Assert.Single(presentation.Slides);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
            using var stream = new MemoryStream(presentation.ToBytes(),
                writable: false);
            using PresentationDocument package =
                PresentationDocument.Open(stream, false);
            Assert.Empty(package.DataParts);
        }

        [Fact]
        public void RemovingTableRowsAndColumnsReleasesDiscardedRunSounds() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTable(2, 2);
            PowerPointTextRun removedRowRun = table.GetCell(0, 0)
                .Paragraphs.Single().Runs.Single();
            using (var rowSound = new MemoryStream(CreateWave(),
                       writable: false)) {
                removedRowRun.SetClickSound(rowSound, "Row");
            }
            PowerPointTextRun removedColumnRun = table.GetCell(1, 1)
                .Paragraphs.Single().Runs.Single();
            using (var columnSound = new MemoryStream(CreateWave(),
                       writable: false)) {
                removedColumnRun.SetMouseOverSound(columnSound, "Column");
            }
            Assert.Equal(2, presentation.OpenXmlDocument.DataParts.Count());

            table.RemoveRow(0);

            Assert.Single(presentation.OpenXmlDocument.DataParts);
            Assert.Single(slide.SlidePart.DataPartReferenceRelationships);

            table.RemoveColumn(1);

            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships);
            Assert.Equal(1, table.Rows);
            Assert.Equal(1, table.Columns);
            Assert.Empty(presentation.ValidateDocument());
        }

        private static byte[] CreateWave() => new byte[] {
            0x52, 0x49, 0x46, 0x46, 0x24, 0x00, 0x00, 0x00,
            0x57, 0x41, 0x56, 0x45, 0x66, 0x6D, 0x74, 0x20,
            0x10, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00,
            0x40, 0x1F, 0x00, 0x00, 0x40, 0x1F, 0x00, 0x00,
            0x01, 0x00, 0x08, 0x00, 0x64, 0x61, 0x74, 0x61,
            0x00, 0x00, 0x00, 0x00
        };
    }
}
