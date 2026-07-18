using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointExportTests {
        [Fact]
        public void CanExportSingleSlideAsStandalonePresentation() {
            string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string exportedPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide first = presentation.AddSlide();
                    first.AddTitle("Keep me");
                    PowerPointSlide second = presentation.AddSlide();
                    second.AddTitle("Export me");
                    second.AddSmartArt().SetNodeText(0, "Preview boundary");
                    presentation.Save();

                    presentation.ExportSlide(1, exportedPath);
                }

                using PowerPointPresentation exported = PowerPointPresentation.Load(exportedPath);
                Assert.Single(exported.Slides);
                Assert.Contains(exported.Slides[0].TextBoxes, box => box.Text == "Export me");
                PowerPointSmartArt smartArt = Assert.Single(exported.Slides[0].SmartArts);
                Assert.Equal("Preview boundary", smartArt.GetNodeText(0));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
                if (File.Exists(exportedPath)) {
                    File.Delete(exportedPath);
                }
            }
        }

        [Fact]
        public void ExportSlide_OmitsInternallyLinkedTargetSlides() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide exportedSource = source.AddSlide();
            PowerPointTextRun run = exportedSource.AddTextBox("Target")
                .Paragraphs.Single().Runs.Single();
            PowerPointSlide linkedTarget = source.AddSlide();
            linkedTarget.AddTitle("Do not export");
            run.SetHyperlink(linkedTarget);
            using var destination = new MemoryStream();

            source.ExportSlide(0, destination);

            destination.Position = 0;
            using PowerPointPresentation exported =
                PowerPointPresentation.Load(destination);
            Assert.Single(exported.Slides);
            Assert.Empty(exported.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Single(exported.OpenXmlDocument.PresentationPart!
                .SlideParts);
            Assert.Empty(exported.ValidateDocument());
        }
    }
}
