using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

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

                using PowerPointPresentation exported = PowerPointPresentation.Open(exportedPath);
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
    }
}
