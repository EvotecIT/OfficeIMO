using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPresentationSizes {
        [Fact]
        public void SlideAndNotesSizesAreSet() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                Assert.NotNull(document.PresentationPart?.Presentation?.SlideSize);
                Assert.NotNull(document.PresentationPart?.Presentation?.NotesSize);
            }

            File.Delete(filePath);
        }
    }
}