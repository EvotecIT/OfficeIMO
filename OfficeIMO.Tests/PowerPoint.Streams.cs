using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointStreamTests {
        [Fact]
        public void Create_ToStream_WritesPackage() {
            using var stream = new MemoryStream();
            using (var presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide();
            }

            Assert.True(stream.Length > 0);
            stream.Position = 0;

            using var document = PresentationDocument.Open(stream, false);
            Assert.NotNull(document.PresentationPart);
            Assert.NotNull(document.PresentationPart!.Presentation);
            Assert.NotNull(document.PresentationPart.Presentation.SlideIdList);
            Assert.True(document.PresentationPart.Presentation.SlideIdList!.ChildElements.Count > 0);
        }
    }
}
