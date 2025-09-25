using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioFluentDocumentTests {
        [Fact]
        public void CanBuildDocumentFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            VisioDocument result = document.AsFluent()
                .Page("Page1", 8.5, 11, VisioMeasurementUnit.Inches, p => { })
                .End();

            Assert.Same(document, result);
            Assert.Single(document.Pages);
            Assert.Equal("Page1", document.Pages[0].Name);
            document.Save();
        }
    }
}
