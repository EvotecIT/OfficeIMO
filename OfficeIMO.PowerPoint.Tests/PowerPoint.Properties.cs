using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPropertiesTests {
        [Fact]
        public void CanRoundTripBuiltinAndApplicationProperties() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.BuiltinDocumentProperties.Title = "Roundtrip deck";
                    presentation.BuiltinDocumentProperties.Creator = "OfficeIMO Tests";
                    presentation.BuiltinDocumentProperties.Keywords = "officeimo,powerpoint";
                    presentation.BuiltinDocumentProperties.Subject = "Properties";
                    presentation.BuiltinDocumentProperties.Category = "Testing";
                    presentation.ApplicationProperties.Company = "Evotec";
                    presentation.ApplicationProperties.Manager = "Automation";
                    presentation.ApplicationProperties.Application = "OfficeIMO.PowerPoint";
                    presentation.ApplicationProperties.ApplicationVersion = "1.0";

                    presentation.AddSlide().AddTitle("Properties");
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    Assert.Equal("Roundtrip deck", presentation.BuiltinDocumentProperties.Title);
                    Assert.Equal("OfficeIMO Tests", presentation.BuiltinDocumentProperties.Creator);
                    Assert.Equal("officeimo,powerpoint", presentation.BuiltinDocumentProperties.Keywords);
                    Assert.Equal("Properties", presentation.BuiltinDocumentProperties.Subject);
                    Assert.Equal("Testing", presentation.BuiltinDocumentProperties.Category);
                    Assert.Equal("Evotec", presentation.ApplicationProperties.Company);
                    Assert.Equal("Automation", presentation.ApplicationProperties.Manager);
                    Assert.Equal("OfficeIMO.PowerPoint", presentation.ApplicationProperties.Application);
                    Assert.Equal("1.0", presentation.ApplicationProperties.ApplicationVersion);
                    Assert.Equal("1", presentation.ApplicationProperties.Slides);
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    Assert.Equal("Roundtrip deck", document.PackageProperties.Title);
                    Assert.Equal("Evotec", document.ExtendedFilePropertiesPart!.Properties!.Company!.Text);
                    Assert.Equal("Automation", document.ExtendedFilePropertiesPart.Properties.Manager!.Text);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
