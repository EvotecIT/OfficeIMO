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
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    presentation.Properties.Title = "Roundtrip deck";
                    presentation.Properties.Creator = "OfficeIMO Tests";
                    presentation.Properties.Keywords = "officeimo,powerpoint";
                    presentation.Properties.Subject = "Properties";
                    presentation.Properties.Category = "Testing";
                    presentation.ApplicationProperties.Company = "Evotec";
                    presentation.ApplicationProperties.Manager = "Automation";
                    presentation.ApplicationProperties.Application = "OfficeIMO.PowerPoint";
                    presentation.ApplicationProperties.ApplicationVersion = "1.0";

                    presentation.AddSlide().AddTitle("Properties");
                    presentation.Save();
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    Assert.Equal("Roundtrip deck", presentation.Properties.Title);
                    Assert.Equal("OfficeIMO Tests", presentation.Properties.Creator);
                    Assert.Equal("officeimo,powerpoint", presentation.Properties.Keywords);
                    Assert.Equal("Properties", presentation.Properties.Subject);
                    Assert.Equal("Testing", presentation.Properties.Category);
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
