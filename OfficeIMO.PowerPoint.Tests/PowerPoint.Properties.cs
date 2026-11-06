using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPropertiesTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void SavingTitlePreservesCoreAuthorAndCreationProperties(bool hasCoreProperties) {
            string? creator = hasCoreProperties ? "Original creator" : null;
            string? lastModifiedBy = hasCoreProperties ? "Original editor" : null;
            var originalModified = new DateTime(2001, 1, 2, 0, 0, 0, DateTimeKind.Utc);
            DateTime? created = hasCoreProperties ? originalModified.AddDays(-1) : null;

            using var source = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create()) {
                presentation.AddSlide().AddTitle("Original slide");
                presentation.Save(source);
            }

            source.Position = 0;
            using (PresentationDocument package = PresentationDocument.Open(source, true)) {
                if (hasCoreProperties) {
                    package.PackageProperties.Creator = creator;
                    package.PackageProperties.LastModifiedBy = lastModifiedBy;
                    package.PackageProperties.Created = created;
                    package.PackageProperties.Modified = originalModified;
                } else {
                    package.DeletePart(package.CoreFilePropertiesPart!);
                }
            }

            using var input = new MemoryStream(source.ToArray(), writable: false);
            using var output = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Load(input)) {
                Assert.Equal(creator, presentation.Properties.Creator);
                presentation.Properties.Title = "Updated title";
                presentation.Save(output);
            }

            output.Position = 0;
            using PresentationDocument saved = PresentationDocument.Open(output, false);
            Assert.Equal("Updated title", saved.PackageProperties.Title);
            Assert.Equal(creator, saved.PackageProperties.Creator);
            Assert.Equal(lastModifiedBy, saved.PackageProperties.LastModifiedBy);
            Assert.Equal(created, saved.PackageProperties.Created?.ToUniversalTime());
            Assert.True(saved.PackageProperties.Modified?.ToUniversalTime() > originalModified);
        }

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
