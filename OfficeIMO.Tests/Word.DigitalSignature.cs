using System.IO;
using DocumentFormat.OpenXml.ExtendedProperties;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_MissingPart_ReturnsNull() {
            string tempFile = Path.GetTempFileName();
            using (WordDocument document = WordDocument.Create(tempFile)) {
                Assert.True(document.ApplicationProperties.DigitalSignature == null);
            }
        }

        [Fact]
        public void Test_DigitalSignature_PartDeleted_ReturnsNull() {
            string tempFile = Path.GetTempFileName();
            using (WordDocument document = WordDocument.Create(tempFile)) {
                document.ApplicationProperties.DigitalSignature = new DigitalSignature();
                Assert.True(document.ApplicationProperties.DigitalSignature != null);
                var extendedPart = document._wordprocessingDocument!.ExtendedFilePropertiesPart;
                Assert.NotNull(extendedPart);
                document._wordprocessingDocument!.DeletePart(extendedPart);
                Assert.True(document.ApplicationProperties.DigitalSignature == null);
            }
        }
    }
}
