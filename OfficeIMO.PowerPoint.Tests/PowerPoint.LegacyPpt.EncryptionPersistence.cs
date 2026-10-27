using System.Security.Cryptography;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointLegacyEncryptionPersistenceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedBinarySavesKeepTheCurrentDestinationPassword(bool removePassword) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-ppt-password-state-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string source = Path.Combine(directory, "source.ppt"), output = Path.Combine(directory, "output.ppt");
            using (var document = PowerPointPresentation.Create()) {
                document.AddSlide().AddTextBox("Encrypted source");
                document.Save(source, new PowerPointSaveOptions { Password = "first", LegacyPptEncryptionKeySizeBits = 40 });
            }
            using (var document = PowerPointPresentation.Load(source, new PowerPointLoadOptions { Password = "first" })) {
                document.Save(output, new PowerPointSaveOptions { Password = removePassword ? null : "second", LegacyPptEncryptionKeySizeBits = 128 });
                document.Slides[0].TextBoxes.Single(box => box.Text == "Encrypted source").Text = "Saved again";
                document.Save();
            }
            using var reopened = PowerPointPresentation.Load(output, removePassword ? null : new PowerPointLoadOptions { Password = "second" });
            Assert.Contains(reopened.Slides[0].TextBoxes, box => box.Text == "Saved again");
            if (!removePassword) {
                Assert.ThrowsAny<CryptographicException>(() => PowerPointPresentation.Load(output, new PowerPointLoadOptions { Password = "first" }));
                Assert.Equal(128, LegacyPptPresentation.Load(File.ReadAllBytes(output), new LegacyPptImportOptions { Password = "second" }).Package.EncryptionKeySizeBits);
            }
            else Assert.False(LegacyPptPresentation.Load(File.ReadAllBytes(output)).Package.WasEncryptedSource);
        } finally { Directory.Delete(directory, true); }
    }
}
