using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = System.Drawing.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingDocumentWithSettings() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Settings.ProtectionPassword = "Test";

                Assert.True(document.Settings.ProtectionType == DocumentProtectionValues.ReadOnly);

                Assert.True(document.Settings.Language == "en-US");

                document.Settings.Language = "pl-PL";

                Assert.True(document.Settings.Language == "pl-PL");

                document.Settings.SetBackgroundColor(Color.BlueViolet);

                Assert.True(document.Settings.BackgroundColor == "8A2BE2");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx"))) {
                Assert.True(document.Settings.Language == "pl-PL");

                document.Settings.Language = "en-US";

                Assert.True(document.Settings.Language == "en-US");

                Assert.True(document.Settings.ProtectionType == DocumentProtectionValues.ReadOnly);

                document.Settings.RemoveProtection();

                Assert.True(document.Settings.BackgroundColor == "8A2BE2");

                document.Settings.SetBackgroundColor("FFA07A");


                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx"))) {
                Assert.True(document.Settings.ProtectionType == null);
                Assert.True(document.Settings.BackgroundColor == "FFA07A");
                document.Save();
            }
        }
    }
}
