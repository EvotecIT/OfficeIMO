using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Helper;
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

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx"))) {
                Assert.True(document.Settings.Language == "pl-PL");

                document.Settings.Language = "en-US";

                Assert.True(document.Settings.Language == "en-US");

                Assert.True(document.Settings.ProtectionType == DocumentProtectionValues.ReadOnly);

                document.Settings.RemoveProtection();



                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx"))) {
                Assert.True(document.Settings.ProtectionType == null);
                
                document.Save();
            }
        }
    }
}
