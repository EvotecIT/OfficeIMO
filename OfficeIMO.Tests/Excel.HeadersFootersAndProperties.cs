using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Enums;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        [Trait("Category","ExcelPropsHF")]
        public void Excel_HeaderFooter_And_Properties_Roundtrip() {
            string filePath = Path.Combine(_directoryWithFiles, "PropsHeaderFooter.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                // Set workbook info via fluent
                doc.AsFluent().Info(i => i
                    .Title("Roundtrip Title")
                    .Author("Roundtrip Author")
                    .Company("Roundtrip Co")
                    .Application("OfficeIMO.Excel")
                    .Keywords("test,excel,header,footer")
                ).End();

                var sheet = doc.AddWorkSheet("Summary");
                sheet.Cell(1, 1, "Hello");

                // Header/footer with tokens and image
                string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
                byte[]? logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : null;
                sheet.SetHeaderFooter(headerCenter: "Domain Detective", headerRight: "Page &P of &N");
                if (logo != null) sheet.SetHeaderImage(HeaderFooterPosition.Center, logo, "image/png", widthPoints: 96, heightPoints: 32);

                // Save and close
                doc.Save(false);
            }

            // Reopen read-only and verify
            using (var verify = ExcelDocument.Load(filePath, readOnly: true))
            {
                Assert.Equal("Roundtrip Title", verify.BuiltinDocumentProperties.Title);
                Assert.Equal("Roundtrip Author", verify.BuiltinDocumentProperties.Creator);
                Assert.Equal("Roundtrip Co", verify.ApplicationProperties.Company);

                var summary = verify.Sheets.FirstOrDefault(s => s.Name == "Summary");
                Assert.NotNull(summary);
                var hf = summary!.GetHeaderFooter();
                // Header center must contain our text; right must include tokens
                Assert.Contains("Domain Detective", hf.HeaderCenter);
                Assert.Contains("&P", hf.HeaderRight);
                Assert.Contains("&N", hf.HeaderRight);
                Assert.True(hf.HeaderHasPicturePlaceholder);
            }
        }
    }
}
