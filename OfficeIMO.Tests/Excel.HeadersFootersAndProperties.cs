using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
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

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public void Excel_HeaderImage_Roundtrips_ContentType() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderImageContentType.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var sheet = doc.AddWorkSheet("Sheet1");
                var pngPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
                var pngBytes = File.ReadAllBytes(pngPath);
                sheet.SetHeaderImage(HeaderFooterPosition.Center, pngBytes, "image/png");
                doc.Save(false);
            }

            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                var sheetPart = package.WorkbookPart!.WorksheetParts.First();
                var vmlPart = sheetPart.VmlDrawingParts.FirstOrDefault();
                Assert.NotNull(vmlPart);
                var imagePart = Assert.Single(vmlPart!.ImageParts);
                Assert.Equal("image/png", imagePart.ContentType);
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public void Excel_FooterImage_Roundtrips_ContentType() {
            string filePath = Path.Combine(_directoryWithFiles, "FooterImageContentType.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var sheet = doc.AddWorkSheet("Sheet1");
                var jpegPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                var jpegBytes = File.ReadAllBytes(jpegPath);
                sheet.SetFooterImage(HeaderFooterPosition.Center, jpegBytes, "image/jpeg");
                doc.Save(false);
            }

            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                var sheetPart = package.WorkbookPart!.WorksheetParts.First();
                var vmlPart = sheetPart.VmlDrawingParts.FirstOrDefault();
                Assert.NotNull(vmlPart);
                var imagePart = Assert.Single(vmlPart!.ImageParts);
                Assert.Equal("image/jpeg", imagePart.ContentType);
            }
        }
    }
}
