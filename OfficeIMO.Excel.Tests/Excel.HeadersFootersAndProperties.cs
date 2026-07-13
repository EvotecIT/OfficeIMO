using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
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

                var sheet = doc.AddWorksheet("Summary");
                sheet.Cell(1, 1, "Hello");

                // Header/footer with tokens and image
                string logoPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
                byte[]? logo = File.Exists(logoPath) ? File.ReadAllBytes(logoPath) : null;
                sheet.SetHeaderFooter(headerCenter: "Domain Detective", headerRight: "Page &P of &N");
                if (logo != null) sheet.SetHeaderImage(HeaderFooterPosition.Center, logo, "image/png", widthPoints: 96, heightPoints: 32);

                // Save and close
                doc.Save();
            }

            // Reopen read-only and verify
            using (var verify = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
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
                Assert.Empty(verify.ValidateOpenXml());
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public void Excel_HeaderImage_Roundtrips_ContentType() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderImageContentType.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var sheet = doc.AddWorksheet("Sheet1");
                var pngPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
                var pngBytes = File.ReadAllBytes(pngPath);
                sheet.SetHeaderImage(HeaderFooterPosition.Center, pngBytes, "image/png");
                doc.Save();
            }

            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                var sheetPart = package.WorkbookPart!.WorksheetParts.First();
                var vmlPart = sheetPart.VmlDrawingParts.FirstOrDefault();
                Assert.NotNull(vmlPart);
                var imagePart = Assert.Single(vmlPart!.ImageParts);
                Assert.Equal("image/png", imagePart.ContentType);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public void Excel_HeaderImage_Normalizes_Known_ContentType_Alias() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderImageContentTypeAlias.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var sheet = doc.AddWorksheet("Sheet1");
                var jpegPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                var jpegBytes = File.ReadAllBytes(jpegPath);
                sheet.SetHeaderImage(HeaderFooterPosition.Center, jpegBytes, " image/jpg; charset=binary ");
                doc.Save();
            }

            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                var sheetPart = package.WorkbookPart!.WorksheetParts.First();
                var vmlPart = sheetPart.VmlDrawingParts.FirstOrDefault();
                Assert.NotNull(vmlPart);
                var imagePart = Assert.Single(vmlPart!.ImageParts);
                Assert.Equal(OfficeImageInfo.GetMimeType(OfficeImageFormat.Jpeg), imagePart.ContentType);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public void Excel_FooterImage_Roundtrips_ContentType() {
            string filePath = Path.Combine(_directoryWithFiles, "FooterImageContentType.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var sheet = doc.AddWorksheet("Sheet1");
                var jpegPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                var jpegBytes = File.ReadAllBytes(jpegPath);
                sheet.SetFooterImage(HeaderFooterPosition.Center, jpegBytes, "image/jpeg");
                doc.Save();
            }

            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                var sheetPart = package.WorkbookPart!.WorksheetParts.First();
                var vmlPart = sheetPart.VmlDrawingParts.FirstOrDefault();
                Assert.NotNull(vmlPart);
                var imagePart = Assert.Single(vmlPart!.ImageParts);
                Assert.Equal("image/jpeg", imagePart.ContentType);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public void Excel_HeaderFooter_TextUpdate_RemovesStalePictureArtifacts() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderFooterPictureCleanup.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var sheet = doc.AddWorksheet("Sheet1");
                var pngPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
                var pngBytes = File.ReadAllBytes(pngPath);
                sheet.SetHeaderImage(HeaderFooterPosition.Center, pngBytes, "image/png");
                sheet.SetHeaderFooter(headerCenter: "Plain header");
                doc.Save();
            }

            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                var sheetPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Empty(sheetPart.VmlDrawingParts);
                Assert.Null(sheetPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.LegacyDrawingHeaderFooter>().FirstOrDefault());
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                var sheet = document.Sheets.First();
                var hf = sheet.GetHeaderFooter();
                Assert.False(hf.HeaderHasPicturePlaceholder);
                Assert.Equal("Plain header", hf.HeaderCenter);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public async Task Excel_HeaderImageFromUrlAsync_Roundtrips_ContentType() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderImageUrlContentType.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var pngPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            var pngBytes = File.ReadAllBytes(pngPath);

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var url = $"http://127.0.0.1:{port}/logo.png";
            var acceptTask = ServeSingleImageAsync(listener, pngBytes, "image/png");

            try {
                using (var doc = ExcelDocument.Create(filePath))
                {
                    var sheet = doc.AddWorksheet("Sheet1");
                    await sheet.SetHeaderImageFromUrlAsync(HeaderFooterPosition.Center, url);
                    doc.Save();
                }
            } finally {
                listener.Stop();
                await acceptTask;
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
        public async Task Excel_FooterImageFromUrlAsync_Roundtrips_ContentType() {
            string filePath = Path.Combine(_directoryWithFiles, "FooterImageUrlContentType.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var jpegPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            var jpegBytes = File.ReadAllBytes(jpegPath);

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var url = $"http://127.0.0.1:{port}/logo.jpg";
            var acceptTask = ServeSingleImageAsync(listener, jpegBytes, "image/jpeg");

            try {
                using (var doc = ExcelDocument.Create(filePath))
                {
                    var sheet = doc.AddWorksheet("Sheet1");
                    await sheet.SetFooterImageFromUrlAsync(HeaderFooterPosition.Center, url);
                    doc.Save();
                }
            } finally {
                listener.Stop();
                await acceptTask;
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

        private static Task ServeSingleImageAsync(TcpListener listener, byte[] payload, string contentType) {
            return Task.Run(async () =>
            {
                try
                {
                    using var client = await listener.AcceptTcpClientAsync();
                    using var stream = client.GetStream();
                    using (var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true))
                    {
                        string? line;
                        while (!string.IsNullOrEmpty(line = await reader.ReadLineAsync())) { }
                    }

                    var header = $"HTTP/1.1 200 OK\r\nContent-Type: {contentType}\r\nContent-Length: {payload.Length}\r\nConnection: close\r\n\r\n";
                    var headerBytes = Encoding.ASCII.GetBytes(header);
                    await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
                    await stream.WriteAsync(payload, 0, payload.Length);
                    await stream.FlushAsync();
                }
                catch (SocketException)
                {
                    // Listener stopped before accepting a connection; ignore for test cleanup.
                }
                catch (ObjectDisposedException)
                {
                    // Listener disposed before accept completed; ignore for cleanup.
                }
            });
        }

    }
}
