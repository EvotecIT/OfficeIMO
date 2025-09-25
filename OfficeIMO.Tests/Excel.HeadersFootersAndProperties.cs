using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
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

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public async Task ImageDownloader_Reuses_Cache_For_Repeat_Urls()
        {
            OfficeIMO.Excel.ImageDownloader.ClearCache();

            var pngPath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            var pngBytes = File.ReadAllBytes(pngPath);

            var listener = new TcpListener(IPAddress.Loopback, 0);
            try
            {
                listener.Start();
                var port = ((IPEndPoint)listener.LocalEndpoint).Port;
                var url = $"http://127.0.0.1:{port}/logo.png";
                int requestCount = 0;

                var acceptTask = Task.Run(async () =>
                {
                    try
                    {
                        using var client = await listener.AcceptTcpClientAsync();
                        requestCount++;
                        using var stream = client.GetStream();
                        using (var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true))
                        {
                            string? line;
                            while (!string.IsNullOrEmpty(line = await reader.ReadLineAsync())) { }
                        }

                        var header = $"HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nContent-Length: {pngBytes.Length}\r\nConnection: close\r\n\r\n";
                        var headerBytes = Encoding.ASCII.GetBytes(header);
                        await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
                        await stream.WriteAsync(pngBytes, 0, pngBytes.Length);
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

                try
                {
                    Assert.True(OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 2_000_000, out var firstBytes, out var firstContentType));
                    Assert.NotNull(firstBytes);
                    Assert.Equal("image/png", firstContentType);
                    Assert.Equal(pngBytes, firstBytes);
                }
                catch
                {
                    listener.Stop();
                    await acceptTask;
                    throw;
                }

                listener.Stop();
                await acceptTask;

                // Second request should be served from cache even though the listener is stopped.
                Assert.True(OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 2_000_000, out var cachedBytes, out var cachedContentType));
                Assert.NotNull(cachedBytes);
                Assert.Equal("image/png", cachedContentType);
                Assert.Equal(pngBytes, cachedBytes);
                Assert.Equal(1, requestCount);
            }
            finally
            {
                listener.Stop();
                OfficeIMO.Excel.ImageDownloader.ClearCache();
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public async Task Excel_HeaderImageUrl_Roundtrips_ContentType() {
            OfficeIMO.Excel.ImageDownloader.ClearCache();

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
                    var sheet = doc.AddWorkSheet("Sheet1");
                    sheet.SetHeaderImageUrl(HeaderFooterPosition.Center, url);
                    doc.Save(false);
                }
            } finally {
                listener.Stop();
                await acceptTask;
                OfficeIMO.Excel.ImageDownloader.ClearCache();
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
        public async Task Excel_FooterImageUrl_Roundtrips_ContentType() {
            OfficeIMO.Excel.ImageDownloader.ClearCache();

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
                    var sheet = doc.AddWorkSheet("Sheet1");
                    sheet.SetFooterImageUrl(HeaderFooterPosition.Center, url);
                    doc.Save(false);
                }
            } finally {
                listener.Stop();
                await acceptTask;
                OfficeIMO.Excel.ImageDownloader.ClearCache();
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
