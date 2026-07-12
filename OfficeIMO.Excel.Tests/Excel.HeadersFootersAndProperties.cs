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

                var sheet = doc.AddWorkSheet("Summary");
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
                var sheet = doc.AddWorkSheet("Sheet1");
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
                var sheet = doc.AddWorkSheet("Sheet1");
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
                var sheet = doc.AddWorkSheet("Sheet1");
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
                var sheet = doc.AddWorkSheet("Sheet1");
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
        public async Task ImageDownloader_Normalizes_Image_Content_Type_Alias()
        {
            OfficeIMO.Excel.ImageDownloader.ClearCache();

            var jpegPath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            var jpegBytes = File.ReadAllBytes(jpegPath);

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var url = $"http://127.0.0.1:{port}/logo.jpg";
            var acceptTask = ServeSingleImageAsync(listener, jpegBytes, "image/jpg; charset=binary");

            try
            {
                Assert.True(OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 2_000_000, out var bytes, out var contentType));
                Assert.NotNull(bytes);
                Assert.Equal(jpegBytes, bytes);
                Assert.Equal(OfficeImageInfo.GetMimeType(OfficeImageFormat.Jpeg), contentType);
            }
            finally
            {
                listener.Stop();
                await acceptTask;
                OfficeIMO.Excel.ImageDownloader.ClearCache();
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public async Task ImageDownloader_Rejects_Redirect_To_NonHttp_Target() {
            OfficeIMO.Excel.ImageDownloader.ClearCache();

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var url = $"http://127.0.0.1:{port}/redirect.png";
            var response = "HTTP/1.1 302 Found\r\nLocation: file:///C:/Windows/win.ini\r\nConnection: close\r\n\r\n";
            var acceptTask = ServeSingleRawResponseAsync(listener, Encoding.ASCII.GetBytes(response));

            try {
                Assert.False(OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 2_000_000, out var bytes, out var contentType));
                Assert.Null(bytes);
                Assert.Null(contentType);
            } finally {
                listener.Stop();
                await acceptTask;
                OfficeIMO.Excel.ImageDownloader.ClearCache();
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public async Task ImageDownloader_Rejects_CrossOrigin_Redirect_Before_Fetching_Target() {
            OfficeIMO.Excel.ImageDownloader.ClearCache();

            var redirectListener = new TcpListener(IPAddress.Loopback, 0);
            var targetListener = new TcpListener(IPAddress.Loopback, 0);
            redirectListener.Start();
            targetListener.Start();
            var redirectPort = ((IPEndPoint)redirectListener.LocalEndpoint).Port;
            var targetPort = ((IPEndPoint)targetListener.LocalEndpoint).Port;
            var url = $"http://127.0.0.1:{redirectPort}/redirect.png";
            var targetUrl = $"http://127.0.0.1:{targetPort}/private.png";
            var response = $"HTTP/1.1 302 Found\r\nLocation: {targetUrl}\r\nConnection: close\r\n\r\n";
            var redirectTask = ServeSingleRawResponseAsync(redirectListener, Encoding.ASCII.GetBytes(response));
            var targetRequestCount = 0;
            var targetTask = Task.Run(async () =>
            {
                try
                {
                    using var client = await targetListener.AcceptTcpClientAsync();
                    targetRequestCount++;
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

            try {
                Assert.False(OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 2_000_000, out var bytes, out var contentType));
                Assert.Null(bytes);
                Assert.Null(contentType);
                await Task.Delay(100);
                Assert.Equal(0, targetRequestCount);
            } finally {
                redirectListener.Stop();
                targetListener.Stop();
                await redirectTask;
                await targetTask;
                OfficeIMO.Excel.ImageDownloader.ClearCache();
            }
        }

        [Fact]
        [Trait("Category","ExcelHeaderFooterImages")]
        public async Task ImageDownloader_Rejects_Response_When_Stream_Exceeds_Limit() {
            OfficeIMO.Excel.ImageDownloader.ClearCache();

            var payload = Enumerable.Repeat((byte)0x41, 64).ToArray();
            var header = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nConnection: close\r\n\r\n");
            var response = header.Concat(payload).ToArray();

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var url = $"http://127.0.0.1:{port}/large.png";
            var acceptTask = ServeSingleRawResponseAsync(listener, response);

            try {
                Assert.False(OfficeIMO.Excel.ImageDownloader.TryFetch(url, 5, 16, out var bytes, out var contentType));
                Assert.Null(bytes);
                Assert.Null(contentType);
            } finally {
                listener.Stop();
                await acceptTask;
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
                    doc.Save();
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
                    doc.Save();
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

        private static Task ServeSingleRawResponseAsync(TcpListener listener, byte[] responseBytes) {
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

                    await stream.WriteAsync(responseBytes, 0, responseBytes.Length);
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
