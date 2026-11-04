using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void ImageLengths_MovingATwoCellImageAcrossDifferentColumnWidthsKeepsItsSize() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Images");
        sheet.SetColumnWidth(5, ExcelColumnWidth.Characters(40));
        sheet.SetColumnWidth(6, ExcelColumnWidth.Characters(40));
        ExcelImage image = sheet.AddImage("A1:C10", bytes, new OfficeIMO.Excel.ExcelImageOptions {
            OffsetX = OfficeLength.Pixels(30), EndOffsetX = OfficeLength.Pixels(10)
        });
        OfficeLength width = image.Width, height = image.Height;

        image.MoveTo(1, 5);

        Assert.True(image.ToOffsetX >= OfficeLength.Zero);
        Assert.Equal(width.ToEmus(), image.Width.ToEmus(), tolerance: 1D);
        Assert.Equal(height.ToEmus(), image.Height.ToEmus(), tolerance: 1D);
    }

    [Fact]
    public void ImageCleanup_KeepsThePartBehindAnSvgBlipExtension() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        var stream = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            document.AddSheet("Images").AddImage(1, 1, bytes);
            document.Save(stream);
        }
        stream.Position = 0;
        using (var package = SpreadsheetDocument.Open(stream, true)) {
            DrawingsPart drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
            ImagePart vector = drawing.AddImagePart(ImagePartType.Svg);
            using (var svg = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'/>"))) vector.FeedData(svg);
            drawing.WorksheetDrawing!.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single().Append(
                new DocumentFormat.OpenXml.Drawing.BlipExtensionList(new DocumentFormat.OpenXml.Drawing.BlipExtension(
                    new DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip { Embed = drawing.GetIdOfPart(vector) }) {
                    Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
                }));
        }

        stream.Position = 0;
        using var loaded = ExcelDocument.Load(stream);
        loaded.Sheets[0].Cell(3, 3).SetValue("edited");
        using var saved = SpreadsheetDocument.Open(new MemoryStream(loaded.ToBytes()), false);
        DrawingsPart savedDrawing = saved.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
        string vectorId = savedDrawing.WorksheetDrawing!.Descendants<DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip>().Single().Embed!.Value!;
        Assert.Equal("image/svg+xml", savedDrawing.GetPartById(vectorId).ContentType);
    }

    [Fact]
    public void ImageLengths_RetainNativeEmusAndAttachedReportHandleAfterSave() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            var sheet = document.AddSheet("Images");
            ExcelImage image = sheet.Report().AddImage(2, 3, bytes, new OfficeIMO.Excel.ExcelImageOptions {
                Width = OfficeLength.Emus(9526),
                Height = OfficeLength.Emus(19051),
                OffsetX = OfficeLength.Emus(-1),
                OffsetY = OfficeLength.Emus(17)
            });
            Assert.Equal(OfficeLength.Emus(9526), image.Width);
            Assert.Equal(OfficeLength.Emus(-1), image.OffsetX);
            image.Name = "Measured image";
            image.MoveTo(4, 5, OfficeLength.Emus(23), OfficeLength.Emus(-29));
            image.SetSize(OfficeLength.Emus(37), OfficeLength.Emus(41));
            document.Save(output);
        }
        output.Position = 0;
        using var saved = SpreadsheetDocument.Open(output, false);
        var anchor = Assert.Single(saved.WorkbookPart!.WorksheetParts.Single().DrawingsPart!
            .WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
        Assert.Equal("23", anchor.FromMarker!.ColumnOffset!.Text);
        Assert.Equal("-29", anchor.FromMarker.RowOffset!.Text);
        Assert.Equal(37L, anchor.Extent!.Cx!.Value);
        Assert.Equal(41L, anchor.Extent.Cy!.Value);
        Assert.Equal("Measured image", anchor.Descendants<Xdr.NonVisualDrawingProperties>().Single().Name!.Value);
    }

    [Fact]
    public void ImageLengths_EndMarkerRejectsNegativeSubPixelOffsetsBeforeMutation() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Images");
        var image = sheet.AddImage("A1:B2", bytes);
        image.SetTwoCellEndingMarker(3, 4, OfficeLength.Emus(1), OfficeLength.Emus(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => image.SetTwoCellEndingMarker(5, 6, OfficeLength.Emus(-1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => image.SetTwoCellEndingMarker(5, 6, toOffsetY: OfficeLength.Emus(-1)));
        Assert.Equal(3, image.ToRowIndex);
        Assert.Equal(4, image.ToColumnIndex);
        Assert.Equal(OfficeLength.Emus(1), image.ToOffsetX);
        Assert.Equal(OfficeLength.Emus(2), image.ToOffsetY);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImageLengths_RemoteDefaultsUseSourceDimensionsAndExplicitLengthsStayExact(bool reportImage) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        OfficeImageInfo info = OfficeImageReader.Identify(bytes);
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        Task response = ServeSingleImageAsync(listener, bytes, "image/png");
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Remote");
            var policy = new OfficeRemoteImageLoadOptions { AllowPrivateNetworkAddresses = true };
            string url = $"http://127.0.0.1:{port}/logo.png";
            ExcelImage image = reportImage
                ? await sheet.Report().AddImageAsync(1, 1, new System.Uri(url), new OfficeIMO.Excel.ExcelImageOptions {
                    RemoteImageOptions = policy,
                    Width = OfficeLength.Emus(9526),
                    Height = OfficeLength.Emus(19051)
                }, cancellationToken: timeout.Token)
                : await sheet.AddImageAsync(1, 1, new System.Uri(url), new OfficeIMO.Excel.ExcelImageOptions { RemoteImageOptions = policy }, cancellationToken: timeout.Token);
            Assert.Equal(reportImage ? OfficeLength.Emus(9526) : OfficeLength.Pixels(info.Width), image.Width);
            Assert.Equal(reportImage ? OfficeLength.Emus(19051) : OfficeLength.Pixels(info.Height), image.Height);
            image.Name = "Remote handle";
            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using var saved = SpreadsheetDocument.Open(output, false);
            Assert.Equal("Remote handle", saved.WorkbookPart!.WorksheetParts.Single().DrawingsPart!
                .WorksheetDrawing!.Descendants<Xdr.NonVisualDrawingProperties>().Single().Name!.Value);
        } finally {
            listener.Stop();
            await response;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageLengths_RangeCopiesRetainPhysicalSizeWithSignedNativeEndOffsets(bool transpose) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var source = new MemoryStream();
        using (var original = ExcelDocument.Create()) {
            var sheet = original.AddSheet("Copy images");
            sheet.Cell(1, 1).SetValue("Seed");
            sheet.AddImage("A1:B3", bytes, new OfficeIMO.Excel.ExcelImageOptions { EndOffsetX = OfficeLength.Emus(-1), EndOffsetY = OfficeLength.Emus(-2) });
            original.Save(source);
        }
        source.Position = 0;
        using var document = ExcelDocument.Load(source);
        var target = document.Sheets.Single();
        var image = Assert.Single(target.Images);
        var width = image.Width;
        var height = image.Height;
        if (transpose) target.Range("A1:B3").TransposeTo("E5");
        else target.Range("A1:B3").CopyTo("E5");
        var copy = Assert.Single(target.Images, value => value.RowIndex == 5 && value.ColumnIndex == 5);
        Assert.Equal(transpose ? height : width, copy.Width);
        Assert.Equal(transpose ? width : height, copy.Height);
        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var savedCopy = Assert.Single(reopened.Sheets.Single().Images,
            value => value.RowIndex == 5 && value.ColumnIndex == 5);
        Assert.Equal(copy.Width, savedCopy.Width);
        Assert.Equal(copy.Height, savedCopy.Height);
    }

}
