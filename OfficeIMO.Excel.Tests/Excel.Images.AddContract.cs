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
    [Theory]
    [InlineData(ExcelImagePlacement.MoveAndSize, "twoCell")]
    [InlineData(ExcelImagePlacement.MoveOnly, "oneCell")]
    [InlineData(ExcelImagePlacement.FreeFloating, "absolute")]
    public void AddImageContract_SavesDistinctNativeAnchorsAndExactMetadata(ExcelImagePlacement placement, string editAs) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            var sheet = document.AddSheet("Images");
            var cell = sheet.AddImage("C2", bytes, new ExcelImageOptions {
                Width = OfficeLength.Emus(9526), Height = OfficeLength.Emus(19051),
                OffsetX = OfficeLength.Emus(-1), OffsetY = OfficeLength.Emus(7),
                Name = "Cell", Description = "Native bytes", Title = "Cell title",
                RotationDegrees = 12.5, LockAspectRatio = false
            });
            Assert.Equal(bytes, cell.ToBytes());
            cell.Title = "Attached title";
            sheet.AddImage("A1:B3", bytes, new ExcelImageOptions {
                OffsetX = OfficeLength.Emus(11), OffsetY = OfficeLength.Emus(-13),
                EndOffsetX = OfficeLength.Emus(-17), EndOffsetY = OfficeLength.Emus(19),
                Placement = placement, Name = "Range"
            });
            sheet.AddImage(OfficeLength.Emus(23), OfficeLength.Emus(29), bytes,
                new ExcelImageOptions { Width = OfficeLength.Emus(31), Height = OfficeLength.Emus(37), Name = "Absolute" });
            document.Save(output);
        }
        output.Position = 0;
        using var saved = SpreadsheetDocument.Open(output, false);
        var drawings = saved.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
        var cellAnchor = Assert.Single(drawings.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
        Assert.Equal("2", cellAnchor.FromMarker!.ColumnId!.Text);
        Assert.Equal("1", cellAnchor.FromMarker.RowId!.Text);
        Assert.Equal("-1", cellAnchor.FromMarker.ColumnOffset!.Text);
        Assert.Equal("7", cellAnchor.FromMarker.RowOffset!.Text);
        Assert.Equal(9526L, cellAnchor.Extent!.Cx!.Value);
        Assert.Equal(19051L, cellAnchor.Extent.Cy!.Value);
        var picture = cellAnchor.GetFirstChild<Xdr.Picture>()!;
        Assert.Equal("Attached title", picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Title!.Value);
        Assert.Equal("Native bytes", picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description!.Value);
        Assert.False(picture.NonVisualPictureProperties.NonVisualPictureDrawingProperties!.GetFirstChild<DocumentFormat.OpenXml.Drawing.PictureLocks>()!.NoChangeAspect!.Value);
        Assert.Equal(750000, picture.ShapeProperties!.Transform2D!.Rotation!.Value);
        var range = Assert.Single(drawings.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>());
        Assert.Equal(editAs, range.GetAttribute("editAs", string.Empty).Value);
        Assert.Equal("11", range.FromMarker!.ColumnOffset!.Text);
        Assert.Equal("-13", range.FromMarker.RowOffset!.Text);
        Assert.Equal("2", range.ToMarker!.ColumnId!.Text);
        Assert.Equal("3", range.ToMarker.RowId!.Text);
        Assert.Equal("-17", range.ToMarker.ColumnOffset!.Text);
        Assert.Equal("19", range.ToMarker.RowOffset!.Text);
        var absolute = Assert.Single(drawings.WorksheetDrawing.Elements<Xdr.AbsoluteAnchor>());
        Assert.Equal(23L, absolute.Position!.X!.Value);
        Assert.Equal(29L, absolute.Position.Y!.Value);
        Assert.Equal(31L, absolute.Extent!.Cx!.Value);
        Assert.Equal(37L, absolute.Extent.Cy!.Value);
        Assert.Equal(3, drawings.ImageParts.Count());
        foreach (ImagePart part in drawings.ImageParts) {
            using Stream source = part.GetStream();
            using var buffer = new MemoryStream();
            source.CopyTo(buffer);
            Assert.Equal(bytes, buffer.ToArray());
        }
    }

    [Fact]
    public async Task AddImageContract_PreservesByteDefaultsAndLoadedSourceSizing() {
        string path = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        byte[] bytes = File.ReadAllBytes(path);
        OfficeImageInfo info = OfficeImageReader.Identify(bytes);
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            var sheet = document.AddSheet("Sizing");
            sheet.AddImage(1, 1, bytes, new ExcelImageOptions { Name = "Bytes" });
            sheet.AddImage(2, 1, path, new ExcelImageOptions { Name = "File" });
            await sheet.AddImageAsync(3, 1, path, new ExcelImageOptions { ScalePercent = 20, Name = "Scaled" });
            var report = sheet.Report().AddImage(4, 1, bytes,
                new ExcelImageOptions { Width = OfficeLength.Emus(43), Height = OfficeLength.Emus(47), Name = "Report" });
            report.Title = "Attached report handle";
            document.Save(output);
        }
        output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        var images = reopened.Sheets.Single().Images.ToDictionary(image => image.Name);
        Assert.Equal(OfficeLength.Pixels(96), images["Bytes"].Width);
        Assert.Equal(OfficeLength.Pixels(32), images["Bytes"].Height);
        Assert.Equal(OfficeLength.Pixels(info.Width), images["File"].Width);
        Assert.Equal(OfficeLength.Pixels(info.Height), images["File"].Height);
        Assert.Equal(OfficeLength.Emus(Math.Max(1L, (OfficeLength.Pixels(info.Width) * .2).ToEmus())), images["Scaled"].Width);
        Assert.Equal(OfficeLength.Emus(Math.Max(1L, (OfficeLength.Pixels(info.Height) * .2).ToEmus())), images["Scaled"].Height);
        Assert.Equal(OfficeLength.Emus(43), images["Report"].Width);
        Assert.Equal("Attached report handle", images["Report"].Title);
        foreach (ExcelImage image in images.Values) Assert.Equal(bytes, image.ToBytes());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task AddImageContract_StreamUsesAsyncReadsAndHonorsOwnershipOnSuccessAndCancellation(bool leaveOpen) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Streams");
        using var stream = new AsyncImageContractStream(bytes) { Position = 5 };
        var image = await sheet.AddImageAsync(1, 1, stream, new ExcelImageOptions { LeaveOpen = leaveOpen });
        Assert.True(stream.AsyncReads > 0);
        Assert.Equal(leaveOpen, !stream.Disposed);
        Assert.Equal(5L, stream.PositionAtDispose ?? stream.Position);
        Assert.Equal(bytes, image.ToBytes());
        using var cancelled = new AsyncImageContractStream(bytes) { Position = 3 };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => sheet.AddImageAsync(2, 1, cancelled,
            new ExcelImageOptions { LeaveOpen = leaveOpen }, cancellation.Token));
        Assert.Equal(0, cancelled.AsyncReads);
        Assert.Equal(leaveOpen, !cancelled.Disposed);
        Assert.Equal(3L, cancelled.PositionAtDispose ?? cancelled.Position);
        Assert.Single(sheet.Images);
        using var invalid = new AsyncImageContractStream(new byte[] { 1, 2, 3 });
        await Assert.ThrowsAsync<ArgumentException>(() => sheet.AddImageAsync(2, 1, invalid,
            new ExcelImageOptions { LeaveOpen = leaveOpen }));
        Assert.Equal(leaveOpen, !invalid.Disposed);
        Assert.Single(sheet.Images);
    }

    [Fact]
    public void AddImageContract_OptionalDedupRetainsIndependentPicturesAndNativePayload() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            var sheet = document.AddSheet("Dedup");
            sheet.AddImage(1, 1, bytes, new ExcelImageOptions { Deduplicate = true, Name = "First", Title = "First title" });
            sheet.AddImage(3, 1, bytes, new ExcelImageOptions { Deduplicate = true, Name = "Second", Title = "Second title" });
            document.Save(output);
        }
        output.Position = 0;
        using var saved = SpreadsheetDocument.Open(output, false);
        var drawing = saved.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
        Assert.Single(drawing.ImageParts);
        var pictures = drawing.WorksheetDrawing!.Descendants<Xdr.Picture>().ToArray();
        Assert.Equal(2, pictures.Length);
        Assert.Equal(pictures[0].BlipFill!.Blip!.Embed!.Value, pictures[1].BlipFill!.Blip!.Embed!.Value);
        Assert.Equal("First title", pictures[0].NonVisualPictureProperties!.NonVisualDrawingProperties!.Title!.Value);
        Assert.Equal("Second title", pictures[1].NonVisualPictureProperties!.NonVisualDrawingProperties!.Title!.Value);
        using Stream payload = drawing.ImageParts.Single().GetStream();
        using var buffer = new MemoryStream();
        payload.CopyTo(buffer);
        Assert.Equal(bytes, buffer.ToArray());

        using var editableSource = new MemoryStream();
        output.WriteTo(editableSource);
        editableSource.Position = 0;
        using (var nativeEdit = SpreadsheetDocument.Open(editableSource, true)) {
            var nativeDrawing = nativeEdit.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            nativeDrawing.Elements<Xdr.OneCellAnchor>().First().Remove();
            nativeDrawing.Save();
        }
        editableSource.Position = 0;
        using var editable = ExcelDocument.Load(editableSource);
        var sheetAfterLoad = editable.Sheets.Single();
        sheetAfterLoad.DeleteRows(1);
        var survivingImage = Assert.Single(sheetAfterLoad.Images);
        Assert.Equal("Second", survivingImage.Name);
        Assert.Equal(bytes, survivingImage.ToBytes());
        using var afterRemoval = new MemoryStream();
        editable.Save(afterRemoval);
        afterRemoval.Position = 0;
        using var reloaded = ExcelDocument.Load(afterRemoval);
        Assert.Equal(bytes, Assert.Single(reloaded.Sheets.Single().Images).ToBytes());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task AddImageContract_RemotePolicyAndAnchorSurviveWorkbookSave(int anchorKind) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var uri = new Uri($"http://127.0.0.1:{((IPEndPoint)listener.LocalEndpoint).Port}/logo.png");
        Task response = ServeSingleImageAsync(listener, bytes, "image/png");
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Remote");
            var options = new ExcelImageOptions { Name = "Remote", Title = "Remote title",
                RemoteImageOptions = new OfficeRemoteImageLoadOptions { AllowPrivateNetworkAddresses = true } };
            ExcelImage image = anchorKind == 0 ? await sheet.AddImageAsync(2, 3, uri, options, timeout.Token)
                : anchorKind == 1 ? await sheet.AddImageAsync("A1:B3", uri, options, timeout.Token)
                : await sheet.AddImageAsync(OfficeLength.Emus(53), OfficeLength.Emus(59), uri, options, timeout.Token);
            Assert.Equal(bytes, image.ToBytes());
            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using var saved = SpreadsheetDocument.Open(output, false);
            var anchor = saved.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.ChildElements.Single();
            Assert.Equal(anchorKind == 0 ? "oneCellAnchor" : anchorKind == 1 ? "twoCellAnchor" : "absoluteAnchor", anchor.LocalName);
            Assert.Equal("Remote title", anchor.Descendants<Xdr.NonVisualDrawingProperties>().Single().Title!.Value);
        } finally {
            listener.Stop();
            await response;
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AddImageContract_SynchronousStreamsKeepPositionAndHonorOwnership(bool leaveOpen) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Sync stream");
        using var stream = new MemoryStream(bytes, writable: false) { Position = 5 };
        var image = sheet.AddImage(1, 1, stream, new ExcelImageOptions { LeaveOpen = leaveOpen });
        Assert.Equal(leaveOpen, stream.CanRead);
        if (leaveOpen) Assert.Equal(5L, stream.Position);
        Assert.Equal(bytes, image.ToBytes());
        using var invalid = new MemoryStream(new byte[] { 1, 2, 3 }, writable: false);
        Assert.Throws<ArgumentException>(() => sheet.AddImage(2, 1, invalid, new ExcelImageOptions { LeaveOpen = leaveOpen }));
        Assert.Equal(leaveOpen, invalid.CanRead);
        Assert.Single(sheet.Images);
    }

    [Fact]
    public void AddImageContract_ConflictingRangeSizingAndInvalidAbsolutePositionDoNotAddPictures() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Rejected geometry");
        Assert.Throws<ArgumentException>(() => sheet.AddImage("A1:B2", bytes,
            new ExcelImageOptions { Width = OfficeLength.Emus(1) }));
        Assert.Throws<ArgumentException>(() => sheet.AddImage("A1:B2", bytes,
            new ExcelImageOptions { ScalePercent = 20 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddImage(OfficeLength.Emus(-1), OfficeLength.Zero, bytes));
        Assert.Empty(sheet.Images);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var saved = SpreadsheetDocument.Open(output, false);
        Assert.Null(saved.WorkbookPart!.WorksheetParts.Single().DrawingsPart);
    }

    private sealed class AsyncImageContractStream : Stream {
        private readonly MemoryStream _inner;
        internal AsyncImageContractStream(byte[] bytes) => _inner = new MemoryStream(bytes, writable: false);
        internal int AsyncReads { get; private set; }
        internal bool Disposed { get; private set; }
        internal long? PositionAtDispose { get; private set; }
        public override bool CanRead => !Disposed;
        public override bool CanSeek => !Disposed;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Expected asynchronous source reads.");
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            AsyncReads++;
            return _inner.ReadAsync(buffer, offset, count, cancellationToken);
        }
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing && !Disposed) { PositionAtDispose = _inner.Position; Disposed = true; _inner.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
