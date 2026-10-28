using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderLoadedNativeProjectionTests {
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData("word", "officeimo.reader.word")]
    [InlineData("excel", "officeimo.reader.excel")]
    [InlineData("powerpoint", "officeimo.reader.powerpoint")]
    public void LoadedProjectionUsesCurrentModelAndPreservesPayloadOwner(string format, string capability) {
        using Fixture fixture = Create(format);
        OfficeDocumentReadResult full = fixture.Read(null, default);
        Assert.Contains(full.Chunks, chunk => chunk.Text!.Contains("Current model evidence", StringComparison.Ordinal));
        Assert.NotEmpty(full.Assets);
        Assert.False(string.IsNullOrEmpty(full.Markdown));
        Assert.Contains(capability, full.CapabilitiesUsed);

        // These load-admission settings do not retroactively constrain a caller-owned model.
        var options = new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, MaxInputBytes = 0,
            OpenXmlMaxPartUncompressedBytes = 1, MaxOpenXmlImageAssetBytes = 1
        };
        OfficeDocumentReadResult lean = fixture.Read(options, default);
        Assert.Contains(lean.Chunks, chunk => chunk.Text!.Contains("Current model evidence", StringComparison.Ordinal));
        Assert.All(lean.Chunks, chunk => Assert.Null(chunk.Markdown));
        Assert.Null(lean.Markdown);
        Assert.Empty(lean.Assets);
        Assert.Empty(lean.Visuals);
        Assert.Contains("officeimo.reader.projection.text-and-metadata", lean.CapabilitiesUsed);
        Assert.Equal("archive!/current-model", lean.Source.Path);
        Assert.Null(lean.Source.SourceHash);
        Assert.Null(lean.Source.LengthBytes);
        Assert.All(lean.Chunks, chunk => Assert.Equal("archive!/current-model", chunk.Location.Path));

        using OfficeEmbeddedPayloadInspection payloads = fixture.Inspect();
        Assert.Single(payloads.Payloads);
        Stream detached = payloads.OpenRead(0);
        fixture.Dispose();
        using (detached) {
            Assert.False(detached.CanWrite);
            using var bytes = new MemoryStream(); detached.CopyTo(bytes);
            Assert.Equal(new byte[] { 1, 2, 3 }, bytes.ToArray());
        }
        Assert.Contains(lean.Chunks, chunk => chunk.Text!.Contains("Current model evidence", StringComparison.Ordinal));
        Assert.Equal(ReaderProjection.TextAndMetadata, options.Projection);
        Assert.Equal(0, options.MaxInputBytes);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    public void LoadedProjectionEnforcesOutputBudgetAndRestoresScopeAfterFailure(string format) {
        using Fixture fixture = Create(format);
        var options = new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata,
            ResourceLimits = new ReaderResourceLimits { MaxChunkCharacters = 1 }
        };
        ReaderResourceLimitException limit = Assert.Throws<ReaderResourceLimitException>(() => fixture.Read(options, default));
        Assert.Equal(nameof(ReaderResourceLimits.MaxChunkCharacters), limit.LimitName);
        Assert.Equal(1, limit.Maximum);
        // A failed lean scope must not leak into a later default Full call.
        Assert.NotEmpty(fixture.Read(null, default).Assets);
        Assert.Equal(1, options.ResourceLimits.MaxChunkCharacters);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    public void LoadedProjectionCancellationAndInvalidOptionsLeaveOwnerUsable(string format) {
        using Fixture fixture = Create(format);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => fixture.Read(null, canceled.Token));
        Assert.Throws<ArgumentOutOfRangeException>(() => fixture.Read(new ReaderOptions { Projection = (ReaderProjection)99 }, default));
        Assert.NotEmpty(fixture.Read(null, default).Chunks);
        using OfficeEmbeddedPayloadInspection payloads = fixture.Inspect();
        Assert.Single(payloads.Payloads);
    }

    [Fact]
    public void LoadedExcelProjectionSnapshotsOptionsBeforeCallerConverterRuns() {
        using ExcelDocument document = ExcelDocument.Create();
        var sheet = document.AddSheet("Evidence");
        sheet.CellValue(1, 1, "Header"); sheet.CellFormula(2, 1, "1+1"); sheet.CellValue(3, 1, "Stable value");
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, ResourceLimits = new ReaderResourceLimits { MaxChunkCharacters = 1000 } };
        var format = new ReaderExcelOptions { SheetName = "Evidence", A1Range = "A1:A3", ChunkRows = 0, ReadOptions = new ExcelReaderOptions() };
        int calls = 0;
        format.ReadOptions.CellValueConverter = _ => {
            calls++;
            options.Projection = ReaderProjection.Full;
            options.ResourceLimits.MaxChunkCharacters = 0;
            format.A1Range = "Z1:Z2";
            format.ReadOptions.CellValueConverter = _ => throw new InvalidOperationException("Caller replacement must not replace the captured converter.");
            return ExcelCellValue.NotHandled;
        };
        OfficeDocumentReadResult result = document.ReadDocument("logical-sheet", options, format);
        Assert.True(calls > 1);
        Assert.Null(result.Markdown);
        Assert.All(result.Chunks, chunk => Assert.Equal("A1:A3", chunk.Location.A1Range));
        Assert.Equal("1", Assert.Single(result.Metadata, entry => entry.Id == "excel-formula-count").Value);
        Assert.Equal(0, format.ChunkRows);
        Assert.Equal("Z1:Z2", format.A1Range);
    }

    private static Fixture Create(string format) {
        if (format == "word") {
            WordDocument document = WordDocument.Create();
            using var picture = new MemoryStream(Pixel, writable: false);
            document.AddParagraph("Current model evidence").AddImage(picture, "pixel.png", OfficeIMO.OfficeLength.Pixels(1), OfficeIMO.OfficeLength.Pixels(1));
            AddPayload(document.OpenXmlDocument.MainDocumentPart!);
            return new Fixture(document, (options, token) => document.ReadDocument("archive!/current-model", options, cancellationToken: token), () => document.InspectEmbeddedPayloads());
        }
        if (format == "excel") {
            ExcelDocument document = ExcelDocument.Create();
            var sheet = document.AddSheet("Evidence");
            sheet.CellValue(1, 1, "Header"); sheet.CellValue(2, 1, "Current model evidence");
            sheet.AddImage(1, 1, Pixel, width: OfficeIMO.OfficeLength.Pixels(1), height: OfficeIMO.OfficeLength.Pixels(1));
            AddPayload(document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single());
            return new Fixture(document, (options, token) => document.ReadDocument("archive!/current-model", options, cancellationToken: token), () => document.InspectEmbeddedPayloads());
        }
        PowerPointDocument presentation = PowerPointDocument.Create();
        var slide = presentation.AddSlide(); slide.AddTextBox("Current model evidence",global::OfficeIMO.OfficeLength.Emus(100000),global::OfficeIMO.OfficeLength.Emus(100000),global::OfficeIMO.OfficeLength.Emus(2000000),global::OfficeIMO.OfficeLength.Emus(500000));
        using (var picture = new MemoryStream(Pixel, writable: false)) slide.AddImage(picture, OfficeImageFormat.Png);
        AddPayload(presentation.OpenXmlDocument.PresentationPart!.SlideParts.Single());
        return new Fixture(presentation, (options, token) => presentation.ReadDocument("archive!/current-model", options, cancellationToken: token), () => presentation.InspectEmbeddedPayloads());
    }

    private static void AddPayload(OpenXmlPart owner) {
        EmbeddedPackagePart part = owner.AddNewPart<EmbeddedPackagePart>("application/octet-stream", "rPayload");
        using var bytes = new MemoryStream(new byte[] { 1, 2, 3 }, writable: false); part.FeedData(bytes);
    }

    private sealed class Fixture : IDisposable {
        private readonly IDisposable _owner;
        private bool _disposed;
        internal Fixture(IDisposable owner, Func<ReaderOptions?, CancellationToken, OfficeDocumentReadResult> read,
            Func<OfficeEmbeddedPayloadInspection> inspect) { _owner = owner; Read = read; Inspect = inspect; }
        internal Func<ReaderOptions?, CancellationToken, OfficeDocumentReadResult> Read { get; }
        internal Func<OfficeEmbeddedPayloadInspection> Inspect { get; }
        public void Dispose() { if (_disposed) return; _disposed = true; _owner.Dispose(); }
    }
}
