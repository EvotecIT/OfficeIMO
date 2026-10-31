using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EmbeddedPayloadInspectionContracts {
    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    public void DescriptorInventoryDoesNotScanUnrequestedPayloads(string format) {
        using Fixture fixture = Create(format);
        using OfficeEmbeddedPayloadInspection inspection = fixture.Inspect(new OfficeEmbeddedPayloadInspectionOptions { MaxReadBytes = 0 }, default);
        Assert.Equal(2, inspection.Payloads.Count);
        Assert.Equal(new[] { 0, 1 }, inspection.Payloads.Select(item => item.Ordinal));
        Assert.All(inspection.Payloads, item => {
            Assert.False(item.HasKnownLength);
            Assert.Equal(-1, item.Length);
            Assert.Null(item.Sha256);
            Assert.Equal(OfficeEmbeddedPayloadKind.EmbeddedPackage, item.Kind);
            Assert.NotEmpty(item.OwnerPartUri);
            Assert.NotEmpty(item.RelationshipId);
        });
        Assert.Equal(0, inspection.ReadBytes);
        OfficeInspectionResourceLimitException exception = Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(0));
        Assert.Equal("embedded-payload", exception.Owner);
        Assert.Equal("MaxReadBytes", exception.LimitName);
        Assert.Equal(0, exception.Maximum);
        long consumed = inspection.ReadBytes;
        Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(0));
        Assert.Equal(consumed, inspection.ReadBytes);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    public void SelectedReadsUseSnapshottedAggregateWorkAndDetachedLifetime(string format) {
        using Fixture fixture = Create(format);
        var options = new OfficeEmbeddedPayloadInspectionOptions { MaxPayloadBytes = 4, MaxTotalPayloadBytes = 7, MaxReadBytes = 4 };
        OfficeEmbeddedPayloadInspection inspection = fixture.Inspect(options, default);
        options.MaxReadBytes = 0;
        options.MaxPayloadBytes = 0;
        Stream detached = inspection.OpenRead(0);
        Assert.False(detached.CanWrite);
        Assert.Equal(4, inspection.ReadBytes);
        OfficeInspectionResourceLimitException exception = Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(1));
        Assert.Equal("MaxReadBytes", exception.LimitName);
        Assert.Equal(4, exception.Maximum);
        inspection.Dispose();
        Assert.Throws<ObjectDisposedException>(() => inspection.OpenRead(0));
        fixture.Dispose();
        using (detached) Assert.Equal(new byte[] { 1, 2, 3, 4 }, Read(detached));
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    public void RequestedHashesAndSelectedReadsShareOneFiniteOperationBudget(string format) {
        using Fixture fixture = Create(format);
        using OfficeEmbeddedPayloadInspection inspection = fixture.Inspect(new OfficeEmbeddedPayloadInspectionOptions {
            IncludeSha256 = true, MaxPayloadBytes = 4, MaxTotalPayloadBytes = 7, MaxReadBytes = 7
        }, default);
        Assert.Equal(7, inspection.ReadBytes);
        Assert.Equal(new long[] { 4, 3 }, inspection.Payloads.Select(item => item.Length));
        Assert.All(inspection.Payloads, item => Assert.Equal(64, item.Sha256!.Length));
        Assert.Equal("9f64a747e1b97f131fabb6b447296c9b6f0201e79fb3c5356e6c77e89b6a806a", inspection.Payloads[0].Sha256);
        OfficeInspectionResourceLimitException exception = Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(0));
        Assert.Equal("MaxReadBytes", exception.LimitName);
        Assert.Equal(7, exception.Maximum);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    public void NativeSinglePayloadReadEnforcesMemberLimitAndCancellation(string format) {
        using Fixture fixture = Create(format);
        using OfficeEmbeddedPayloadInspection inspection = fixture.Inspect(new OfficeEmbeddedPayloadInspectionOptions(), default);
        string id = inspection.Payloads[0].Id;
        OfficeInspectionResourceLimitException exception = Assert.Throws<OfficeInspectionResourceLimitException>(() => fixture.Open(id, 3, default));
        Assert.Equal("MaxPayloadBytes", exception.LimitName);
        Assert.Equal(3, exception.Maximum);
        using Stream source = fixture.Open(id, 4, default);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, Read(source));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => fixture.Open(id, 4, canceled.Token));
        Assert.Throws<OperationCanceledException>(() => fixture.Inspect(new OfficeEmbeddedPayloadInspectionOptions(), canceled.Token));
        OfficeInspectionResourceLimitException count = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            fixture.Inspect(new OfficeEmbeddedPayloadInspectionOptions { MaxPayloads = 1 }, default));
        Assert.Equal("MaxPayloads", count.LimitName);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    public void UnknownPayloadLengthsStillEnforceAggregateAdmissionBeforeReturningBytes(string format) {
        using Fixture fixture = Create(format);
        using OfficeEmbeddedPayloadInspection inspection = fixture.Inspect(new OfficeEmbeddedPayloadInspectionOptions {
            MaxPayloadBytes = 4, MaxTotalPayloadBytes = 6, MaxReadBytes = 10
        }, default);
        using Stream first = inspection.OpenRead(0);
        OfficeInspectionResourceLimitException aggregate = Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(1));
        Assert.Equal("MaxTotalPayloadBytes", aggregate.LimitName);
        Assert.Equal(6, aggregate.Maximum);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, Read(first));
    }

    private static Fixture Create(string format) {
        IDisposable document;
        OpenXmlPartContainer owner;
        Func<OfficeEmbeddedPayloadInspectionOptions, CancellationToken, OfficeEmbeddedPayloadInspection> inspect;
        Func<string, long, CancellationToken, Stream> open;
        if (format == "word") {
            WordDocument word = WordDocument.Create(); document = word;
            owner = word.OpenXmlDocument.MainDocumentPart!;
            inspect = word.InspectEmbeddedPayloads; open = word.OpenEmbeddedPayload;
        } else if (format == "excel") {
            ExcelDocument excel = ExcelDocument.Create(); excel.AddSheet("Payloads"); document = excel;
            owner = excel.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
            inspect = excel.InspectEmbeddedPayloads; open = excel.OpenEmbeddedPayload;
        } else {
            PowerPointDocument presentation = PowerPointDocument.Create(); presentation.AddSlide(); document = presentation;
            owner = presentation.OpenXmlDocument.PresentationPart!.SlideParts.Single();
            inspect = presentation.InspectEmbeddedPayloads; open = presentation.OpenEmbeddedPayload;
        }
        Add(owner, "rIdPayload1", new byte[] { 1, 2, 3, 4 });
        Add(owner, "rIdPayload2", new byte[] { 5, 6, 7 });
        return new Fixture(document, inspect, open);
    }

    private static void Add(OpenXmlPartContainer owner, string id, byte[] data) {
        EmbeddedPackagePart part = owner.AddNewPart<EmbeddedPackagePart>("application/octet-stream", id);
        using var source = new MemoryStream(data, writable: false); part.FeedData(source);
    }
    private static byte[] Read(Stream stream) { using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray(); }
    private sealed class Fixture : IDisposable {
        private readonly IDisposable _owner;
        private bool _disposed;
        internal Fixture(IDisposable owner, Func<OfficeEmbeddedPayloadInspectionOptions, CancellationToken, OfficeEmbeddedPayloadInspection> inspect,
            Func<string, long, CancellationToken, Stream> open) { _owner = owner; Inspect = inspect; Open = open; }
        internal Func<OfficeEmbeddedPayloadInspectionOptions, CancellationToken, OfficeEmbeddedPayloadInspection> Inspect { get; }
        internal Func<string, long, CancellationToken, Stream> Open { get; }
        public void Dispose() { if (!_disposed) { _disposed = true; _owner.Dispose(); } }
    }
}
