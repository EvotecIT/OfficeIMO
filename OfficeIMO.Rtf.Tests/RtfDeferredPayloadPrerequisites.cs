using System.Security.Cryptography;
using System.Threading;
using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;
using Xunit;

namespace OfficeIMO.Rtf.Tests;

public sealed class RtfDeferredPayloadPrerequisiteTests {
    [Theory]
    [InlineData(@"{\rtf1 text{\pict\pngblip 41\bin-1 42}}", "RTF003")]
    [InlineData(@"{\rtf1 text{\pict\pngblip 41\bin2147483648 42}}", "RTF002")]
    [InlineData(@"{\rtf1 text{\object\objemb{\*\objdata 41{\bin-1 42}}}}", "RTF003")]
    [InlineData(@"{\rtf1 text{\object\objemb{\*\objdata 41{\bin2147483648 42}}}}", "RTF002")]
    [InlineData(@"{\rtf1 text{\pict\pngblip 41\bin 42}}", null)]
    public void InvalidOriginalBinaryDeclarationsCannotPublishACompletePayloadOrHash(string source, string? diagnostic) {
        foreach (bool deferred in new[] { false, true }) {
            RtfReadResult result = RtfDocument.ParseResult(source, new RtfReadOptions { DeferBinaryPayloads = deferred, ReadEmbeddedObjects = true });
            if (diagnostic != null) Assert.Contains(result.Diagnostics, item => item.Code == diagnostic);
            byte[] semantic = result.Document.Paragraphs[0].Inlines.OfType<RtfImage>().FirstOrDefault()?.Data
                ?? result.Document.Paragraphs[0].Inlines.OfType<RtfObject>().Single().Data;
            Assert.Equal(new byte[] { 65, 66 }, semantic);
            using OfficeEmbeddedPayloadInspection inspection = result.InspectEmbeddedPayloads();
            Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0));
            Assert.Throws<InvalidDataException>(() => result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions { IncludeSha256 = true }));
        }
    }

    [Fact]
    public void InvalidDeclarationRefusesItsWholeMemberAndDoesNotReplaceValidSiblingSourceBytes() {
        RtfReadResult result = RtfDocument.ParseResult(@"{\rtf1 text{\pict\pngblip 414243}{\object\objemb{\*\objdata 41\bin-1 42}}}",
            new RtfReadOptions { DeferBinaryPayloads = true });
        using OfficeEmbeddedPayloadInspection inspection = result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions {
            MaxPayloads = 2, MaxPayloadBytes = 3, MaxTotalPayloadBytes = 6, MaxReadBytes = 7
        });
        using Stream valid = inspection.OpenRead(0); Assert.Equal(Encoding.ASCII.GetBytes("ABC"), Read(valid));
        Assert.Throws<InvalidDataException>(() => inspection.OpenRead(1)); Assert.Equal(4, inspection.ReadBytes);
        Assert.Throws<InvalidDataException>(() => result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions { IncludeSha256 = true }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativePicturesAndExplicitObjectsRetainNormalGetterAndSaveBytes(bool deferred) {
        const string source = @"{\rtf1 text{\pict\pngblip 414243}{\object\objemb{\*\objdata 444546}{\result fallback}}}";
        RtfReadResult result = RtfDocument.ParseResult(source, new RtfReadOptions { DeferBinaryPayloads = deferred, ReadEmbeddedObjects = true });
        RtfParagraph paragraph = result.Document.Paragraphs[0];
        RtfImage picture = Assert.Single(paragraph.Inlines.OfType<RtfImage>());
        RtfObject embedded = Assert.Single(paragraph.Inlines.OfType<RtfObject>());
        Assert.Equal(3, picture.PayloadLength); Assert.Equal(3, embedded.PayloadLength);
        Assert.Equal(new byte[] { 65, 66, 67 }, picture.Data);
        Assert.Equal(new byte[] { 68, 69, 70 }, embedded.Data);
        Assert.Equal(source, result.ToRtfLossless());
        picture.Data = new byte[] { 71, 72 };
        RtfDocument reopened = RtfDocument.Parse(result.Document.ToRtf(), new RtfReadOptions { ReadEmbeddedObjects = true });
        Assert.Equal(new byte[] { 71, 72 }, reopened.Paragraphs[0].Inlines.OfType<RtfImage>().Single().Data);
        Assert.Equal(new byte[] { 68, 69, 70 }, reopened.Paragraphs[0].Inlines.OfType<RtfObject>().Single().Data);
    }

    [Fact]
    public void OriginalSourceInventoryIncludesBlockedObjectsWithoutUnrequestedHexScansOrHashes() {
        RtfReadResult result = RtfDocument.ParseResult(@"{\rtf1 text{\pict\pngblip 414243}{\object\objemb{\*\objdata 444546}}}",
            new RtfReadOptions { DeferBinaryPayloads = true });
        Assert.Empty(result.Document.Paragraphs[0].Inlines.OfType<RtfObject>());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF105");
        using OfficeEmbeddedPayloadInspection inspection = result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions {
            MaxPayloads = 2, MaxPayloadBytes = 3, MaxTotalPayloadBytes = 6, MaxReadBytes = 6
        });
        Assert.Equal(2, inspection.Payloads.Count); Assert.Equal(0, inspection.ReadBytes);
        Assert.All(inspection.Payloads, info => {
            Assert.False(info.HasKnownLength); Assert.Equal(-1, info.Length); Assert.Null(info.Sha256);
            Assert.StartsWith("rtf-source:/group/", info.Id); Assert.Equal(string.Empty, info.RelationshipId);
        });
        result.Document.Paragraphs[0].Inlines.OfType<RtfImage>().Single().Data = new byte[] { 1, 2 };
        using Stream picture = inspection.OpenRead(0); using Stream embedded = inspection.OpenRead(1);
        inspection.Dispose();
        Assert.Equal(new byte[] { 65, 66, 67 }, Read(picture)); Assert.Equal(new byte[] { 68, 69, 70 }, Read(embedded));
    }

    [Fact]
    public void OriginalBinarySlicesAndLosslessBytesSurviveCallerMutationAndIndependentCancellation() {
        byte[] bytes = Encoding.ASCII.GetBytes(@"{\rtf1 text{\pict\pngblip\bin3 ABC}}");
        using var opening = new CancellationTokenSource();
        RtfReadResult result = RtfDocument.LoadResult(bytes, new RtfReadOptions { DeferBinaryPayloads = true }, opening.Token);
        Array.Clear(bytes, 0, bytes.Length); opening.Cancel();
        using OfficeEmbeddedPayloadInspection inspection = result.InspectEmbeddedPayloads();
        Assert.Equal(3, Assert.Single(inspection.Payloads).Length);
        Assert.Throws<OperationCanceledException>(() => inspection.OpenRead(0, opening.Token)); Assert.Equal(0, inspection.ReadBytes);
        using Stream selected = inspection.OpenRead(0);
        Assert.Equal(Encoding.ASCII.GetBytes("ABC"), Read(selected));
        Assert.Equal(@"{\rtf1 text{\pict\pngblip\bin3 ABC}}", Encoding.ASCII.GetString(result.ToBytesLossless()));
        RtfBinary binary = result.SyntaxTree.Root.Children.OfType<RtfGroup>().Single(group => group.Destination == "pict").Children.OfType<RtfBinary>().Single();
        Assert.Equal(3, binary.Length); Assert.Equal(Encoding.ASCII.GetBytes("ABC"), binary.Data);
    }

    [Theory]
    [InlineData("image")]
    [InlineData("object")]
    [InlineData("bin")]
    public void DeferredParsingStillRefusesNativeDeclaredAndDecodedPayloadBounds(string kind) {
        var options = new RtfReadOptions { DeferBinaryPayloads = true, ReadEmbeddedObjects = true,
            MaxImageBytesPerImage = 2, MaxObjectBytesPerObject = 2, MaxBinaryBytesPerPayload = 2 };
        string source = kind == "image" ? @"{\rtf1{\pict\pngblip 414243}}"
            : kind == "object" ? @"{\rtf1{\object\objemb{\*\objdata 414243}}}" : @"{\rtf1{\pict\pngblip\bin3 ABC}}";
        Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(source, options));
    }

    [Fact]
    public void RequestedCompleteHashesAndCumulativeReadsObserveSelectedLimits() {
        RtfReadResult result = RtfDocument.ParseResult(@"{\rtf1{\pict\pngblip 414243}}", new RtfReadOptions { DeferBinaryPayloads = true });
        using OfficeEmbeddedPayloadInspection hashes = result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions {
            IncludeSha256 = true, MaxPayloadBytes = 3, MaxTotalPayloadBytes = 3, MaxReadBytes = 3
        });
        using SHA256 sha = SHA256.Create();
        string expected = BitConverter.ToString(sha.ComputeHash(Encoding.ASCII.GetBytes("ABC"))).Replace("-", string.Empty).ToLowerInvariant();
        Assert.Equal(expected, Assert.Single(hashes.Payloads).Sha256); Assert.Equal(3, hashes.ReadBytes);
        Assert.Throws<OfficeInspectionResourceLimitException>(() => hashes.OpenRead(0));
        Assert.Throws<OfficeInspectionResourceLimitException>(() => result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions {
            IncludeSha256 = true, MaxPayloadBytes = 2, MaxTotalPayloadBytes = 3, MaxReadBytes = 3
        }));
    }

    [Theory]
    [InlineData(@"{\rtf1{\pict\pngblip\bin6 AB", 6L, 2L)]
    [InlineData(@"{\rtf1{\pict\pngblip 41\bin6 AB", -1L, 3L)]
    public void TruncatedDeclaredBinaryCannotPublishACompletePayloadOrHash(string source, long declaredLength, long readBytes) {
        RtfReadResult result = RtfDocument.ParseResult(source, new RtfReadOptions { DeferBinaryPayloads = true });
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF004");
        using OfficeEmbeddedPayloadInspection inspection = result.InspectEmbeddedPayloads();
        Assert.Equal(declaredLength, Assert.Single(inspection.Payloads).Length);
        Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0)); Assert.Equal(readBytes, inspection.ReadBytes);
        Assert.Throws<InvalidDataException>(() => result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions { IncludeSha256 = true }));
    }

    [Theory]
    [InlineData("41F")]
    [InlineData("41GG")]
    public void OriginalMalformedHexCannotPublishAPartialHashWhileSemanticGetterCompatibilityRemains(string hex) {
        RtfReadResult result = RtfDocument.ParseResult(@"{\rtf1 text{\pict\pngblip " + hex + "}}", new RtfReadOptions { DeferBinaryPayloads = true });
        Assert.Equal(new byte[] { 65 }, result.Document.Paragraphs[0].Inlines.OfType<RtfImage>().Single().Data);
        using OfficeEmbeddedPayloadInspection inspection = result.InspectEmbeddedPayloads();
        Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0)); Assert.Equal(1, inspection.ReadBytes);
        Assert.Throws<InvalidDataException>(() => result.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions { IncludeSha256 = true }));
    }

    [Fact]
    public void FromHtmlBinarySourceIsRefusedByDeferredProfileAndFullKeepsItsBaseline() {
        const string source = @"{\rtf1\fromhtml1 text{\pict\pngblip 414243}}";
        Assert.Throws<NotSupportedException>(() => RtfDocument.ParseResult(source, new RtfReadOptions { DeferBinaryPayloads = true }));
        RtfReadResult full = RtfDocument.ParseResult(source);
        Assert.NotNull(full.Document.HtmlEncapsulation); Assert.True(full.Document.IsHtmlEncapsulationCurrent);
        Assert.Equal(new byte[] { 65, 66, 67 }, full.Document.Paragraphs[0].Inlines.OfType<RtfImage>().Single().Data);
    }

    private static byte[] Read(Stream stream) { using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray(); }
}
