using System.Buffers.Binary;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class RasterContentSafetyTests {
    [Fact]
    public async Task InspectSendsMetadataFreeNormalizedPngToOcr() {
        byte[] image = CreateImage(20, 10, OfficeColor.White, null, null);
        byte[]? normalized = null;
        IOcrEngine engine = CreateEngine(request => {
            normalized = request.Payload;
            return new OcrResult();
        });

        await OfficeRasterContentSafety.InspectAsync(image, engine);

        Assert.NotNull(normalized);
        Assert.False(ContainsAscii(normalized!, "pHYs"));
    }

    [Fact]
    public async Task RedactionReturnsMetadataFreePng() {
        byte[] image = CreateImage(
            24,
            12,
            OfficeColor.White,
            new PixelBox(3, 3, 12, 5),
            OfficeColor.FromRgb(248, 248, 248));
        IOcrEngine engine = CreateEngine(request =>
            OfficeRasterImageDecoder.TryDecode(request.Payload, out OfficeRasterImage? raster) &&
            raster != null && raster.GetPixel(3, 3) == OfficeColor.Black
                ? new OcrResult()
                : Result("concealed", new OcrRegion { X = 3, Y = 3, Width = 12, Height = 5 }, 0.99D));
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionColor = OfficeColor.Black
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        OfficeContentCleanupResult cleanup = await OfficeRasterContentSafety.RedactSelectedContentAsync(
            image,
            engine,
            new OfficeContentCleanupSelection(new[] { finding.Id }),
            options);

        Assert.True(cleanup.Changed);
        Assert.False(ContainsAscii(cleanup.Output, "pHYs"));
    }

    [Fact]
    public async Task EmptySelectionKeepsAnIndependentInputSnapshot() {
        byte[] image = CreateImage(20, 10, OfficeColor.White, null, null);
        byte expectedFirstByte = image[0];
        var options = new OfficeRasterContentSafetyOptions { EnableOpaqueRectangleRedaction = true };

        OfficeContentCleanupResult cleanup = await OfficeRasterContentSafety.RedactSelectedContentAsync(
            image,
            CreateEngine(_ => new OcrResult()),
            new OfficeContentCleanupSelection(Array.Empty<string>()),
            options);
        image[0] ^= byte.MaxValue;

        Assert.False(cleanup.Changed);
        Assert.Equal(expectedFirstByte, cleanup.Output[0]);
    }

    [Fact]
    public async Task InspectAcceptsCanonicalSrgbTiffColorimetry() {
        byte[] tiff = OfficeTiffCodec.Encode(new OfficeRasterImage(8, 8, OfficeColor.White));
        byte[] canonical = AddSrgbTiffColorimetry(tiff, corruptTransferFunction: false);

        OfficeContentSafetyReport report = await OfficeRasterContentSafety.InspectAsync(
            canonical,
            CreateEngine(_ => new OcrResult()));

        Assert.Empty(report.Findings);
    }

    [Fact]
    public async Task InspectRejectsNonCanonicalTiffTransferFunction() {
        byte[] tiff = OfficeTiffCodec.Encode(new OfficeRasterImage(8, 8, OfficeColor.White));
        byte[] nonCanonical = AddSrgbTiffColorimetry(tiff, corruptTransferFunction: true);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRasterContentSafety.InspectAsync(
                nonCanonical,
                CreateEngine(_ => new OcrResult())));
    }

    [Fact]
    public async Task RedactionIgnoresBoundedWhitespaceDuringOverlapVerification() {
        byte[] image = CreateImage(
            24,
            12,
            OfficeColor.White,
            new PixelBox(3, 3, 12, 5),
            OfficeColor.FromRgb(248, 248, 248));
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? new OcrResult {
                Text = "concealed ",
                Spans = new[] {
                    Span(0, "concealed", new OcrRegion { X = 3, Y = 3, Width = 12, Height = 5 }, 0.99D),
                    Span(1, " ", new OcrRegion { X = 10, Y = 3, Width = 1, Height = 5 }, 0.99D,
                        OcrTextSpanLevel.Character)
                }
            }
            : new OcrResult {
                Text = " ",
                Spans = new[] {
                    Span(0, " ", new OcrRegion { X = 10, Y = 3, Width = 1, Height = 5 }, 0.99D,
                        OcrTextSpanLevel.Character)
                }
            });
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionColor = OfficeColor.Black
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        OfficeContentCleanupResult cleanup = await OfficeRasterContentSafety.RedactSelectedContentAsync(
            image,
            engine,
            new OfficeContentCleanupSelection(new[] { finding.Id }),
            options);

        Assert.True(cleanup.Changed);
    }

    [Fact]
    public async Task InstructionDetectionBreaksAtVisibleMixedGranularityText() {
        var raster = new OfficeRasterImage(32, 12, OfficeColor.White);
        for (int y = 3; y < 8; y++) {
            for (int x = 2; x < 8; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
            raster.SetPixel(10, y, OfficeColor.Black);
            for (int x = 12; x < 20; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
        }
        byte[] image = OfficePngWriter.Encode(raster);
        IOcrEngine engine = CreateEngine(_ => new OcrResult {
            Text = "ignore x previous",
            Spans = new[] {
                Span(0, "ignore", new OcrRegion { X = 2, Y = 3, Width = 6, Height = 5 }, 0.99D),
                Span(1, "x", new OcrRegion { X = 10, Y = 3, Width = 1, Height = 5 }, 0.99D,
                    OcrTextSpanLevel.Character),
                Span(2, "previous", new OcrRegion { X = 12, Y = 3, Width = 8, Height = 5 }, 0.99D)
            }
        });

        OfficeContentSafetyReport report = await OfficeRasterContentSafety.InspectAsync(image, engine);

        Assert.Equal(2, report.Findings.Count);
        Assert.False(report.HasPotentiallyDangerousContent);
        Assert.All(report.Findings, finding => Assert.False(finding.IsInstructionLike));
    }

    private static byte[] AddSrgbTiffColorimetry(byte[] tiff, bool corruptTransferFunction) {
        if (tiff.Length < 8 || tiff[0] != (byte)'I' || tiff[1] != (byte)'I') {
            throw new InvalidDataException("The test helper requires a little-endian classic TIFF.");
        }
        int oldIfdOffset = BinaryPrimitives.ReadInt32LittleEndian(tiff.AsSpan(4, 4));
        int oldEntryCount = BinaryPrimitives.ReadUInt16LittleEndian(tiff.AsSpan(oldIfdOffset, 2));
        int oldEntriesOffset = oldIfdOffset + 2;
        if (oldEntryCount < 1 || oldEntriesOffset > tiff.Length - oldEntryCount * 12 - 4) {
            throw new InvalidDataException("The source TIFF IFD is malformed.");
        }

        var entries = new List<(ushort Tag, byte[] Bytes)>(oldEntryCount + 3);
        for (int index = 0; index < oldEntryCount; index++) {
            byte[] entry = tiff.AsSpan(oldEntriesOffset + index * 12, 12).ToArray();
            entries.Add((BinaryPrimitives.ReadUInt16LittleEndian(entry), entry));
        }

        int newIfdOffset = (tiff.Length + 1) & ~1;
        int newIfdLength = 2 + (oldEntryCount + 3) * 12 + 4;
        int transferOffset = newIfdOffset + newIfdLength;
        int whitePointOffset = transferOffset + 256 * 2;
        int primariesOffset = whitePointOffset + 2 * 8;
        byte[] result = new byte[primariesOffset + 6 * 8];
        Buffer.BlockCopy(tiff, 0, result, 0, tiff.Length);
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(4, 4), newIfdOffset);

        entries.Add((301, CreateTiffEntry(301, type: 3, count: 256, transferOffset)));
        entries.Add((318, CreateTiffEntry(318, type: 5, count: 2, whitePointOffset)));
        entries.Add((319, CreateTiffEntry(319, type: 5, count: 6, primariesOffset)));
        entries.Sort((left, right) => left.Tag.CompareTo(right.Tag));
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(newIfdOffset, 2), checked((ushort)entries.Count));
        for (int index = 0; index < entries.Count; index++) {
            entries[index].Bytes.CopyTo(result, newIfdOffset + 2 + index * 12);
        }

        for (int index = 0; index < 256; index++) {
            double encoded = index / 255D;
            double linear = encoded <= 0.04045D
                ? encoded / 12.92D
                : Math.Pow((encoded + 0.055D) / 1.055D, 2.4D);
            ushort value = checked((ushort)Math.Round(
                linear * ushort.MaxValue,
                MidpointRounding.AwayFromZero));
            BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(transferOffset + index * 2, 2), value);
        }
        if (corruptTransferFunction) {
            int changed = transferOffset + 128 * 2;
            ushort value = BinaryPrimitives.ReadUInt16LittleEndian(result.AsSpan(changed, 2));
            BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(changed, 2), checked((ushort)(value + 2)));
        }

        WriteTiffRationals(result, whitePointOffset, new[] { 3127, 3290 }, new[] { 10000, 10000 });
        WriteTiffRationals(
            result,
            primariesOffset,
            new[] { 640, 330, 300, 600, 150, 60 },
            new[] { 1000, 1000, 1000, 1000, 1000, 1000 });
        return result;
    }

    private static byte[] CreateTiffEntry(ushort tag, ushort type, uint count, int valueOffset) {
        var entry = new byte[12];
        BinaryPrimitives.WriteUInt16LittleEndian(entry.AsSpan(0, 2), tag);
        BinaryPrimitives.WriteUInt16LittleEndian(entry.AsSpan(2, 2), type);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.AsSpan(4, 4), count);
        BinaryPrimitives.WriteInt32LittleEndian(entry.AsSpan(8, 4), valueOffset);
        return entry;
    }

    private static void WriteTiffRationals(
        byte[] destination,
        int offset,
        IReadOnlyList<int> numerators,
        IReadOnlyList<int> denominators) {
        Assert.Equal(numerators.Count, denominators.Count);
        for (int index = 0; index < numerators.Count; index++) {
            BinaryPrimitives.WriteInt32LittleEndian(
                destination.AsSpan(offset + index * 8, 4),
                numerators[index]);
            BinaryPrimitives.WriteInt32LittleEndian(
                destination.AsSpan(offset + index * 8 + 4, 4),
                denominators[index]);
        }
    }
}
