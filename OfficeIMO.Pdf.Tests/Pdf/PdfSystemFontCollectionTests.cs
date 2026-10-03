using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSystemFontCollectionTests {
    [Fact]
    public void TrueTypeSelectionSkipsStandaloneAndCollectionCffPrograms() {
        byte[] cff = LoadCff();
        Assert.Empty(OfficeTrueTypeCollection.ExtractPrograms(cff));
        byte[] collection = CreateCollection(cff, cff, cff);
        Assert.Empty(OfficeTrueTypeCollection.ExtractPrograms(collection));
        // Extraction rebuilds SFNT offsets; decoded glyph behavior is the contract.
        byte[] extracted = OfficeTrueTypeCollection.ExtractFace(collection, 1, cff.Length);
        var original = OfficeOpenTypeCffFont.TryLoad(cff, null, out _);
        var rebuilt = OfficeOpenTypeCffFont.TryLoad(extracted, null, out _);
        Assert.NotNull(original);
        Assert.NotNull(rebuilt);
        Assert.True(rebuilt.HasGlyphs("AŁéΩ"));
        Assert.Equal(original.Measure("AŁéΩ", 12), rebuilt.Measure("AŁéΩ", 12));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MixedCollectionRetainsTheSupportedTrueTypeFace(bool cffFirst) {
        byte[] cff = LoadCff();
        byte[] trueType = PdfSystemFontNameTests.WithNames(
            (3, 1033, 1, ManagedTextShapingTestAssets.FamilyName), (3, 1033, 2, "Regular"));
        byte[] collection = cffFirst ? CreateCollection(cff, trueType) : CreateCollection(trueType, cff);
        byte[] selected = Assert.Single(OfficeTrueTypeCollection.ExtractPrograms(collection));
        Assert.True(OfficeTrueTypeFont.TryLoad(selected)!.HasGlyphs("A"));

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.MixedFontCollection." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "mixed.ttc");
            File.WriteAllBytes(path, collection);
            Assert.True(PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                ManagedTextShapingTestAssets.FamilyName, new[] { path }, out var family));
            Assert.True(OfficeTrueTypeFont.TryLoad(family!.Regular)!.HasGlyphs("A"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void SkippedCffFaceStillValidatesItsTableRanges() {
        byte[] collection = CreateCollection(LoadCff(), ManagedTextShapingTestAssets.CreateFont('A'));
        int face = (int)ReadUInt32(collection, 12);
        WriteUInt32(collection, face + 12 + 8, (uint)collection.Length);
        Assert.Throws<NotSupportedException>(() => OfficeTrueTypeCollection.ExtractPrograms(collection));
    }

    private static byte[] LoadCff() => File.ReadAllBytes(
        PdfComplianceTestFonts.FindBundledOpenTypeCffFont()
            ?? throw new InvalidOperationException("The bundled independent CFF font fixture is required."));

    private static byte[] CreateCollection(params byte[][] faces) {
        int cursor = Align4(12 + faces.Length * 4);
        var offsets = new int[faces.Length];
        for (int i = 0; i < faces.Length; i++) {
            offsets[i] = cursor;
            cursor = Align4(cursor + faces[i].Length);
        }
        byte[] collection = new byte[cursor];
        WriteUInt32(collection, 0, 0x74746366);
        WriteUInt32(collection, 4, 0x00010000);
        WriteUInt32(collection, 8, (uint)faces.Length);
        for (int i = 0; i < faces.Length; i++) {
            int target = offsets[i];
            WriteUInt32(collection, 12 + i * 4, (uint)target);
            Array.Copy(faces[i], 0, collection, target, faces[i].Length);
            int count = (faces[i][4] << 8) | faces[i][5];
            for (int table = 0; table < count; table++) {
                int record = 12 + table * 16;
                WriteUInt32(collection, target + record + 8, (uint)target + ReadUInt32(faces[i], record + 8));
            }
        }
        return collection;
    }

    private static int Align4(int value) => checked((value + 3) & ~3);
    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
