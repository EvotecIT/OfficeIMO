using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSystemFontStyleTests {
    private const string MacFontDirectory = "/System/Library/Fonts/Supplemental/";

    [Fact]
    public void InstalledTrebuchetFacesKeepTheirBoldAndItalicPrograms() {
        string[] paths = {
            MacFontDirectory + "Trebuchet MS.ttf",
            MacFontDirectory + "Trebuchet MS Bold.ttf",
            MacFontDirectory + "Trebuchet MS Italic.ttf",
            MacFontDirectory + "Trebuchet MS Bold Italic.ttf"
        };
        if (Array.Exists(paths, path => !File.Exists(path))) return;

        Assert.True(PdfEmbeddedFontFamily.TryFromSystemFontFiles("Trebuchet MS", paths, out PdfEmbeddedFontFamily? family));
        Assert.Equal(File.ReadAllBytes(paths[1]), family!.Bold);
        Assert.Equal(File.ReadAllBytes(paths[2]), family.Italic);
        Assert.Equal(File.ReadAllBytes(paths[3]), family.BoldItalic);
    }

    [Fact]
    public void InstalledFamilyPrefersRegularAndBoldOverAdjacentWeightsRegardlessOfFileOrder() {
        string regularPath = MacFontDirectory + "Trebuchet MS.ttf";
        string boldPath = MacFontDirectory + "Trebuchet MS Bold.ttf";
        if (!File.Exists(regularPath) || !File.Exists(boldPath)) return;

        byte[] regular = File.ReadAllBytes(regularPath);
        byte[] bold = File.ReadAllBytes(boldPath);
        string directory = CreateTestDirectory();
        try {
            string lightPath = Path.Combine(directory, "light.ttf");
            string semiBoldPath = Path.Combine(directory, "semibold.ttf");
            File.WriteAllBytes(lightPath, WithOs2Style(regular, weight: 300));
            File.WriteAllBytes(semiBoldPath, WithOs2Style(bold, weight: 600));

            Assert.True(PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                "Trebuchet MS", new[] { lightPath, regularPath, semiBoldPath, boldPath },
                out PdfEmbeddedFontFamily? family));
            Assert.Equal(regular, family!.Regular);
            Assert.Equal(bold, family.Bold);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ObliqueSelectionFlagKeepsFaceSlantedWithoutItalicFlag() {
        string regularPath = MacFontDirectory + "Arial.ttf";
        string italicPath = MacFontDirectory + "Arial Italic.ttf";
        if (!File.Exists(regularPath) || !File.Exists(italicPath)) return;

        byte[] oblique = WithOs2Style(File.ReadAllBytes(italicPath), version: 4, selection: 0x200);
        int head = FindTableOffset(oblique, "head");
        WriteUInt16(oblique, head + 44, (ushort)(ReadUInt16(oblique, head + 44) & ~0x0002));
        string directory = CreateTestDirectory();
        try {
            string obliquePath = Path.Combine(directory, "oblique.ttf");
            File.WriteAllBytes(obliquePath, oblique);
            string[] paths = { regularPath, obliquePath };

            Assert.True(PdfEmbeddedFontFamily.TryFromSystemFontFiles("Arial", paths, out PdfEmbeddedFontFamily? family));
            Assert.Equal(oblique, family!.Italic);
            Assert.True(PdfEmbeddedFontFamily.TryResolveSystemFaceFromFiles(
                "Arial", paths, new OfficeFontFaceDescriptor(400, 100D, OfficeFontSlant.Oblique),
                "A", out PdfEmbeddedFontFamily? selected));
            Assert.Equal(oblique, selected!.Regular);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string CreateTestDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.FontStyles." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static byte[] WithOs2Style(byte[] source, ushort? version = null, ushort? weight = null, ushort? selection = null) {
        byte[] copy = (byte[])source.Clone();
        int os2 = FindTableOffset(copy, "OS/2");
        if (version.HasValue) WriteUInt16(copy, os2, version.Value);
        if (weight.HasValue) WriteUInt16(copy, os2 + 4, weight.Value);
        if (selection.HasValue) WriteUInt16(copy, os2 + 62, selection.Value);
        return copy;
    }

    private static int FindTableOffset(byte[] data, string tag) {
        int count = ReadUInt16(data, 4);
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (data[record] == tag[0] && data[record + 1] == tag[1]
                && data[record + 2] == tag[2] && data[record + 3] == tag[3]) {
                return (int)((uint)data[record + 8] << 24 | (uint)data[record + 9] << 16
                    | (uint)data[record + 10] << 8 | data[record + 11]);
            }
        }
        throw new InvalidOperationException("Required font table was not found: " + tag);
    }

    private static ushort ReadUInt16(byte[] data, int offset) =>
        (ushort)((data[offset] << 8) | data[offset + 1]);

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }
}
