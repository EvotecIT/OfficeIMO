using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSystemFontNameTests {
    [Theory]
    [InlineData("Tipus de lletra del sistema")]
    [InlineData("System Font")]
    [InlineData(".SF NS")]
    [InlineData("System Font Display")]
    [InlineData("Localized Full Name")]
    [InlineData("SystemFont-Regular")]
    public void LocalizedAndPlatformNamesResolveTheSameFontProgram(string requestedName) {
        byte[] font = WithNames(
            (3, 1027, 1, "Tipus de lletra del sistema"),
            (3, 1033, 1, "System Font"),
            (0, 0, 1, ".SF NS"),
            (3, 1027, 4, "Tipus de lletra del sistema"),
            (3, 1027, 4, "Localized Full Name"),
            (3, 1033, 4, "System Font Display"),
            (3, 1033, 6, "SystemFont-Regular"));
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.FontNames." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "unrelated-file-name.ttf");
            File.WriteAllBytes(path, font);
            Assert.True(PdfEmbeddedFontFamily.TryFromSystemFontFiles(requestedName, new[] { path }, out var family));
            Assert.Equal(font, family!.Regular);
            Assert.True(PdfEmbeddedFontFamily.TryResolveSystemFaceFromFiles(requestedName, new[] { path },
                new OfficeFontFaceDescriptor(400, 100D, OfficeFontSlant.Normal), "A", out var selected));
            Assert.Equal(font, selected!.Regular);
            Assert.False(PdfEmbeddedFontFamily.TryFromSystemFontFiles("Missing Family", new[] { path }, out _));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactFamilyWinsOverSameDescriptorTypographicRelative(bool reverseOrder) {
        byte[] exact = WithNames((3, 1033, 1, "Shared Family"), (3, 1033, 2, "Regular"),
            (3, 1033, 4, "Shared Family Regular"), (3, 1033, 16, "Shared Family"));
        byte[] relative = WithNames((3, 1033, 1, "Shared Family Ornaments"), (3, 1033, 2, "Regular"),
            (3, 1033, 4, "Shared Family Ornaments"), (3, 1033, 16, "Shared Family"));
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.FontFamilyPriority." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string exactPath = Path.Combine(directory, "exact.ttf");
            string relativePath = Path.Combine(directory, "relative.ttf");
            File.WriteAllBytes(exactPath, exact);
            File.WriteAllBytes(relativePath, relative);
            string[] paths = reverseOrder ? new[] { relativePath, exactPath } : new[] { exactPath, relativePath };
            Assert.True(PdfEmbeddedFontFamily.TryFromSystemFontFiles("Shared Family", paths, out var family));
            Assert.Equal(exact, family!.Regular);
            Assert.True(PdfEmbeddedFontFamily.TryResolveSystemFaceFromFiles("Shared Family", paths,
                new OfficeFontFaceDescriptor(400, 100D, OfficeFontSlant.Normal), "A", out var selected));
            Assert.Equal(exact, selected!.Regular);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RetainedNameBudgetCountsDistinctAliases(bool distinct) {
        var names = Enumerable.Range(0, PdfEmbeddedFontFamily.MaxSystemFontNameAliases + 1)
            .Select(index => (3, 1033, 1, distinct ? "Alias" + index : "Alias0")).ToArray();
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.FontNames." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "unrelated.ttf");
            File.WriteAllBytes(path, WithNames(names));
            Assert.Equal(!distinct, PdfEmbeddedFontFamily.TryFromSystemFontFiles("Alias0", new[] { path }, out _));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    internal static byte[] WithNames(params (int Platform, int Language, int Id, string Text)[] names) {
        byte[][] strings = names.Select(name => Encoding.BigEndianUnicode.GetBytes(name.Text)).ToArray();
        int stringOffset = 6 + names.Length * 12;
        byte[] table = new byte[stringOffset + strings.Sum(value => value.Length)];
        WriteUInt16(table, 2, names.Length);
        WriteUInt16(table, 4, stringOffset);
        int cursor = stringOffset;
        for (int index = 0; index < names.Length; index++) {
            int record = 6 + index * 12;
            WriteUInt16(table, record, names[index].Platform);
            WriteUInt16(table, record + 2, names[index].Platform == 3 ? 1 : 4);
            WriteUInt16(table, record + 4, names[index].Language);
            WriteUInt16(table, record + 6, names[index].Id);
            WriteUInt16(table, record + 8, strings[index].Length);
            WriteUInt16(table, record + 10, cursor - stringOffset);
            strings[index].CopyTo(table, cursor);
            cursor += strings[index].Length;
        }

        byte[] source = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] result = new byte[source.Length + table.Length];
        source.CopyTo(result, 0);
        table.CopyTo(result, source.Length);
        int tableCount = (source[4] << 8) | source[5];
        for (int index = 0; index < tableCount; index++) {
            int record = 12 + index * 16;
            if (Encoding.ASCII.GetString(source, record, 4) != "name") continue;
            WriteUInt32(result, record + 8, source.Length);
            WriteUInt32(result, record + 12, table.Length);
            return result;
        }
        throw new InvalidOperationException("Test font has no name table.");
    }

    private static void WriteUInt16(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteUInt32(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
