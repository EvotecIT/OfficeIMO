using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeCloudFontDiscoveryTests {
    [Theory]
    [InlineData("等线", 1, false)]
    [InlineData("等线 Light", 16, false)]
    [InlineData("等线", 1, true)]
    [InlineData("等线 Light", 16, true)]
    public void CachedFamilyDiscoveryUsesLocalizedMetadataOutsideTheFolderName(string alias, int nameId, bool collection) {
        string cloud = Path.Combine(Path.GetTempPath(), "OfficeIMO.CloudFontDiscovery", Guid.NewGuid().ToString("N"));
        try {
            string folder = Directory.CreateDirectory(Path.Combine(cloud, "DengXian")).FullName;
            string path = Path.Combine(folder, collection ? "123456789.ttc" : "123456789.ttf");
            byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
            byte[] font = AddFamilyAlias(source, alias, (ushort)nameId);
            File.WriteAllBytes(path, collection ? ToCollection(font) : font);
            string key = string.Concat(alias.Where(char.IsLetterOrDigit)).ToLowerInvariant();
            Assert.Equal(new[] { path }, OfficeCloudFontFiles.EnumerateFamilyFiles(cloud, key));
            Assert.Empty(OfficeCloudFontFiles.EnumerateFamilyFiles(cloud, "unrelated"));
            Assert.True(OfficeTrueTypeFont.TryLoad(path)!.MatchesFamilyName(alias));
        } finally {
            if (Directory.Exists(cloud)) Directory.Delete(cloud, recursive: true);
        }
    }

    private static byte[] ToCollection(byte[] source) {
        byte[] collection = new byte[source.Length + 16];
        Array.Copy(source, 0, collection, 16, source.Length);
        Array.Copy(Encoding.ASCII.GetBytes("ttcf"), collection, 4);
        collection[5] = 1; collection[11] = 1; collection[15] = 16;
        int count = source[4] * 256 + source[5];
        for (int index = 0; index < count; index++) {
            int at = 16 + 12 + index * 16 + 8;
            uint offset = ((uint)collection[at] << 24) | ((uint)collection[at + 1] << 16) | ((uint)collection[at + 2] << 8) | collection[at + 3];
            offset += 16;
            collection[at] = (byte)(offset >> 24); collection[at + 1] = (byte)(offset >> 16);
            collection[at + 2] = (byte)(offset >> 8); collection[at + 3] = (byte)offset;
        }
        return collection;
    }

    private static byte[] AddFamilyAlias(byte[] source, string alias, ushort nameId) {
        int Read16(int at) => source[at] * 256 + source[at + 1];
        int Read32(int at) => checked((int)(((uint)source[at] << 24) | ((uint)source[at + 1] << 16) | ((uint)source[at + 2] << 8) | source[at + 3]));
        int record = Enumerable.Range(0, Read16(4)).Select(index => 12 + index * 16)
            .Single(at => Encoding.ASCII.GetString(source, at, 4) == "name");
        int oldOffset = Read32(record + 8), oldLength = Read32(record + 12);
        int count = Read16(oldOffset + 2), storage = Read16(oldOffset + 4), end = 6 + count * 12;
        byte[] name = Encoding.BigEndianUnicode.GetBytes(alias);
        int next = (source.Length + 3) & ~3;
        byte[] result = new byte[next + oldLength + 12 + name.Length];
        Array.Copy(source, result, source.Length);
        Array.Copy(source, oldOffset, result, next, end);
        Array.Copy(source, oldOffset + end, result, next + end + 12, oldLength - end);
        void Write16(int at, int value) { result[at] = (byte)(value >> 8); result[at + 1] = (byte)value; }
        void Write32(int at, int value) { result[at] = (byte)(value >> 24); result[at + 1] = (byte)(value >> 16); Write16(at + 2, value); }
        Write16(next + 2, count + 1); Write16(next + 4, storage + 12);
        int aliasRecord = next + end;
        Write16(aliasRecord, 3); Write16(aliasRecord + 2, 1); Write16(aliasRecord + 4, 0x0804);
        Write16(aliasRecord + 6, nameId); Write16(aliasRecord + 8, name.Length); Write16(aliasRecord + 10, oldLength - storage);
        Array.Copy(name, 0, result, next + oldLength + 12, name.Length);
        Write32(record + 8, next); Write32(record + 12, oldLength + 12 + name.Length);
        return result;
    }

    [Fact]
    public void CachedFamilyDiscoveryAcceptsNumericFontFilenamesAndKeepsFamiliesSeparate() {
        string cache = Path.Combine(Path.GetTempPath(), "OfficeIMO.CloudFontDiscovery", Guid.NewGuid().ToString("N"));
        try {
            Assert.Empty(OfficeCloudFontFiles.EnumerateRoots(cache));
            string cloud = Path.Combine(cache, "4", "CloudFonts");
            string family = Path.Combine(cloud, "Example Sans");
            string other = Path.Combine(cloud, "Other Family");
            Directory.CreateDirectory(family);
            Directory.CreateDirectory(other);
            string numeric = Path.Combine(family, "123456789.ttf");
            // Discovery returns candidates. The existing font loader validates
            // their programs and declared family before selecting a face.
            File.WriteAllBytes(numeric, new byte[] { 0, 1 });
            File.WriteAllBytes(Path.Combine(family, "metadata.json"), new byte[] { 0, 1 });
            File.WriteAllBytes(Path.Combine(other, "987654321.ttf"), new byte[] { 0, 1 });

            Assert.Equal(new[] { cloud }, OfficeCloudFontFiles.EnumerateRoots(cache).ToArray());
            Assert.Equal(new[] { numeric }, OfficeCloudFontFiles.EnumerateFamilyFiles(cloud, "examplesans").ToArray());
            Assert.Empty(OfficeCloudFontFiles.EnumerateFamilyFiles(cloud, "missingfamily"));
        } finally {
            if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
        }
    }
}
