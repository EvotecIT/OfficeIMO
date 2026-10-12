using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class InstalledCffFontTests {
    [Fact]
    public void RetainedApproximateFaceDoesNotHideAnExactFaceInLaterCandidates() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.InstalledCff", Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            byte[] regular = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
            OfficeOpenTypeReader reader = OfficeOpenTypeReader.TryCreate(regular)!;
            string family = reader.ReadFamilyName()!;
            byte[] light = (byte[])regular.Clone();
            Assert.True(reader.TryGetTable("OS/2", out int offset, out _));
            light[offset + 4] = 1; light[offset + 5] = 44; // 300, with identical glyph metrics.
            string lightPath = Path.Combine(root, "123.otf"), regularPath = Path.Combine(root, "456.otf");
            File.WriteAllBytes(lightPath, light); File.WriteAllBytes(regularPath, regular);
            var fonts = new OfficeFontFaceCollection();
            Assert.True(fonts.TryAddInstalledCffFamily(family, OfficeFontFaceDescriptor.Regular,
                new[] { lightPath }, regular.Length * 4, regular.Length, CancellationToken.None, out _, out string? error), error);
            Assert.Equal(300, Assert.Single(fonts.Faces).Descriptor.Weight);
            Assert.True(fonts.TryAddInstalledCffFamily(family, OfficeFontFaceDescriptor.Regular,
                new[] { lightPath, regularPath }, regular.Length * 4, regular.Length, CancellationToken.None, out _, out error), error);
            Assert.Contains(fonts.Faces, face => face.Descriptor == OfficeFontFaceDescriptor.Regular);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void InstalledCffSelectionUsesFaceMetadataAndRetainsScopedBudgets() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.InstalledCff", Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            byte[] regular = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
            OfficeOpenTypeReader reader = OfficeOpenTypeReader.TryCreate(regular)!;
            string family = reader.ReadFamilyName()!;
            byte[] bold = (byte[])regular.Clone();
            Assert.True(reader.TryGetTable("OS/2", out int offset, out _));
            bold[offset + 4] = 2; bold[offset + 5] = 188; // 700, with identical glyph metrics.
            string regularPath = Path.Combine(root, "123.otf"), boldPath = Path.Combine(root, "456.otf");
            File.WriteAllBytes(regularPath, regular); File.WriteAllBytes(boldPath, bold);
            string[] paths = { regularPath, boldPath };
            var requested = new OfficeFontFaceDescriptor(700);
            var fonts = new OfficeFontFaceCollection();
            Assert.True(fonts.TryAddInstalledCffFamily(family, requested, paths, regular.Length * 4, regular.Length,
                CancellationToken.None, out int bytes, out string? error), error);
            OfficeFontFace face = Assert.Single(fonts.Faces);
            Assert.True(face.Program.IsOpenTypeCff);
            Assert.Equal(700, face.Descriptor.Weight);
            Assert.True(bytes > 0);
            Assert.True(face.Program.Measure("Office fonts 0123456789", 11D) > 0D);
            Assert.True(fonts.TryAddInstalledCffFamily(family, requested, paths, 1, 1,
                CancellationToken.None, out int multiCandidateReuse, out error), error);
            Assert.Equal(0, multiCandidateReuse);
            File.Delete(boldPath);
            Assert.True(fonts.TryAddInstalledCffFamily(family, requested, new[] { boldPath }, 1, 1,
                CancellationToken.None, out int reused, out error), error);
            Assert.Equal(0, reused);
            Assert.Single(fonts.Faces);
            Assert.False(new OfficeFontFaceCollection().TryAddInstalledCffFamily(family, requested,
                new[] { regularPath }, regular.Length * 4, regular.Length - 1,
                CancellationToken.None, out _, out error));
            Assert.Contains("per-resource", error);
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            Assert.Throws<OperationCanceledException>(() => fonts.TryAddInstalledCffFamily(family, requested,
                paths, int.MaxValue, long.MaxValue, canceled.Token, out _, out _));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }
}
