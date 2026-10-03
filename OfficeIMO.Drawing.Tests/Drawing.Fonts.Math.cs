using System;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingTests {
    [Fact]
    public void GenericMathFamilyPrefersSuppliedGenericThenCanonicalNamedFaces() {
        byte[] data = ManagedTextShapingTestAssets.CreateFont('x');
        var fonts = new OfficeFontFaceCollection();
        fonts.Add("Latin Modern Math", data);
        fonts.Add("STIX Two Math", data);
        fonts.Add("Cambria Math", data);
        Assert.True(fonts.TryResolveFaceForText("x", "math", OfficeFontStyle.Regular, out OfficeFontFace? named));
        Assert.Equal("STIX Two Math", named!.FamilyName);
        fonts.Add("math", data);
        Assert.True(fonts.TryResolveFaceForText("x", "math", OfficeFontStyle.Regular, out OfficeFontFace? generic));
        Assert.Equal("math", generic!.FamilyName);
    }

    [Fact]
    public void InstalledMathFontRetainsOperationBoundsAndCancellationWhenAvailable() {
        var fonts = new OfficeFontFaceCollection();
        if (!fonts.TryAddInstalledFamily("math", OfficeFontFaceDescriptor.Regular, 32 * 1024 * 1024,
                CancellationToken.None, out int bytes, out _, 16 * 1024 * 1024)) return;
        Assert.True(bytes > 0);
        Assert.True(fonts.TryResolveFaceForText("x2", "math", OfficeFontStyle.Regular, out OfficeFontFace? firstFace));
        // Allow the original collection as well as the retained standalone face, but
        // do not charge a speculative CFF copy before reaching the TrueType route.
        int exactBudget = bytes;
        if (!firstFace!.Program.IsOpenTypeCff
            && OfficeTrueTypeFont.TryLoadFontFamily("math", out string? sourcePath) != null
            && sourcePath != null) {
            exactBudget = Math.Max(bytes, checked((int)new System.IO.FileInfo(sourcePath).Length));
        }
        var exact = new OfficeFontFaceCollection();
        Assert.True(exact.TryAddInstalledFamily("math", OfficeFontFaceDescriptor.Regular,
            exactBudget, CancellationToken.None, out int exactBytes, out string? exactError,
            16 * 1024 * 1024), exactError);
        Assert.Equal(bytes, exactBytes);
        // A second request for another style must not charge the same retained static
        // program again; this also protects the following fonts' operation-wide budget.
        Assert.True(fonts.TryAddInstalledFamily("math", new OfficeFontFaceDescriptor(700, 100D, OfficeFontSlant.Italic),
            1, CancellationToken.None, out int repeatedBytes, out _, 1));
        Assert.Equal(0, repeatedBytes);
        if (firstFace.Program.IsOpenTypeCff) {
            var configured = new OfficeFontFaceCollection();
            int resolverCalls = 0;
            configured.FontVariationResolver = _ => { resolverCalls++; return null; };
            Assert.True(configured.TryAddInstalledFamily("math", OfficeFontFaceDescriptor.Regular, 32 * 1024 * 1024,
                CancellationToken.None, out int configuredFirstBytes, out _));
            Assert.True(configured.TryAddInstalledFamily("math", new OfficeFontFaceDescriptor(700, 100D, OfficeFontSlant.Italic),
                32 * 1024 * 1024, CancellationToken.None, out int configuredSecondBytes, out _));
            Assert.Equal(2, resolverCalls);
            Assert.True(configuredFirstBytes > 0 && configuredSecondBytes > 0);
        }
        var limited = new OfficeFontFaceCollection();
        Assert.False(limited.TryAddInstalledFamily("math", OfficeFontFaceDescriptor.Regular, 32 * 1024 * 1024,
            CancellationToken.None, out _, out string? sourceError, 1));
        Assert.Contains("per-resource", sourceError);
        Assert.Empty(limited.Faces);
        Assert.False(limited.TryAddInstalledFamily("math", OfficeFontFaceDescriptor.Regular, 1,
            CancellationToken.None, out _, out string? decodedError));
        Assert.Contains("byte limit", decodedError);
        Assert.Empty(limited.Faces);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => limited.TryAddInstalledFamily("math", OfficeFontFaceDescriptor.Regular,
            32 * 1024 * 1024, cancellation.Token, out _, out _));
    }
}
