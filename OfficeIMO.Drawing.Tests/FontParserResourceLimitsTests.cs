using System.IO;
using System.IO.Compression;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class FontParserResourceLimitsTests {
    [Fact]
    public void SharedCff2SubroutineReferencesRemainLoadable() {
        // Counting the same 16,384-entry index for all 256 references would exceed the limit.
        var font = OfficeOpenTypeCffFont.TryLoad(ReadFont("TestSharedCff2.otf.gz"), null, out string? error);

        Assert.Null(error);
        Assert.NotNull(font);
        Assert.Equal(2.4D, font!.Measure("A", 4D), 6);
    }

    [Fact]
    public void DistinctCff2IndicesHaveACumulativeParsingLimit() {
        var font = OfficeOpenTypeCffFont.TryLoad(ReadFont("TestCumulativeCff2.otf.gz"), null, out string? error);

        Assert.Null(font);
        Assert.Contains("cumulative", error);
    }

    [Fact]
    public void NonmatchingCoverageRangesConsumeTheShapingWorkAllowance() {
        var request = new OfficeTextShapingRequest(
            new string('A', 512), "Coverage limit", ReadFont("TestSubstitutionCoverage.ttf.gz"),
            isOpenTypeCff: false, unitsPerEm: 1000,
            direction: OfficeTextDirection.LeftToRight, language: null,
            featureSettings: new OfficeTextFeatureSettings().With("liga", 1));

        Assert.Throws<InvalidDataException>(() => OfficeManagedTextShapingProvider.Instance.ShapeText(request));
    }

    private static byte[] ReadFont(string name) {
        using var stream = typeof(FontParserResourceLimitsTests).Assembly.GetManifestResourceStream(
            "OfficeIMO.Drawing.Tests.TestAssets.FontParserLimits." + name);
        Assert.NotNull(stream);
        using var decompressor = new GZipStream(stream!, CompressionMode.Decompress);
        using var output = new MemoryStream();
        decompressor.CopyTo(output);
        return output.ToArray();
    }
}
