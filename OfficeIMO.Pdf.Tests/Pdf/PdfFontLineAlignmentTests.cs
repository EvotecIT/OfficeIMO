using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFontLineAlignmentTests {
    [Theory]
    [InlineData("Top", 7.08D, 14.16D)]
    [InlineData("Center", 2.76D, 5.4D)]
    [InlineData("Bottom", -.84D, -1.8D)]
    [InlineData("Baseline", 0D, 0D)]
    public void ExactLineAlignmentUsesEachRunDeviceMetrics(string alignmentName, double small, double large) {
        var alignment = (PdfFontLineAlignment)Enum.Parse(typeof(PdfFontLineAlignment), alignmentName);
        var options = new PdfCreateOptions();
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Device metrics", ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
        var doc = PdfDocument.Create(options);
        doc.Paragraph(p => p.FontFamily("Device metrics").FontSize(8).Text("A")
            .FontSize(16).Text("B"), style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exactly(30).WithFixedLineBoxBaseline(24).WithFontAlignment(alignment, .12D)
            });
        using var pdf = PdfPigDocument.Open(doc.ToBytes());
        var letters = pdf.GetPage(1).Letters;
        var a = Assert.Single(letters, x => x.Value == "A");
        var b = Assert.Single(letters, x => x.Value == "B");
        Assert.InRange(a.StartBaseLine.Y - b.StartBaseLine.Y, large - small - .002D, large - small + .002D);
    }

    [Theory]
    [InlineData(8, 60, 15)]
    [InlineData(16, 119, 30)]
    [InlineData(12, 80, 20)]
    public void DeviceMetricsUseMatchingHeightsAndScaleMissingHeights(double size, double ascent, double descent) {
        var metrics = OfficeOpenTypeLineMetrics.TryRead(ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics());
        Assert.True(metrics.HasValue);
        metrics!.Value.GetDeviceMetrics(size, .12D, out double actualAscent, out double actualDescent);
        Assert.Equal(ascent, actualAscent);
        Assert.Equal(descent, actualDescent);
    }

    [Theory]
    [InlineData(0, 0, 54D, 13D)]
    [InlineData(0, 1, 60D, 15D)]
    [InlineData(0, 2, 54D, 13D)]
    [InlineData(1, 0, 60D, 15D)]
    [InlineData(1, 1, 60D, 15D)]
    [InlineData(1, 2, 54D, 13D)]
    public void DeviceMetricsRespectVersionSpecificCharacterSets(int version, int characterSet, double ascent, double descent) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics();
        var reader = OfficeOpenTypeReader.TryCreate(font)!;
        Assert.True(reader.TryGetTable("VDMX", out int offset, out _));
        font[offset + 1] = (byte)version;
        font[offset + 6] = (byte)characterSet;
        var metrics = OfficeOpenTypeLineMetrics.TryRead(font);
        Assert.True(metrics.HasValue);
        metrics!.Value.GetDeviceMetrics(8, .12D, out double actualAscent, out double actualDescent);
        Assert.Equal(ascent, actualAscent);
        Assert.Equal(descent, actualDescent);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 2)]
    public void DeviceMetricsSkipUnsupportedCharacterSetsBeforeAMatchingGroup(int version, int firstCharacterSet) {
        byte[] source = ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics();
        var reader = OfficeOpenTypeReader.TryCreate(source)!;
        Assert.True(reader.TryGetTable("VDMX", out int offset, out _));
        var table = new byte[50];
        table[1] = (byte)version;
        table[3] = 2;
        table[5] = 2;
        table[6] = (byte)firstCharacterSet;
        table[7] = table[8] = table[9] = 1;
        table[10] = table[11] = table[12] = table[13] = 1;
        table[15] = 18;
        table[17] = 34;
        Array.Copy(source, offset + 12, table, 18, 16);
        Array.Copy(source, offset + 12, table, 34, 16);
        table[25] = 64;
        var metrics = OfficeOpenTypeLineMetrics.TryRead(ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics(table));
        Assert.True(metrics.HasValue);
        metrics!.Value.GetDeviceMetrics(8, .12D, out double ascent, out double descent);
        Assert.Equal(60D, ascent);
        Assert.Equal(15D, descent);
    }

    [Theory]
    [InlineData(5)]
    [InlineData(11)]
    [InlineData(15)]
    [InlineData(27)]
    public void TruncatedOptionalDeviceMetricsFallBackWithinTheFontTable(int length) {
        byte[] valid = ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics();
        var reader = OfficeOpenTypeReader.TryCreate(valid)!;
        Assert.True(reader.TryGetTable("VDMX", out int offset, out _));
        var truncated = new byte[length];
        Array.Copy(valid, offset, truncated, 0, length);
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics(truncated);
        var metrics = OfficeOpenTypeLineMetrics.TryRead(font);
        Assert.True(metrics.HasValue);
        metrics!.Value.GetDeviceMetrics(8, .12D, out double ascent, out double descent);
        Assert.Equal(54D, ascent);
        Assert.Equal(13D, descent);
    }
}
