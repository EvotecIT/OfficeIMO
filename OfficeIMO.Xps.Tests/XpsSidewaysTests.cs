using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsSidewaysTests {
    [Theory]
    [InlineData("os2", XpsFormat.Xps)]
    [InlineData("hhea", XpsFormat.OpenXps)]
    [InlineData("os2-distinct", XpsFormat.OpenXps)]
    [InlineData("vertical", XpsFormat.Xps)]
    [InlineData("vertical", XpsFormat.OpenXps)]
    public void SidewaysOutlinesUseTopCenterOriginAndVerticalAdvance(string metrics, XpsFormat format) {
        var doc = XpsSidewaysFixtures.Create(metrics, format);
        var page = doc.Pages[0]; var xml = page.GetMarkup(); var run = xml.Elements().Last();
        run.SetAttributeValue("UnicodeString", "AA"); run.SetAttributeValue("Indices", "36;36"); page.ReplaceMarkup(xml);
        var sideways = Points(page);
        run.SetAttributeValue("Indices", "36"); run.SetAttributeValue("UnicodeString", "A"); run.SetAttributeValue("IsSideways", "false");
        run.SetAttributeValue("OriginX", "0"); run.SetAttributeValue("OriginY", "0"); page.ReplaceMarkup(xml);
        var horizontal = Points(page);
        Assert.Equal(horizontal.Length * 2, sideways.Length);
        var tables = XpsSidewaysFixtures.Tables();
        const int glyph = 36;
        double scale = 48D / XpsSidewaysFixtures.U16(tables["head"], 18);
        int hCount = XpsSidewaysFixtures.U16(tables["hhea"], 34);
        double topX = XpsSidewaysFixtures.U16(tables["hmtx"], Math.Min(glyph, hCount - 1) * 4) * scale / 2;
        double topY, advance;
        if (metrics == "vertical") {
            // Outline extrema are independent evidence of this simple glyph's declared yMax.
            topY = -horizontal.Where((_, i) => i % 2 == 1).Min() + (100 + glyph) * scale;
            advance = 2300 * scale;
        } else {
            byte[] table = tables[metrics.StartsWith("os2", StringComparison.Ordinal) ? "OS/2" : "hhea"];
            int position = metrics.StartsWith("os2", StringComparison.Ordinal) ? 68 : 4;
            topY = (metrics == "os2-distinct" ? 2100 : XpsSidewaysFixtures.I16(table, position)) * scale;
            advance = topY + (metrics == "os2-distinct" ? 450 : Math.Abs(XpsSidewaysFixtures.I16(table, position + 2))) * scale;
        }
        for (int i = 0; i < horizontal.Length; i += 2) {
            Assert.Equal(30 + horizontal[i + 1] + topY, sideways[i], 3);
            Assert.Equal(80 - horizontal[i] + topX, sideways[i + 1], 3);
            Assert.Equal(sideways[i] + advance, sideways[i + horizontal.Length], 3);
            Assert.Equal(sideways[i + 1], sideways[i + horizontal.Length + 1], 3);
        }
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void ExplicitClusterPlacementOverridesAdvanceAndRetainsOffsets(XpsFormat format) {
        var doc = XpsSidewaysFixtures.Create("vertical", format);
        var page = doc.Pages[0]; var xml = page.GetMarkup(); var run = xml.Elements().Last();
        const int glyph = 36;
        run.SetAttributeValue("UnicodeString", "fiA"); run.SetAttributeValue("Indices", $"(2:1){glyph},100;{glyph},100,25,10");
        page.ReplaceMarkup(xml); var points = Points(page); int count = points.Length / 2;
        for (int i = 0; i < count; i += 2) {
            Assert.Equal(points[i] + 60, points[i + count], 3);
            Assert.Equal(points[i + 1] - 4.8, points[i + count + 1], 3);
        }
        Assert.Equal("fiA", XpsDocument.Load(doc.Save()).Pages[0].ExtractText());
    }

    [Theory]
    [InlineData("vertical-truncated")]
    [InlineData("vertical-zero-count")]
    public void MalformedVerticalMetricsCannotReadAcrossFontTables(string metrics) {
        Assert.Throws<InvalidDataException>(() => XpsSidewaysFixtures.Create(metrics).Pages[0].ToSvg());
    }

    [Theory]
    [InlineData("true", "1")]
    [InlineData("1", "3")]
    [InlineData("yes", "0")]
    public void InvalidSidewaysRunsAreRejected(string sideways, string bidi) {
        var page = XpsSidewaysFixtures.Create("os2").Pages[0]; var xml = page.GetMarkup();
        xml.Elements().Last().SetAttributeValue("IsSideways", sideways); xml.Elements().Last().SetAttributeValue("BidiLevel", bidi);
        page.ReplaceMarkup(xml);
        Assert.Throws<InvalidDataException>(() => page.ToSvg(allowPartial: true));
    }

    [Theory]
    [InlineData("os2")]
    [InlineData("hhea")]
    [InlineData("vertical")]
    public void RotatedClippedSidewaysRunsSurviveNativeAndPdfConversion(string metrics) {
        var doc = XpsDocument.Load(XpsSidewaysFixtures.Create(metrics, positioned: true).Save());
        byte[] png = doc.Pages[0].ExportImage(OfficeImageExportFormat.Png).Bytes;
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out var raster));
        var pdf = OfficeIMO.Pdf.PdfDocument.Load(doc.ToPdf());
        var pdfRaster = OfficeDrawingRasterRenderer.Render(pdf.Render.Drawing(1), 96D / 72D);
        int dark = 0;
        for (int y = 0; y < 160; y++) for (int x = 0; x < 260; x++) {
            if (raster!.GetPixel(x, y).R < 80) dark++;
            Assert.InRange(Math.Abs(raster!.GetPixel(x, y).R - pdfRaster.GetPixel(x, y).R), 0, 8);
        }
        Assert.True(dark > 100);
    }

    private static double[] Points(XpsPage page) {
        var svg = XElement.Parse(page.ToSvg().Svg);
        string path = (string)svg.Descendants(XName.Get("path", "http://www.w3.org/2000/svg")).Last().Attribute("d")!;
        return Regex.Matches(path, @"-?\d+(?:\.\d+)?").Cast<Match>().Select(m => double.Parse(m.Value, CultureInfo.InvariantCulture)).ToArray();
    }
}
