using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTextDefaultsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ControlledGraphicDefaultsAndNativeMaterializedBindingsSurviveBothContainers(bool native) {
        string file = native ? "libreoffice-graphic-text-defaults.fodg" : "libreoffice-graphic-text-defaults.source.fodg";
        var document = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", file));
        double[] sizes = native ? new double[] { 18, 22, 28, 18, 24, 22, 18, 18 } : new double[] { 18, 22, 28, 20, 24, 33, 14, 18 };
        string[] colors = native
            ? new[] { "#9B2020", "#804090", "#906010", "#9B2020", "#805020", "#704070", "#9B2020", "#9B2020" }
            : new[] { "#207A35", "#804090", "#906010", "#107090", "#805020", "#704070", "#204FA0", "#207A35" };
        string[] body = document.Pages[0].Shapes.Select(s => s.Text).ToArray();
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            var result = read.Pages[0].ToDrawing();
            var text = result.Value.Elements.OfType<OfficeDrawingRichText>().ToArray();
            Assert.Equal(8, text.Length);
            Assert.Equal(body, text.Select(t => t.PlainText));
            Assert.Equal(sizes, text.Select(t => Assert.Single(Assert.Single(t.Paragraphs).Runs).FontSize));
            Assert.Equal(colors, text.Select(t => Assert.Single(Assert.Single(t.Paragraphs).Runs).Color.ToString()));
            Assert.True(Assert.Single(text[6].Paragraphs).Runs.Single().Bold);
            Assert.True(Assert.Single(text[7].Paragraphs).Runs.Single().Bold);
            Assert.Equal(native, result.Report.Mappings.Any(m => m.Feature.EndsWith(":text-window-color", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported));
            Assert.Equal(before, Parts(read));
        }
    }

    private static string[] Parts(OdgDocument document) => new[] { "content.xml", "styles.xml" }.Select(p => document.GetXml(p).ToString()).ToArray();
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))), OdgDocument.LoadFlatXml(flat) };
    }
}
