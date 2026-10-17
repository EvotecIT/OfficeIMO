using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    private static void AssertRenderedWordBackgrounds(WordDocument document, params OfficeColor[] colors) {
        // Background paint is independent of the reader's unsupported embedded @font-face resources.
        var options = new WordImageExportOptions { BackgroundColor = OfficeColor.White };
        byte[] pngBytes = document.ToPng(options);
        OfficeRasterImage png = VisualBaselineTestSupport.DecodePng(pngBytes, "Word background proof is not a valid PNG.");
        string svg = document.ToSvg(options);
        bool decoded = OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported);
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_HTML_LAYOUT_EVIDENCE_DIR");
        if (!string.IsNullOrWhiteSpace(evidence)) {
            Directory.CreateDirectory(evidence);
            string name = "word-background-" + string.Join("-", colors.Select(color => color.ToRgbHex()));
            File.WriteAllText(Path.Combine(evidence, name + ".svg"), svg);
            File.WriteAllBytes(Path.Combine(evidence, name + ".png"), pngBytes);
            File.WriteAllText(Path.Combine(evidence, name + ".txt"),
                "SVG reader succeeded: " + decoded + "\nUnsupported feature count: " + unsupported
                + "\nEmbedded @font-face stylesheet: " + svg.Contains("@font-face", StringComparison.OrdinalIgnoreCase)
                + "\nKnown reader limit: embedded @font-face CSS resources are not imported."
                + "\nProof scope: decoded PNG and SVG background colors; font-resource ingestion and glyph fidelity are not qualified here.\n");
        }
        Assert.True(decoded, "Word background SVG could not be read.");
        OfficeRasterImage svgRaster = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);

        foreach (OfficeColor color in colors) {
            AssertPaintedBackground(png, color, "PNG");
            AssertPaintedBackground(svgRaster, color, "SVG");
        }
    }

    private static void AssertPaintedBackground(OfficeRasterImage image, OfficeColor color, string format) {
        int pixels = 0;
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            OfficeColor actual = image.GetPixel(x, y);
            if (actual.A == 255 && actual.R == color.R && actual.G == color.G && actual.B == color.B) pixels++;
        }
        Assert.True(pixels >= 40, format + " did not paint the Word run background #" + color.ToRgbHex() + ".");
    }
}
