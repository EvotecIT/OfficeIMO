using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTypographyEvidenceCorpusTests {
    [Fact]
    public void ManagedRasterAndSvgUseTheSharedFontProgramsAndLogicalText() {
        Assert.Equal(OfficeTextShapingBackend.Managed,
            ((IOfficeTextShapingProviderMetadata)OfficeManagedTextShapingProvider.Instance).Backend);

        foreach (TypographyEvidenceCase evidence in TypographyEvidenceCorpus.Cases) {
            byte[] fontData = LoadFontData(evidence);
            OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection()
                .Add(evidence.Family, fontData).Faces);
            Assert.True(face.Program.HasGlyphs(evidence.Text), evidence.Name);

            OfficeTextShapingResult? shaped = OfficeManagedTextShapingProvider.Instance.ShapeText(
                new OfficeTextShapingRequest(
                    evidence.Text,
                    evidence.Family,
                    face.Program.GetFontDataForShaping(),
                    face.Program.IsOpenTypeCff,
                    face.Program.UnitsPerEm,
                    evidence.Direction,
                    evidence.Language));
            Assert.True(evidence.ManagedShapingExpected == (shaped != null), evidence.Name);

            var drawing = new OfficeDrawing(560, evidence.Direction == OfficeTextDirection.TopToBottom ? 180 : 100)
                .AddFont(evidence.Family, fontData);
            if (evidence.Direction == OfficeTextDirection.TopToBottom) {
                drawing.AddVerticalText(evidence.Text, 240, 10, 80, 160, new OfficeFontInfo(evidence.Family, 28));
            } else {
                drawing.AddText(evidence.Text, 10, 10, 540, 80,
                    new OfficeFontInfo(evidence.Family, 28),
                    wrapText: false);
            }
            drawing.ApplyImageExportOptions(new OfficeImageExportOptions {
                TextShapingProvider = OfficeManagedTextShapingProvider.Instance,
                TextShapingLanguage = evidence.Language
            });

            var diagnostics = new List<OfficeImageExportDiagnostic>();
            drawing.AppendFontDiagnostics(diagnostics, evidence.Name);
            Assert.DoesNotContain(diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);

            string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
            string logicalText = string.Concat(XDocument.Parse(svg)
                .Descendants()
                .Where(element => element.Name.LocalName == "text")
                .Select(element => element.Value));
            Assert.Contains(evidence.Text, logicalText, StringComparison.Ordinal);

            var renderDiagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                TextShapingProvider = OfficeManagedTextShapingProvider.Instance,
                TextShapingLanguage = evidence.Language,
                DiagnosticSink = renderDiagnostics,
                DiagnosticSource = evidence.Name
            });
            Assert.Contains(raster.GetPixels(), pixel => pixel != 0);
            if (evidence.Direction == OfficeTextDirection.TopToBottom) {
                Assert.Contains(renderDiagnostics, diagnostic =>
                    diagnostic.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback &&
                    diagnostic.LossKind == OfficeConversionLossKind.Approximation);
                XElement verticalText = Assert.Single(XDocument.Parse(svg).Descendants(), element => element.Name.LocalName == "text");
                Assert.Equal("browser-native", verticalText.Attribute("data-officeimo-shaping-backend")?.Value);
            }
        }
    }

    [Fact]
    public void RenderingProfilesExposeManagedAndHostProviderIdentity() {
        Assert.Equal(OfficeTextShapingBackend.Managed, OfficeRenderingProfile.Managed.TextShapingBackend);
        Assert.Equal(
            OfficeTextShapingBackend.HostProvided,
            new OfficeRenderingProfile("host", textShapingProvider: new DecliningProvider()).TextShapingBackend);
    }

    private static string FontPath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "TestAssets", fileName);

    private static byte[] LoadFontData(TypographyEvidenceCase evidence) {
        if (!string.IsNullOrEmpty(evidence.FontFileName)) return File.ReadAllBytes(FontPath(evidence.FontFileName));
        return evidence.Name == "Hebrew"
            ? ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x05E9, 0x05DC, 0x05D5, 0x05DD, 0x05E2)
            : ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'C', 'a', 'f', 'e', 0x0301);
    }

    private sealed class DecliningProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => null;
    }
}
