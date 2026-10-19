#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using ChartForgeX.Core;
using ChartForgeX.Primitives;
using ChartForgeX.Rendering;
using ChartForgeX.Themes;
using ChartForgeX.Typography;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

/// <summary>Exercises producer-to-wire preservation, strict failure and a real labeled chart.</summary>
public sealed class OfficeDrawingPortableTests {
    private const string Family = "OfficeIMO Baseline Sans";

    [Fact]
    public void ScaledClipRetainsPainterOrderNegativeGeometryAndPositionedText() {
        var child = new OfficeDrawing(100, 80);
        child.Fonts.AddRange(Fonts());
        var rectangle = OfficeShape.Rectangle(30, 20); rectangle.FillColor = OfficeColor.Black;
        child.AddShape(rectangle, 0, 0);
        double advance = child.Fonts.Faces[0].Program.Measure("ABC", 12);
        child.AddPositionedText("ABC", 10, 20, advance, 20, new OfficeImageFrameTransform(0, 0, 0),
            new OfficeFontInfo(Family, 12), OfficeColor.Black, textAdvanceWidth: advance);
        var container = new OfficeDrawing(120, 100);
        container.AddClippedDrawing(child, 7, 9, OfficeClipPath.Rectangle(90, 70), -10, -7);
        var root = new OfficeDrawing(240, 200);
        root.AddEffectDrawing(container, OfficeTransform.Scale(2, 2));
        Assert.True(OfficeDrawingPortableExporter.TryExportJson(root, new OfficeDrawingPortableOptions(Family), out string? json, out var diagnostics));
        Assert.Empty(diagnostics);
        using var document = JsonDocument.Parse(json!);
        JsonElement group = document.RootElement.GetProperty("elements")[0];
        Assert.Equal(14, group.GetProperty("clip").GetProperty("x").GetDouble());
        Assert.Equal(18, group.GetProperty("clip").GetProperty("y").GetDouble());
        Assert.Equal(180, group.GetProperty("clip").GetProperty("width").GetDouble());
        JsonElement shape = group.GetProperty("elements")[0];
        Assert.Equal("path", shape.GetProperty("kind").GetString());
        Assert.Equal(-6, shape.GetProperty("commands")[0].GetProperty("x").GetDouble());
        JsonElement textClip = group.GetProperty("elements")[1];
        JsonElement text = Texts(textClip).Single();
        Assert.Equal("ABC", text.GetProperty("text").GetString());
        Assert.Equal(14, text.GetProperty("x").GetDouble());
        Assert.Equal(68, text.GetProperty("y").GetDouble());
        Assert.Equal(24, text.GetProperty("fontSize").GetDouble());
    }

    [Fact]
    public void UnsupportedPaintAndResourceLimitsNeverReturnPartialJson() {
        var drawing = new OfficeDrawing(100, 100);
        var shape = OfficeShape.Rectangle(30, 20); shape.FillColor = OfficeColor.Black;
        drawing.AddShape(shape, 0, 0);
        var translucent = OfficeShape.Rectangle(30, 20); translucent.FillColor = OfficeColor.FromRgba(12, 30, 50, 128);
        drawing.AddShape(translucent, 30, 30);
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family), out var rejected, out var diagnostics));
        Assert.Null(rejected);
        Assert.Equal("DRAWING_ALPHA_UNSUPPORTED", Assert.Single(diagnostics).Code);
        Assert.Equal("elements[1]", diagnostics[0].Location);
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family) { MaximumElements = 1 }, out rejected, out diagnostics));
        Assert.Null(rejected); Assert.Equal("DRAWING_ELEMENT_LIMIT", Assert.Single(diagnostics).Code);
        var valid = new OfficeDrawing(100, 100); valid.AddShape(shape, 0, 0);
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(valid, new OfficeDrawingPortableOptions(Family) { MaximumCommands = 2 }, out rejected, out diagnostics));
        Assert.Null(rejected); Assert.Equal("DRAWING_COMMAND_LIMIT", Assert.Single(diagnostics).Code);
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(valid, new OfficeDrawingPortableOptions(Family) { MaximumJsonCharacters = 30 }, out rejected, out diagnostics));
        Assert.Null(rejected); Assert.Equal("DRAWING_JSON_LIMIT", Assert.Single(diagnostics).Code);
    }

    [Fact]
    public void PositionedPairKerningAndRightAlignmentDoNotBecomeScalarAdvanceSpacing() {
        var drawing = new OfficeDrawing(240, 90); drawing.Fonts.AddRange(Fonts());
        IOfficeFontProgram face = drawing.Fonts.Faces[0].Program;
        const double size = 18;
        double natural = face.Measure("AV", size);
        face.TryGetGlyphMetrics('A', out _, out int firstAdvance);
        face.TryGetGlyphMetrics('V', out _, out int secondAdvance);
        Assert.True(natural < (firstAdvance + secondAdvance) * size / face.UnitsPerEm);
        drawing.AddPositionedText("AV", 20, 15, 180, 25, new OfficeImageFrameTransform(0, 0, 0),
            new OfficeFontInfo(Family, size), alignment: OfficeTextAlignment.Right, textAdvanceWidth: natural);
        Assert.True(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family), out var json, out var diagnostics));
        Assert.Empty(diagnostics);
        using var document = JsonDocument.Parse(json!);
        JsonElement[] letters = Texts(document.RootElement).ToArray();
        Assert.Equal(new[] { "A", "V" }, letters.Select(item => item.GetProperty("text").GetString()).ToArray());
        double left = 200 - natural;
        Assert.Equal(left, letters[0].GetProperty("x").GetDouble(), 6);
        Assert.Equal(left + natural - secondAdvance * size / face.UnitsPerEm, letters[1].GetProperty("x").GetDouble(), 6);
        Assert.All(letters, item => Assert.Equal(33, item.GetProperty("y").GetDouble()));
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family) { MaximumElements = 2 }, out json, out diagnostics));
        Assert.Null(json); Assert.Equal("DRAWING_ELEMENT_LIMIT", Assert.Single(diagnostics).Code);
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions(Family) { MaximumDepth = 1 }, out json, out diagnostics));
        Assert.Null(json); Assert.Equal("DRAWING_DEPTH_LIMIT", Assert.Single(diagnostics).Code);
    }

    [Fact]
    public void MissingFontsAndUnsupportedTextLayoutCannotDisappearFromTheDrawing() {
        var missing = new OfficeDrawing(100, 60);
        missing.AddPositionedText("Test", 5, 5, 80, 30, new OfficeImageFrameTransform(0, 0, 0), new OfficeFontInfo(Family, 12));
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(missing, new OfficeDrawingPortableOptions(Family), out var json, out var diagnostics));
        Assert.Null(json); Assert.Equal("DRAWING_FONT_UNRESOLVED", Assert.Single(diagnostics).Code);
        var wrapped = new OfficeDrawing(100, 60); wrapped.Fonts.AddRange(Fonts());
        wrapped.AddText("A wrapped label", 5, 5, 80, 30, new OfficeFontInfo(Family, 12), wrapText: true);
        Assert.False(OfficeDrawingPortableExporter.TryExportJson(wrapped, new OfficeDrawingPortableOptions(Family), out json, out diagnostics));
        Assert.Null(json); Assert.Equal("DRAWING_TEXT_LAYOUT_UNSUPPORTED", Assert.Single(diagnostics).Code);
    }

    [Theory]
    [InlineData("font-variation-settings:'wght' 600")]
    [InlineData("font-feature-settings:'liga' 0")]
    [InlineData("font-language-override:'TRK'")]
    [InlineData("font-palette:--alternate")]
    public void StrictBridgeRejectsAdvancedFontSelections(string selection) {
        var source = new OfficeVisualSource("<svg xmlns='http://www.w3.org/2000/svg' width='120' height='60'><text x='5' y='30' style=\"" +
            selection + "\">Service</text></svg>");
        Assert.Throws<NotSupportedException>(() => source.ToOfficeVisual(new OfficeVisualConversionOptions { SvgPolicy = OfficeVisualSvgPolicy.RequireVector }));
    }

    [Theory]
    [InlineData("<style>text{font-feature-settings:'liga' 0}</style><text x='5' y='30'>Service</text>")]
    [InlineData("<text x='5' y='30' font-palette='--alternate'>Service</text>")]
    public void StrictBridgeAlsoRejectsFontSettingsFromStylesheetsAndPresentationAttributes(string content) {
        var source = new OfficeVisualSource("<svg xmlns='http://www.w3.org/2000/svg' width='120' height='60'>" + content + "</svg>");
        Assert.Throws<NotSupportedException>(() => source.ToOfficeVisual(new OfficeVisualConversionOptions { SvgPolicy = OfficeVisualSvgPolicy.RequireVector }));
    }

    [Fact]
    public void RegisteredFacesAndLatinPairPositionsSurviveActualServiceCharts() {
        OfficeFontFaceCollection fonts = Fonts();
        foreach (OfficeFontFace face in fonts.Faces)
            FontRegistry.Register(Family, face.Data, face.Descriptor.Weight, face.Descriptor.Slant == OfficeFontSlant.Italic);
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_PORTABLE_DRAWING_EVIDENCE");
        if (evidence != null) Directory.CreateDirectory(evidence);
        foreach (VisualThemeMode mode in new[] { VisualThemeMode.Light, VisualThemeMode.Dark }) {
            var chart = Chart.Create()
                .AddLine("Observed", new[] { new ChartPoint(1,42), new ChartPoint(2,58), new ChartPoint(3,73), new ChartPoint(4,91), new ChartPoint(5,67), new ChartPoint(6,82) })
                .AddLine("Capacity", new[] { new ChartPoint(1,100), new ChartPoint(2,100), new ChartPoint(3,100), new ChartPoint(4,100), new ChartPoint(5,100), new ChartPoint(6,100) });
            var prepared = chart.Prepare(new OfficeVisualDocumentStyle(Family).CreateContext(themeMode: mode,
                frame: new VisualFrame("Service capacity", "Events per minute · six observation intervals", showLegend: true,
                    legendPosition: ChartLegendPosition.BottomLeft, showSurface: true)));
            var conversionOptions = new OfficeVisualConversionOptions { SvgPolicy = OfficeVisualSvgPolicy.RequireVector };
            conversionOptions.Fonts.AddRange(fonts);
            var conversion = prepared.ToArtifact("service-capacity").ToOfficeVisual(conversionOptions);
            conversion.Report.RequireNoLoss();
            Assert.True(conversion.Report.IsVector);
            Assert.False(conversion.Report.UsedRasterFallback);
            Assert.True(OfficeDrawingPortableExporter.TryExportJson(conversion.Drawing, new OfficeDrawingPortableOptions(Family),
                out string? json, out var diagnostics), string.Join("; ", diagnostics.Select(item => item.Code + " " + item.Location + " " + item.Message)));
            using var document = JsonDocument.Parse(json!);
            Assert.Equal(450, document.RootElement.GetProperty("width").GetDouble());
            Assert.Equal(270, document.RootElement.GetProperty("height").GetDouble());
            string labels = string.Concat(Texts(document.RootElement).Select(item => item.GetProperty("text").GetString()));
            Assert.Contains("Service capacity", labels);
            Assert.Contains("Observed", labels);
            Assert.Contains("Capacity", labels);
            Assert.Contains("Events per minute", labels);
            Assert.Contains(Texts(document.RootElement), item => item.GetProperty("font").GetString() == "bold");
            if (evidence != null) {
                string name = "service-capacity-" + mode.ToString().ToLowerInvariant();
                File.WriteAllText(Path.Combine(evidence, name + ".json"), json);
                File.WriteAllText(Path.Combine(evidence, name + ".svg"), prepared.ToSvg());
                File.WriteAllBytes(Path.Combine(evidence, name + ".png"), prepared.ToPng());
            }
        }
        if (evidence != null) {
            var manifest = fonts.Faces.Select(face => new { family = face.FamilyName, slot = Slot(face.Style),
                path = Path.Combine(AppContext.BaseDirectory, "Fonts", "OfficeIMOBaselineSans-" + FaceFile(face.Style) + ".ttf"),
                source = "OfficeIMO.TestAssets/Fonts/OfficeIMOBaselineSans-" + FaceFile(face.Style) + ".ttf",
                sha256 = Convert.ToHexString(SHA256.HashData(face.Data)).ToLowerInvariant(), license = "SIL OFL 1.1", notice = "OfficeIMO.TestAssets/Fonts/OFL-Carlito.txt" });
            File.WriteAllText(Path.Combine(evidence, "font-manifest.json"), JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
        }
    }

    private static IEnumerable<JsonElement> Texts(JsonElement element) {
        if (element.TryGetProperty("kind", out JsonElement kind) && kind.GetString() == "text") yield return element;
        if (element.TryGetProperty("elements", out JsonElement children))
            foreach (JsonElement child in children.EnumerateArray()) foreach (JsonElement text in Texts(child)) yield return text;
    }
    private static OfficeFontFaceCollection Fonts() {
        var fonts = new OfficeFontFaceCollection();
        foreach (OfficeFontStyle style in new[] { OfficeFontStyle.Regular, OfficeFontStyle.Bold, OfficeFontStyle.Italic, OfficeFontStyle.Bold | OfficeFontStyle.Italic })
            fonts.Add(Family, File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "OfficeIMOBaselineSans-" + FaceFile(style) + ".ttf")), style);
        return fonts;
    }
    private static string FaceFile(OfficeFontStyle style) => style == OfficeFontStyle.Regular ? "Regular" : style == OfficeFontStyle.Bold ? "Bold" :
        style == OfficeFontStyle.Italic ? "Italic" : "BoldItalic";
    private static string Slot(OfficeFontStyle style) => style == OfficeFontStyle.Regular ? "regular" : style == OfficeFontStyle.Bold ? "bold" :
        style == OfficeFontStyle.Italic ? "italic" : "boldItalic";
}
#endif
