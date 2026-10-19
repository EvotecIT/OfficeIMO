using System;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed class OfficeDrawingPortableFontTests {
    private const string Family = "OfficeIMO Baseline Sans";

    [Theory]
    [InlineData("OfficeIMO Baseline Sans, sans-serif", true)]
    [InlineData("\"OfficeIMO Baseline Sans\", sans-serif", true)]
    [InlineData("sans-serif, OfficeIMO Baseline Sans", false)]
    public void PortableTextRequiresTheSelectedFamilyBeforeCssFallbacks(string familyList, bool accepted) {
        var drawing = new OfficeDrawing(120, 40);
        drawing.Fonts.Add(Family, File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
            "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        double advance = drawing.Fonts.Faces[0].Program.Measure("AV", 12);
        drawing.AddPositionedText("AV", 0, 0, advance, 20, new OfficeImageFrameTransform(0, 0, 0),
            new OfficeFontInfo(familyList, 12), OfficeColor.Black, textAdvanceWidth: advance);
        Assert.Equal(accepted, OfficeDrawingPortableExporter.TryExportJson(drawing,
            new OfficeDrawingPortableOptions(Family), out string? json, out var diagnostics));
        if (accepted) {
            Assert.NotNull(json);
            Assert.Empty(diagnostics);
        } else {
            Assert.Null(json);
            Assert.Equal("DRAWING_FONT_FAMILY_UNSUPPORTED", Assert.Single(diagnostics).Code);
        }
    }
}
