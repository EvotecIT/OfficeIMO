using System.Buffers.Binary;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingFontStyleSelectionTests {
    [Theory]
    [InlineData(OfficeFontStyle.Regular, 400, 200, OfficeFontSlant.Normal)]
    [InlineData(OfficeFontStyle.Bold, 700, 800, OfficeFontSlant.Normal)]
    [InlineData(OfficeFontStyle.Italic, 400, 200, OfficeFontSlant.Italic)]
    [InlineData(OfficeFontStyle.Bold | OfficeFontStyle.Italic, 700, 800, OfficeFontSlant.Italic)]
    public void StyleResolutionAndMeasurementUseTheRequestedNumericFaceRegardlessOfRegistrationOrder(
        OfficeFontStyle style, int selectedWeight, int competingWeight, OfficeFontSlant slant) {
        byte[] selectedFont = FontWithAdvance(600);
        byte[] competingFont = FontWithAdvance(900);
        var selected = new OfficeFontFaceDescriptor(selectedWeight, 100D, slant);
        var competing = new OfficeFontFaceDescriptor(competingWeight, 100D, slant);
        foreach (bool selectedLast in new[] { false, true }) {
            var fonts = new OfficeFontFaceCollection();
            if (!selectedLast) fonts.Add("Scoped", selectedFont, selected);
            fonts.Add("Scoped", competingFont, competing);
            if (selectedLast) fonts.Add("Scoped", selectedFont, selected);

            Assert.True(fonts.TryResolveFaceForText("A", "Scoped", style, out OfficeFontFace? face));
            Assert.Equal(selected, face!.Descriptor);
            Assert.True(fonts.TryResolveFaceForText("A", "Scoped", style, 20D, out OfficeFontFace? sized));
            Assert.Equal(selected, sized!.Descriptor);
            Assert.True(fonts.TryMeasureText("AA", 20D, "Scoped", style, out double width));
            Assert.Equal(24D, width, 6);
        }
    }

    [Fact]
    public void StyleResolutionUsesTheSameNearestFaceAndFallbackCoverageAsNumericResolution() {
        byte[] nearest = FontWithAdvance(600);
        var fonts = new OfficeFontFaceCollection()
            .Add("Scoped", nearest, new OfficeFontFaceDescriptor(500, 100D))
            .Add("Scoped", FontWithAdvance(900), new OfficeFontFaceDescriptor(300, 100D))
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('B'));
        fonts.AddFallbackFamily("Fallback");

        Assert.True(fonts.TryResolveFaceForText("A", "Scoped", OfficeFontFaceDescriptor.Regular, out OfficeFontFace? numeric));
        Assert.Equal(500, numeric!.Descriptor.Weight);
        Assert.True(fonts.TryResolveFaceForText("A", "Scoped", OfficeFontStyle.Regular, out OfficeFontFace? styled));
        Assert.Same(numeric, styled);
        Assert.True(fonts.TryMeasureText("AA", 20D, "Scoped", OfficeFontStyle.Regular, out double width));
        Assert.Equal(24D, width, 6);
        Assert.True(fonts.TryResolveFaceForText("B", "Scoped, Fallback", OfficeFontStyle.Regular, out OfficeFontFace? fallback));
        Assert.Equal("Fallback", fallback!.FamilyName);
    }

    [Theory]
    [InlineData(400)]
    [InlineData(700)]
    public void RegisteredFallbackDoesNotSuppressOrMisattributeSubstitutionDiagnostics(int fallbackWeight) {
        var fonts = new OfficeFontFaceCollection()
            .Add("Requested", ManagedTextShapingTestAssets.CreateFont('A'))
            .Add("Fallback", ManagedTextShapingTestAssets.CreateFont('B'), new OfficeFontFaceDescriptor(fallbackWeight, 100D))
            .AddFallbackFamily("Fallback");

        OfficeImageExportDiagnostic? diagnostic = fonts.CreateSubstitutionDiagnostic("B", "Requested, Fallback");

        Assert.NotNull(diagnostic);
        Assert.Equal(OfficeImageExportDiagnosticCodes.FontSubstituted, diagnostic!.Code);
        Assert.Contains("caller-supplied 'Fallback'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, diagnostic.LossKind);
    }

    private static byte[] FontWithAdvance(ushort advance) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        int count = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4));
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (System.Text.Encoding.ASCII.GetString(font, record, 4) != "hmtx") continue;
            int offset = (int)BinaryPrimitives.ReadUInt32BigEndian(font.AsSpan(record + 8));
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(offset), advance);
            return font;
        }
        throw new InvalidOperationException("The test font must contain horizontal metrics.");
    }
}
