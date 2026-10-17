using System;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingSvgFontImportTests {
    [Theory]
    [InlineData("src:url(https://fonts.example/font.ttf)")]
    [InlineData("src:url(file:///tmp/font.ttf)")]
    [InlineData("src:url(../font.ttf)")]
    [InlineData("src:local('Arial')")]
    [InlineData("src:url(data:text/html;base64,AAECAw==)")]
    [InlineData("src:url(data:font/ttf;base64,not-base64)")]
    [InlineData("src:url(data:font/ttf;base64,AAECAw==)")]
    [InlineData("src:url(data:font/woff;base64,{FONT})")]
    [InlineData("src:url(data:font/ttf;base64,{FONT}) format('unsupported')")]
    [InlineData("src:url(data:font/ttf;base64,{FONT}) tech(variations)")]
    [InlineData("src:url(data:font/ttf;base64,{FONT});font-weight:inherit")]
    [InlineData("src:url(data:font/ttf;base64,{FONT});font-stretch:NaN%")]
    [InlineData("src:url(data:font/ttf;base64,{FONT});font-style:oblique 90deg")]
    [InlineData("src:url(data:font/ttf;base64,{FONT});unicode-range:U+110000")]
    [InlineData("src:url(data:font/ttf;base64,{FONT});font-weight:500!important")]
    [InlineData("src:url(data:font/ttf;base64,{FONT});broken-descriptor")]
    public void EmbeddedSvgFontRejectedSourcesAndDescriptorsRemainDiagnostic(string descriptors) {
        string font = Convert.ToBase64String(ManagedTextShapingTestAssets.CreateFont('A'));
        byte[] bytes = SvgFontImportDocument("@font-face{font-family:FixtureFont;" + descriptors.Replace("{FONT}", font) + "}"
            + ".paint{fill:red}", "<rect class='paint' width='10' height='10'/>");

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Empty(drawing!.Fonts.Faces);
        Assert.Equal(OfficeColor.Red, Assert.Single(drawing.Shapes).Shape.FillColor);
        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(bytes));
    }

    [Fact]
    public void EmbeddedSvgFontFallbackDoesNotHideRejectedRemoteSources() {
        string source = "@font-face{font-family:FixtureFont;src:url('https://fonts.example/font.ttf'),"
            + "url('data:font/ttf;base64," + Convert.ToBase64String(ManagedTextShapingTestAssets.CreateFont('A')) + "') format('truetype');}";
        byte[] bytes = SvgFontImportDocument(source, "<text x='8' y='24'>A</text>");

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Equal("FixtureFont", Assert.Single(drawing!.Fonts.Faces).FamilyName);
        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(bytes));
    }

    [Theory]
    [InlineData("https://fonts.example/font.ttf")]
    [InlineData("data:text/html;base64,AAECAw==")]
    [InlineData("data:font/ttf;base64,not-base64")]
    public void EmbeddedSvgFontDoesNotHideRejectedSourcesAfterAUsableSource(string rejectedSource) {
        string source = "@font-face{font-family:FixtureFont;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(ManagedTextShapingTestAssets.CreateFont('A')) + "'),url('" + rejectedSource + "');}";
        byte[] bytes = SvgFontImportDocument(source);

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Single(drawing!.Fonts.Faces);
        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(bytes));
    }

    [Fact]
    public void EmbeddedSvgFontsShareOneRegistrationBudgetAcrossStylesheets() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        string first = string.Concat(Enumerable.Range(0, 128).Select(index => FontFace("Face" + index, font)));
        byte[] bytes = SvgFontImportDocument(first, "<style>" + FontFace("OverBudget", font) + "</style>");

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Equal(128, drawing!.Fonts.Faces.Count);
        Assert.DoesNotContain(drawing.Fonts.Faces, face => face.FamilyName == "OverBudget");
        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(bytes));
    }

    [Fact]
    public void EmbeddedSvgFontContainerCannotExpandBeyondTheDecodedByteLimit() {
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(ManagedTextShapingTestAssets.CreateFont('A'));
        // The WOFF header declares the decoded sfnt size independently of its small compressed input.
        uint declared = 8U * 1024U * 1024U + 1U;
        for (int index = 0; index < 4; index++) woff[16 + index] = (byte)(declared >> (24 - index * 8));
        string css = "@font-face{font-family:FixtureFont;src:url(data:font/woff;base64," + Convert.ToBase64String(woff) + ") format('woff')}";
        byte[] bytes = SvgFontImportDocument(css);

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported > 0);
        Assert.Empty(drawing!.Fonts.Faces);
        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(bytes));
    }

    [Fact]
    public void EmbeddedSvgFontImportAndContentSafetyHonorCancellationWithoutMutatingCallerFonts() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] bytes = SvgFontImportDocument(FontFace("FixtureFont", font));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var options = new OfficeSvgDrawingReaderOptions { CancellationToken = cancellation.Token };
        options.Fonts.Add("CallerFont", font);

        Assert.Throws<OperationCanceledException>(() => OfficeSvgDrawingReader.TryRead(bytes, options, out _));
        Assert.Throws<OperationCanceledException>(() => OfficeSvgDrawingReader.InspectContentSafety(bytes, readerOptions: options));
        Assert.Equal("CallerFont", Assert.Single(options.Fonts.Faces).FamilyName);
    }
}
