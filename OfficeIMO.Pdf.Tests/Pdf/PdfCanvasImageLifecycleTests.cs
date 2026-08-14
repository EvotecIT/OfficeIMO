using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasImageLifecycleTests {
    private static readonly byte[] Png = PdfPngTestImages.CreateRgbPng(1, 1);

    [Fact]
    public void ClippedEffectImageTokenCannotBeReusedByALaterImage() {
        byte[] bytes = CreateDocument()
            .Canvas(canvas => canvas
                .Clip(0D, 0D, 10D, 10D, clipped => clipped.Effect(
                    OfficeTransform.Identity,
                    0.5D,
                    effect => effect.Image(Png, 40D, 40D, 12D, 12D)))
                .Image(Png, 20D, 20D, 12D, 12D))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Equal(1, CountOccurrences(raw, "/Im1 Do"));
        Assert.DoesNotContain("/Im2 Do", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void EffectTransformsAndOpacityContainTableCellImages() {
        byte[] bytes = CreateDocument()
            .Canvas(canvas => canvas.Effect(
                OfficeTransform.Translate(12D, 7D),
                0.5D,
                effect => effect.Table(CreateImageRows(), 20D, 20D, 80D, 50D, CreateTableStyle())))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int imageDraw = raw.IndexOf("/Im1 Do", StringComparison.Ordinal);
        int effectDraw = raw.IndexOf("/Fx1 Do", StringComparison.Ordinal);
        Assert.True(imageDraw >= 0, "Expected the table-cell image in the effect Form XObject.");
        Assert.True(effectDraw > imageDraw, "Expected the page to invoke the effect after its image-bearing Form XObject was serialized.");
        Assert.Contains("/GS1 gs", raw, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 12 -7 cm", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void FreeformCanvasClipContainsRotatedTableCellImages() {
        OfficeClipPath triangle = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(100D, 0D),
            OfficePathCommand.LineTo(50D, 80D),
            OfficePathCommand.Close());

        byte[] bytes = CreateDocument()
            .Canvas(canvas => canvas.Clip(20D, 20D, triangle, clipped =>
                clipped.Table(CreateImageRows(), 20D, 20D, 100D, 80D, CreateTableStyle(), rotationAngle: 45D)))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int clip = raw.IndexOf("20 140 m 120 140 l 70 60 l h W* n", StringComparison.Ordinal);
        int imageDraw = raw.IndexOf("/Im1 Do", StringComparison.Ordinal);
        Assert.True(clip >= 0, "Expected the freeform canvas clip path in the page content stream.");
        Assert.True(imageDraw > clip, "Expected the rotated table-cell image to render inside the freeform clip state.");
        Assert.Contains("0.707", raw.Substring(clip, imageDraw - clip), StringComparison.Ordinal);
        Assert.DoesNotContain("20 60 100 80 re W", raw, StringComparison.Ordinal);
    }

    private static PdfDocument CreateDocument() => PdfDocument.Create(new PdfOptions {
        PageWidth = 160D,
        PageHeight = 160D,
        MarginLeft = 0D,
        MarginRight = 0D,
        MarginTop = 0D,
        MarginBottom = 0D,
        CompressContentStreams = false
    });

    private static PdfTableCell[][] CreateImageRows() => new[] {
        new[] {
            PdfTableCell.WithImages(
                string.Empty,
                new[] { new PdfTableCellImage(Png, 20D, 20D) })
        }
    };

    private static PdfTableStyle CreateTableStyle() => new PdfTableStyle {
        RowMinHeights = new List<double?> { 80D },
        CellPaddingX = 6D,
        CellPaddingY = 6D
    };

    private static int CountOccurrences(string value, string pattern) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(pattern, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += pattern.Length;
        }

        return count;
    }
}
