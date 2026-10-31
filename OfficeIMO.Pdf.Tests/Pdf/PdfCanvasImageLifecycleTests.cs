using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasImageLifecycleTests {
    private static readonly byte[] Png = PdfPngTestImages.CreateRgbPng(1, 1);

    [Fact]
    public void ClippedEffectImageTokenCannotBeReusedByALaterImage() {
        byte[] bytes = CreateDocument()
            .Canvas(canvas => canvas
                .Clip(global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(10D), global::OfficeIMO.OfficeLength.Points(10D), clipped => clipped.Effect(
                    OfficeTransform.Identity,
                    0.5D,
                    effect => effect.Image(Png, global::OfficeIMO.OfficeLength.Points(40D), global::OfficeIMO.OfficeLength.Points(40D), global::OfficeIMO.OfficeLength.Points(12D), global::OfficeIMO.OfficeLength.Points(12D))))
                .Image(Png, global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(12D), global::OfficeIMO.OfficeLength.Points(12D)))
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
                effect => effect.Table(CreateImageRows(), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(80D), global::OfficeIMO.OfficeLength.Points(50D), CreateTableStyle())))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        var objects = PdfSyntax.ParseObjects(bytes).Map;
        PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
            dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
        PdfDictionary pageResources = Assert.IsType<PdfDictionary>(PdfObjectLookup.ResolveChain(objects, page.Items["Resources"]));
        PdfDictionary pageXObjects = Assert.IsType<PdfDictionary>(PdfObjectLookup.ResolveChain(objects, pageResources.Items["XObject"]));
        PdfStream effect = Assert.IsType<PdfStream>(PdfObjectLookup.ResolveChain(objects, pageXObjects.Items["Fx1"]));
        Assert.Equal("Form", effect.Dictionary.Get<PdfName>("Subtype")?.Name);
        string effectOperators = Encoding.ASCII.GetString(StreamDecoder.DecodeRequired(effect, objects));
        Assert.Contains("/Im1 Do", effectOperators, StringComparison.Ordinal);
        PdfDictionary effectResources = Assert.IsType<PdfDictionary>(PdfObjectLookup.ResolveChain(objects, effect.Dictionary.Items["Resources"]));
        PdfDictionary effectXObjects = Assert.IsType<PdfDictionary>(PdfObjectLookup.ResolveChain(objects, effectResources.Items["XObject"]));
        PdfStream image = Assert.IsType<PdfStream>(PdfObjectLookup.ResolveChain(objects, effectXObjects.Items["Im1"]));
        Assert.Equal("Image", image.Dictionary.Get<PdfName>("Subtype")?.Name);
        Assert.Contains("/Fx1 Do", raw, StringComparison.Ordinal);
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
            .Canvas(canvas => canvas.Clip(global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(20D), triangle, clipped =>
                clipped.Table(CreateImageRows(), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(100D), global::OfficeIMO.OfficeLength.Points(80D), CreateTableStyle(), rotationAngle: 45D)))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int clip = raw.IndexOf("20 140 m 120 140 l 70 60 l h W* n", StringComparison.Ordinal);
        int imageDraw = raw.IndexOf("/Im1 Do", StringComparison.Ordinal);
        Assert.True(clip >= 0, "Expected the freeform canvas clip path in the page content stream.");
        Assert.True(imageDraw > clip, "Expected the rotated table-cell image to render inside the freeform clip state.");
        Assert.Contains("0.707", raw.Substring(clip, imageDraw - clip), StringComparison.Ordinal);
        Assert.DoesNotContain("20 60 100 80 re W", raw, StringComparison.Ordinal);
    }

    private static PdfDocument CreateDocument() => PdfDocument.Create(new PdfCreateOptions {
        PageWidth = global::OfficeIMO.OfficeLength.Points(160D),
        PageHeight = global::OfficeIMO.OfficeLength.Points(160D),
        MarginLeft = global::OfficeIMO.OfficeLength.Points(0D),
        MarginRight = global::OfficeIMO.OfficeLength.Points(0D),
        MarginTop = global::OfficeIMO.OfficeLength.Points(0D),
        MarginBottom = global::OfficeIMO.OfficeLength.Points(0D),
        CompressContentStreams = false
    });

    private static PdfTableCell[][] CreateImageRows() => new[] {
        new[] {
            PdfTableCell.WithImages(
                string.Empty,
                new[] { new PdfTableCellImage(Png, global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(20D)) })
        }
    };

    private static PdfTableStyle CreateTableStyle() => new PdfTableStyle {
        RowMinHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(80D) },
        CellPaddingX = global::OfficeIMO.OfficeLength.Points(6D),
        CellPaddingY = global::OfficeIMO.OfficeLength.Points(6D)
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
