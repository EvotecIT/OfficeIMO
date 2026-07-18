using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReaderAndFooterRegressionTests {

    [Fact]
    public void PdfSyntax_ParseObjects_ReadsBooleanAndNullObjects() {
        byte[] bytes = BuildPdfWithBooleanAndNullObjects();

        var (map, _) = PdfSyntax.ParseObjects(bytes);

        Assert.True(map[3].Value is PdfBoolean boolTrue && boolTrue.Value);
        Assert.True(map[4].Value is PdfBoolean boolFalse && !boolFalse.Value);
        Assert.Same(PdfNull.Instance, map[5].Value);
    }

    [Fact]
    public void PdfSyntax_ParseObjects_ReadsBooleanAndNullDictionaryValues() {
        byte[] bytes = BuildPdfWithBooleanAndNullObjects();

        var (map, _) = PdfSyntax.ParseObjects(bytes);

        var metadata = Assert.IsType<PdfDictionary>(map[6].Value);
        Assert.True(metadata.Get<PdfBoolean>("IsTagged")?.Value);
        Assert.False(metadata.Get<PdfBoolean>("NeedsRendering")?.Value ?? true);
        Assert.IsType<PdfNull>(metadata.Items["OptionalContent"]);

        var flags = Assert.IsType<PdfArray>(metadata.Items["Flags"]);
        Assert.True(Assert.IsType<PdfBoolean>(flags.Items[0]).Value);
        Assert.False(Assert.IsType<PdfBoolean>(flags.Items[1]).Value);
        Assert.IsType<PdfNull>(flags.Items[2]);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_InheritsMediaBoxFromPagesNode() {
        byte[] pdfBytes = BuildPdfWithInheritedMediaBox(500, 700);

        var doc = PdfReadDocument.Open(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(500, width);
        Assert.Equal(700, height);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_PrefersCropBoxOverMediaBox() {
        byte[] pdfBytes = BuildPdfWithMediaAndCropBoxes(500, 700, 300, 400);

        var doc = PdfReadDocument.Open(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(300, width);
        Assert.Equal(400, height);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_ReadsInheritedIndirectMediaBoxArrays() {
        byte[] pdfBytes = BuildPdfWithInheritedIndirectMediaBox(520, 710);

        var doc = PdfReadDocument.Open(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(520, width);
        Assert.Equal(710, height);
    }

    [Fact]
    public void PdfReadPage_GetPageSize_PrefersInheritedIndirectCropBoxArrays() {
        byte[] pdfBytes = BuildPdfWithInheritedIndirectCropBox(520, 710, 320, 410);

        var doc = PdfReadDocument.Open(pdfBytes);

        Assert.Single(doc.Pages);
        var (width, height) = doc.Pages[0].GetPageSize();
        Assert.Equal(320, width);
        Assert.Equal(410, height);
    }

}
