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
    public void PdfTextExtractor_ExtractAllText_ReadsInlineNestedFormResourceDictionaries() {
        byte[] bytes = BuildPdfWithInlineNestedFormResources();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Inline form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsNestedFormXObjects() {
        byte[] bytes = BuildPdfWithNestedFormInvocations();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Nested form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_PreservesInlineFormOrdering() {
        byte[] bytes = BuildPdfWithInlineFormTextOrdering();

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Before middle after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_UsesInheritedResourcesForFormXObjects() {
        byte[] bytes = BuildPdfWithInheritedFormResources();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Form hello");
        Assert.Equal(110, span.X, 3);
        Assert.Equal(220, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_TracksRepeatedFormInvocations() {
        byte[] bytes = BuildPdfWithRepeatedFormInvocations();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var spans = doc.Pages[0].GetTextSpans().Where(s => s.Text == "Repeated form").OrderBy(s => s.X).ToList();
        Assert.Equal(2, spans.Count);
        Assert.Equal(10, spans[0].X, 3);
        Assert.Equal(110, spans[1].X, 3);
        Assert.Equal(20, spans[0].Y, 3);
        Assert.Equal(20, spans[1].Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_TracksNestedFormInvocations() {
        byte[] bytes = BuildPdfWithNestedFormInvocations();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Nested form");
        Assert.Equal(120, span.X, 3);
        Assert.Equal(232, span.Y, 3);
    }

    [Fact]
    public void PdfLogicalDocument_Load_ReadsImagesReferencedByFormXObjects() {
        byte[] bytes = BuildPdfWithFormXObjectImage();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes);

        PdfLogicalImage image = Assert.Single(logical.Images);
        PdfImagePlacement placement = Assert.Single(image.Placements);
        Assert.Equal("ImNested", image.ResourceName);
        Assert.Equal("ImNested", placement.ResourceName);
        Assert.Equal(7, image.SourceImage.ObjectNumber);
        Assert.Equal(7, placement.ObjectNumber);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_AppliesScaledFormTransformsInOrder() {
        byte[] bytes = BuildPdfWithScaledFormMatrix();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Scaled form");
        Assert.Equal(26, span.X, 3);
        Assert.Equal(42, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsInlineNestedFormResourceDictionaries() {
        byte[] bytes = BuildPdfWithInlineNestedFormResources();

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Inline form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
    }


    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFormResourcesWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: true, contentUsesEscapedName: false);

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Escaped form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFormInvocationsWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: false, contentUsesEscapedName: true);

        string text = PdfTextExtractor.ExtractAllText(bytes);

        Assert.Contains("Escaped form", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsFormResourcesWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: true, contentUsesEscapedName: false);

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Escaped form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsFormInvocationsWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: false, contentUsesEscapedName: true);

        var doc = PdfReadDocument.Load(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Escaped form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
    }

}
