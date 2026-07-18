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

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Form hello");
        Assert.Equal(110, span.X, 3);
        Assert.Equal(220, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_TracksRepeatedFormInvocations() {
        byte[] bytes = BuildPdfWithRepeatedFormInvocations();

        var doc = PdfReadDocument.Open(bytes);

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

        var doc = PdfReadDocument.Open(bytes);

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
    public void PdfLogicalDocument_Load_DoesNotExposeImagesFromUnusedFormXObjects() {
        byte[] bytes = BuildPdfWithUnusedFormXObjectImage();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes);

        Assert.Empty(logical.Images);
    }

    [Fact]
    public void PdfLogicalDocument_Load_DoesNotExposeImagesFromUnusedPageResources() {
        byte[] bytes = BuildPdfWithUnusedPageImageResource();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes);

        Assert.Empty(logical.Images);
    }

    [Fact]
    public void PdfLogicalDocument_Load_DoesNotExposeUnusedImageResourceAliases() {
        byte[] bytes = BuildPdfWithImageResourceAlias();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes);

        PdfLogicalImage image = Assert.Single(logical.Images);
        Assert.Equal("ImUsed", image.ResourceName);
        Assert.True(image.HasPlacements);
        Assert.All(image.Placements, placement => Assert.Equal("ImUsed", placement.ResourceName));
    }

    [Fact]
    public void PdfLogicalDocument_Load_DeduplicatesImagesDiscoveredThroughRepeatedForms() {
        byte[] bytes = BuildPdfWithRepeatedFormImageResource();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes);

        PdfLogicalImage image = Assert.Single(logical.Images);
        Assert.Equal("ImShared", image.ResourceName);
        Assert.Equal(2, image.Placements.Count);
    }

    [Fact]
    public void PdfReadPage_GetImages_PreservesDistinctDirectImagesInFormResources() {
        PdfReadPage page = CreatePdfReadPageWithDistinctDirectFormImages();

        IReadOnlyList<PdfImagePlacement> placements = page.GetImagePlacements();
        IReadOnlyList<PdfExtractedImage> images = GetImagesWithPlacements(page, placements);

        Assert.Equal(2, placements.Count);
        Assert.Equal(2, images.Count);
        Assert.Equal(new[] { "aaa", "bbb" }, images.Select(DecodeSingleRgbPngPixel).OrderBy(value => value, StringComparer.Ordinal).ToArray());
    }

    [Fact]
    public void PdfReadPage_GetImages_DoesNotMatchSiblingImagesByNameOnly() {
        PdfReadPage page = CreatePdfReadPageWithDirectPageAndFormImageNameCollision();

        IReadOnlyList<PdfImagePlacement> placements = page.GetImagePlacements();
        IReadOnlyList<PdfExtractedImage> images = GetImagesWithPlacements(page, placements);

        PdfExtractedImage image = Assert.Single(images);
        Assert.Single(placements);
        Assert.Equal("for", DecodeSingleRgbPngPixel(image));
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_AppliesScaledFormTransformsInOrder() {
        byte[] bytes = BuildPdfWithScaledFormMatrix();

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Scaled form");
        Assert.Equal(26, span.X, 3);
        Assert.Equal(42, span.Y, 3);
    }

    private static IReadOnlyList<PdfExtractedImage> GetImagesWithPlacements(PdfReadPage page, IReadOnlyList<PdfImagePlacement> placements) {
        MethodInfo? method = typeof(PdfReadPage).GetMethod(
            "GetImages",
            BindingFlags.Instance | BindingFlags.NonPublic,
            binder: null,
            new[] { typeof(int), typeof(IReadOnlyList<PdfImagePlacement>) },
            modifiers: null);

        return Assert.IsAssignableFrom<IReadOnlyList<PdfExtractedImage>>(method!.Invoke(page, new object[] { 0, placements })!);
    }

    private static string DecodeSingleRgbPngPixel(PdfExtractedImage image) {
        Assert.True(image.IsImageFile);
        Assert.Equal("png", image.FileExtension);
        byte[] scanline = PdfPngTestImages.DecodePngIdat(image.Bytes);
        Assert.Equal(4, scanline.Length);
        Assert.Equal((byte)0, scanline[0]);
        return Encoding.ASCII.GetString(scanline, 1, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsInlineNestedFormResourceDictionaries() {
        byte[] bytes = BuildPdfWithInlineNestedFormResources();

        var doc = PdfReadDocument.Open(bytes);

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

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Escaped form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
    }

    [Fact]
    public void PdfReadPage_GetTextSpans_ReadsFormInvocationsWithEscapedNames() {
        byte[] bytes = BuildPdfWithFormResourceNameEscapes(dictionaryUsesEscapedName: false, contentUsesEscapedName: true);

        var doc = PdfReadDocument.Open(bytes);

        Assert.Single(doc.Pages);
        var span = Assert.Single(doc.Pages[0].GetTextSpans(), s => s.Text == "Escaped form");
        Assert.Equal(10, span.X, 3);
        Assert.Equal(20, span.Y, 3);
    }

}
