using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfWatermarkWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InteractionIdentifiesThePaintedWatermarkWithoutLabelingOrdinaryContent(bool image) {
        var settings = Options(image);
        byte[] source = PdfDocument.Load(Source()).Stamp.Watermark(settings).ToBytes();
        var map = PdfPageInteractionMap.Create(source, 2);
        var watermark = map.Regions.Where(region => region.WatermarkId == settings.Id).ToArray();
        Assert.NotEmpty(watermark);
        Assert.All(watermark, region => Assert.Equal(image ? PdfInteractionKind.Image : PdfInteractionKind.Text, region.Kind));
        Assert.Contains(map.TextRegions, region => region.WatermarkId is null);
        Assert.All(PdfPageInteractionMap.Create(source, 1).Regions, region => Assert.Null(region.WatermarkId));
    }

    [Fact]
    public void RevisionRetainsResourcesWhenUnusedFormExceedsUsageAnalysisLimit() {
        var settings = Options(false);
        byte[] first = PdfDocument.Load(Source()).Stamp.Watermark(settings).ToBytes();
        var (objects, trailer) = PdfSyntax.ParseObjects(first, null);
        var read = PdfReadDocument.Open(first);
        var page = (PdfDictionary)objects[read.Pages[1].ObjectNumber].Value;
        var resources = (PdfDictionary)PdfObjectLookup.Resolve(objects, page.Items["Resources"])!;
        var xObjects = (PdfDictionary)PdfObjectLookup.Resolve(objects, resources.Items["XObject"])!;
        int number = objects.Keys.Max() + 1;
        var form = new PdfDictionary();
        form.Items["Type"] = new PdfName("XObject");
        form.Items["Subtype"] = new PdfName("Form");
        objects[number] = new PdfIndirectObject(number, 0, new PdfStream(form,
            System.Text.Encoding.ASCII.GetBytes(string.Concat(Enumerable.Repeat("q Q\n", 200)))));
        xObjects.Items["UnusedComplexForm"] = new PdfReference(number, 0);
        byte[] source = PdfPageExtractor.ExtractPages(objects, read.UncheckedMetadata, read.Pages.Select(item => item.ObjectNumber).ToArray(),
            catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailer));

        settings.Text = "UPDATED";
        byte[] revised = PdfDocument.Load(source).Stamp.Watermark(settings, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentOperations = 100 }
        }).ToBytes();
        Assert.Contains("UPDATED", PdfDocument.Load(revised).Read().Text);
        var (outputObjects, _) = PdfSyntax.ParseObjects(revised, null);
        var outputPage = (PdfDictionary)outputObjects[PdfReadDocument.Open(revised).Pages[1].ObjectNumber].Value;
        var outputResources = (PdfDictionary)PdfObjectLookup.Resolve(outputObjects, outputPage.Items["Resources"])!;
        var outputForms = (PdfDictionary)PdfObjectLookup.Resolve(outputObjects, outputResources.Items["XObject"])!;
        Assert.Contains("UnusedComplexForm", outputForms.Items.Keys);
        Assert.True(outputForms.Items.Count >= xObjects.Items.Count);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("Do")]
    [InlineData("gs")]
    public void RevisionRetainsResourcesUsedByOtherPageContent(string? splitOperator) {
        var settings = Options(false);
        settings.Text = "PRIOR";
        byte[] first = PdfDocument.Load(Source()).Stamp.Watermark(settings).ToBytes();
        var (objects, trailer) = PdfSyntax.ParseObjects(first, null);
        var read = PdfReadDocument.Open(first);
        var page = (PdfDictionary)objects[read.Pages[1].ObjectNumber].Value;
        var contents = (PdfArray)PdfObjectLookup.Resolve(objects, page.Items["Contents"])!;
        var watermark = Assert.Single(contents.Items.Select(item => PdfObjectLookup.Resolve(objects, item)).OfType<PdfStream>(),
            stream => stream.Dictionary.Get<PdfStringObj>("OfficeIMOWatermarkId")?.Value == settings.Id);
        int additionalNumber = objects.Keys.Max() + 1;
        objects[additionalNumber] = new PdfIndirectObject(additionalNumber, 0, new PdfStream(new PdfDictionary(), watermark.Data));
        contents.Items.Add(new PdfReference(additionalNumber, 0));
        byte[] shared = PdfPageExtractor.ExtractPages(objects, read.UncheckedMetadata, read.Pages.Select(item => item.ObjectNumber).ToArray(),
            catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailer));
        settings.Text = "UPDATED";
        byte[] expected = Render(PdfDocument.Load(shared).Stamp.Watermark(settings).ToBytes(), 2);
        if (splitOperator is not null) {
            string invocation = System.Text.Encoding.ASCII.GetString(watermark.Data);
            int split = invocation.IndexOf(" " + splitOperator, StringComparison.Ordinal) + 1;
            Assert.True(split > 0, "Fixture must contain the selected resource operator.");
            objects[additionalNumber] = new PdfIndirectObject(additionalNumber, 0, new PdfStream(new PdfDictionary(),
                System.Text.Encoding.ASCII.GetBytes(invocation.Substring(0, split))));
            objects[additionalNumber + 1] = new PdfIndirectObject(additionalNumber + 1, 0, new PdfStream(new PdfDictionary(),
                System.Text.Encoding.ASCII.GetBytes(invocation.Substring(split))));
            contents.Items.Add(new PdfReference(additionalNumber + 1, 0));
            shared = PdfPageExtractor.ExtractPages(objects, read.UncheckedMetadata, read.Pages.Select(item => item.ObjectNumber).ToArray(),
                catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, trailer));
        }
        var revised = PdfDocument.Load(shared).Stamp.Watermark(settings);
        Assert.Equal(expected, Render(revised.ToBytes(), 2));
        var spans = PdfReadDocument.Open(revised.ToBytes()).Pages[1].GetTextSpans();
        Assert.Single(spans, span => span.Text == "PRIOR");
        Assert.Single(spans, span => span.Text == "UPDATED");
        Assert.True(revised.Render.DisplayPage(2).Succeeded);
        var map = PdfPageInteractionMap.Create(revised.ToBytes(), 2);
        Assert.Equal("UPDATED", string.Concat(map.TextRegions.Where(region => region.WatermarkId == settings.Id).Select(region => region.Text)));
        Assert.Contains("PRIOR", string.Concat(map.TextRegions.Where(region => region.WatermarkId is null).Select(region => region.Text)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedRevisionDoesNotRetainSupersededWatermarkResources(bool image) {
        var settings = Options(image);
        var document = PdfDocument.Load(Source()).Stamp.Watermark(settings);
        int initialSize = document.ToBytes().Length;
        byte[] appearance = Render(document.ToBytes(), 2);
        for (int index = 0; index < 25; index++) document = document.Stamp.Watermark(settings);
        int revisedSize = document.ToBytes().Length;
        Assert.True(revisedSize < initialSize * 2, $"Repeated revisions grew from {initialSize} to {revisedSize} bytes.");
        Assert.Single(document.Stamp.ReadWatermarks());
        Assert.Equal(appearance, Render(document.ToBytes(), 2));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SavedWatermarkSettingsCanBeReadAndEditedWithoutOriginalOptions(bool image) {
        var original = Options(image);
        original.X = null;
        original.Y = 120;
        original.BehindContent = true;
        var first = PdfDocument.Load(Source()).Stamp.Watermark(original);
        var reopened = PdfDocument.Load(first.ToBytes());
        var saved = Assert.Single(reopened.Stamp.ReadWatermarks());
        Assert.Equal(original.Id, saved.Id);
        Assert.Equal(original.Text, saved.Text);
        Assert.Equal(original.ImageBytes, saved.ImageBytes);
        Assert.Equal(original.X, saved.X);
        Assert.Equal(original.Y, saved.Y);
        Assert.Equal(original.Width, saved.Width);
        Assert.Equal(original.Height, saved.Height);
        Assert.Equal(original.FontSize, saved.FontSize);
        Assert.Equal(original.Font, saved.Font);
        Assert.Equal(original.RotationDegrees, saved.RotationDegrees);
        Assert.Equal(original.Opacity, saved.Opacity);
        Assert.Equal(original.Color.ToOfficeColor(), saved.Color.ToOfficeColor());
        Assert.True(saved.BehindContent);
        Assert.Equal(new[] { 2 }, saved.TargetPages!.Resolve(2));
        saved.Width = 100;
        saved.Y = 145;
        var changed = reopened.Stamp.Watermark(saved);
        Assert.Single(changed.Stamp.ReadWatermarks());
        Assert.Equal(100, changed.Stamp.ReadWatermarks()[0].Width);
        Assert.False(Render(first.ToBytes(), 2).SequenceEqual(Render(changed.ToBytes(), 2)));
    }

    [Fact]
    public void TextWatermarkCanBeRevisedToAnImageAndMovedBehindContent() {
        var settings = Options(false);
        settings.Text = "REPLACE ME";
        var first = PdfDocument.Load(Source()).Stamp.Watermark(settings);
        settings.ImageBytes = PdfPngTestImages.CreateRgbPng(20, 200, 30);
        settings.BehindContent = true;
        var revised = first.Stamp.Watermark(settings);
        Assert.DoesNotContain("REPLACE ME", revised.Read().Text);
        Assert.Contains("Original second page", revised.Read().Text);
        Assert.NotEmpty(revised.Read().Pages[1].Images);
        Assert.False(Render(first.ToBytes(), 2).SequenceEqual(Render(revised.ToBytes(), 2)));
    }

    [Fact]
    public void SameIdentifierRevisesWatermarkAfterSaveAndReopen() {
        var settings = Options(false);
        settings.Text = "FIRST";
        var first = PdfDocument.Load(Source()).Stamp.Watermark(settings);
        settings.Text = "REVISED";
        settings.Color = PdfColor.FromRgb(20, 200, 40);
        var revised = PdfDocument.Load(first.ToBytes()).Stamp.Watermark(settings);
        string text = revised.Read().Text;
        Assert.DoesNotContain("FIRST", text);
        Assert.Contains("REVISED", text);
        Assert.Contains("Original second page", text);
        Assert.Single(PdfReadDocument.Open(revised.ToBytes()).Pages[1].GetTextSpans(), span => span.Text == "REVISED");
    }

    [Fact]
    public void RevisingPageSelectionRemovesOldTargetsAndPreservesOtherWatermarks() {
        var settings = Options(false);
        settings.Text = "MOVING";
        var first = PdfDocument.Load(Source()).Stamp.Watermark(settings);
        var other = Options(false);
        other.Text = "RETAINED";
        first = first.Stamp.Watermark(other);
        settings.TargetPages = PdfPageSelector.Parse("1");
        var revised = first.Stamp.Watermark(settings).Read();
        Assert.Contains("MOVING", Text(revised.Pages[0]));
        Assert.DoesNotContain("MOVING", Text(revised.Pages[1]));
        Assert.Contains("RETAINED", Text(revised.Pages[1]));
    }

    [Fact]
    public void IdenticalRevisionPreservesAppearanceAndStackingWithLaterContent() {
        var settings = Options(false);
        settings.Text = "LOWER";
        var first = PdfDocument.Load(Source()).Stamp.Watermark(settings);
        var upper = Options(false);
        upper.Text = "UPPER";
        upper.Opacity = 1;
        upper.Color = PdfColor.FromRgb(0, 0, 200);
        first = first.Stamp.Watermark(upper);
        var revised = first.Stamp.Watermark(settings);
        Assert.Equal(Render(first.ToBytes(), 2), Render(revised.ToBytes(), 2));
    }

    [Fact]
    public void TextWatermarkIsTransparentWithoutABorderAndUsesClockwiseRotation() {
        var source = PdfDocument.Load(Source());
        var options = Options(false);
        options.RotationDegrees = 35;
        var output = source.Stamp.Watermark(options);
        Assert.Equal(source.Read().Pages[1].VectorPrimitiveCount, output.Read().Pages[1].VectorPrimitiveCount);
        var text = Assert.Single(PdfReadDocument.Open(output.ToBytes()).Pages[1].GetTextSpans(), span => span.Text == "REVIEW");
        Assert.Equal(-35, text.RotationDegrees, 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectedPageWatermarkPreservesSourceAndChangesOnlyItsTarget(bool image) {
        byte[] source = Source();
        var options = Options(image);
        byte[] output = PdfDocument.Load(source).Stamp.Watermark(options).ToBytes();
        var result = PdfDocument.Load(output).Read();
        Assert.Equal(2, result.Pages.Count);
        Assert.Contains("Original first page", Text(result.Pages[0]));
        Assert.Contains("Original second page", Text(result.Pages[1]));
        Assert.Equal(Render(source, 1), Render(output, 1));
        Assert.False(Render(source, 2).SequenceEqual(Render(output, 2)));
        if (!image) Assert.Contains("REVIEW", Text(result.Pages[1]));
        Assert.DoesNotContain("REVIEW", Text(PdfDocument.Load(source).Read().Pages[1]));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ZeroOpacityRetainsTheOriginalAppearanceForTextAndImages(bool image) {
        byte[] source = Source();
        var options = Options(image);
        options.Opacity = 0D;
        byte[] output = PdfDocument.Load(source).Stamp.Watermark(options).ToBytes();
        Assert.Equal(Render(source, 2), Render(output, 2));
    }

    [Fact]
    public void RotationAndColorChangeTheRenderedWatermark() {
        byte[] source = Source();
        var options = Options(false);
        options.RotationDegrees = 0D;
        byte[] initial = PdfDocument.Load(source).Stamp.Watermark(options).ToBytes();
        options.RotationDegrees = 70D;
        options.Color = PdfColor.FromRgb(0, 0, 240);
        byte[] changed = PdfDocument.Load(source).Stamp.Watermark(options).ToBytes();
        Assert.False(Render(initial, 2).SequenceEqual(Render(changed, 2)));
    }

    private static PdfWatermarkOptions Options(bool image) => new() {
        Text = "REVIEW", ImageBytes = image ? PdfPngTestImages.CreateRgbPng(240, 20, 20) : null,
        X = 100, Y = 160, Width = 160, Height = 60, FontSize = 24, RotationDegrees = 35,
        Color = PdfColor.FromRgb(240, 20, 20), Opacity = 0.5, TargetPages = PdfPageSelector.Parse("2")
    };

    private static byte[] Source() => PdfDocument.Create(document => {
        document.Page(page => page.Size(400, 500).Content(content => content.Text("Original first page")));
        document.Page(page => page.Size(400, 500).Content(content => content.Text("Original second page")));
    }).ToBytes();

    private static byte[] Render(byte[] source, int page) {
        var result = PdfDocument.Load(source).Render.DisplayPage(page, new PdfPageDisplayOptions { Scale = 1D });
        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Diagnostics));
        return Assert.IsType<byte[]>(result.Bytes);
    }

    private static string Text(PdfLogicalPage page) => string.Join(" ", page.TextBlocks.Select(block => block.Text));
}
