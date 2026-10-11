using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTextWrapTests {
    [Fact]
    public void Native_exclusion_references_keep_text_outside_an_intersecting_picture() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherPage page = document.Pages.Single(item => item.TextFrames.Any(frame => frame.Id == 327));
        PublisherTextFrame frame = page.TextFrames.Single(item => item.Id == 327);
        Assert.Equal(new uint[] { 345, 346 }, frame.WrapObjectIds);
        OfficeDrawingImage picture = PublisherNativeTests.Elements(page.Drawing).OfType<OfficeDrawingImage>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-346") == true);
        OfficeImagePlacement bounds = picture.Projection.Placement;
        OfficeDrawingRichText[] regions = PublisherNativeTests.Elements(page.Drawing).OfType<OfficeDrawingRichText>()
            .Where(item => item.SourceElementIds?.Contains("publisher-object-327") == true).ToArray();
        Assert.NotEmpty(regions);
        Assert.True(frame.X < bounds.X + bounds.Width && frame.X + frame.Width > bounds.X);
        Assert.True(frame.Y < bounds.Y + bounds.Height && frame.Y + frame.Height > bounds.Y);
        Assert.All(regions, region => Assert.False(region.X < bounds.X + bounds.Width - .001 && region.X + region.Width > bounds.X + .001
            && region.Y < bounds.Y + bounds.Height - .001 && region.Y + region.Height > bounds.Y + .001));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_WRAP_APPROXIMATED"
            && item.Location == "Contents/object/327" && item.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Native_wrap_counts_are_validated_before_projection() {
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes => {
            PublisherInputContractTests.WriteUInt32(bytes, PublisherTextFlowTests.NativeFrameField(bytes, 326, 0x46), 1);
        }, "SampleNewsletter.pub");
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
        Assert.Contains("text-wrap count", error.Message);
    }
}
