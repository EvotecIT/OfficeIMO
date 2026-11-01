using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTextWrapTests {
    [Fact]
    public void NonTop_wrapped_frames_keep_the_existing_vertical_alignment_path() {
        var document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/WrapObjectOrder/native-wrapobjectorder.pub"));
        PublisherTextFrame frame = document.Pages[0].TextFrames.Single(f => f.Id == 293);
        document.ReplaceStoryParagraphs(frame.StoryId, PublisherAuthoringTests.Paragraphs("Alignment marker"));
        var baselines = new List<double>();
        foreach (var alignment in new[] { OfficeTextVerticalAlignment.Top, OfficeTextVerticalAlignment.Center, OfficeTextVerticalAlignment.Bottom }) {
            document.SetTextFrameLayout(frame.Id, 1, 0, frame.Padding, alignment);
            var pdf = PdfReadDocument.Open(document.ToPdfBytes());
            var span = Assert.Single(pdf.Pages[0].GetTextSpans(), s => s.IsVisible && s.Text == "Alignment marker");
            baselines.Add(span.Y);
        }
        Assert.True(baselines[0] > baselines[1]); Assert.True(baselines[1] > baselines[2]);
        Assert.Equal(baselines[0] - baselines[1], baselines[1] - baselines[2], 4);
    }

    [Theory]
    [InlineData("WrapObjectOrder", true)]
    [InlineData("WrapObjectOrderDefaults", true)]
    [InlineData("WrapObjectOrderLeft", false)]
    [InlineData("WrapObjectOrderReversed", false)]
    public void Qualified_both_side_wrap_places_successive_fragments_on_a_shared_row(string control, bool bothSides) {
        var document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + control + "/native-" + control.ToLowerInvariant() + ".pub"));
        PublisherTextFrame frame = document.Pages[0].TextFrames.Single(f => f.Id == 293);
        var regions = PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingRichText>()
            .Where(t => t.SourceElementIds?.Contains("publisher-object-293") == true).ToArray();
        var splitRows = regions.GroupBy(t => Math.Round(t.Y, 6)).Where(row => row.Count() > 1).ToArray();
        Assert.Equal(bothSides, splitRows.Length != 0);
        if (bothSides) {
            Assert.All(splitRows, row => Assert.True(row.First().X < row.Last().X));
            double right = control == "WrapObjectOrderDefaults" ? 254.88 : 258;
            Assert.Contains(splitRows, row => Math.Abs(row.First().X - 72) < .001 && Math.Abs(row.Last().X - right) < .001);
        }
        if (control == "WrapObjectOrder") {
            Assert.All(regions, region => Assert.False(region.X < 258 - .001 && region.X + region.Width > 174 + .001
                && region.Y < 210 - .001 && region.Y + region.Height > 126 + .001));
        }
        Assert.Equal(0, frame.TextStart);
        int storyLength = document.TextStories.Single(story => story.Id == frame.StoryId).Text.Length;
        Assert.InRange(frame.TextLength!.Value, 1, storyLength);
        Assert.Equal(frame.TextLength < storyLength, frame.HasOverflow);
    }

    [Fact]
    public void Native_exclusion_references_keep_text_outside_an_intersecting_picture() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherPage page = document.Pages.Single(item => item.TextFrames.Any(frame => frame.Id == 327));
        PublisherTextFrame frame = page.TextFrames.Single(item => item.Id == 327);
        Assert.Equal(new uint[] { 345, 346 }, frame.WrapObjectIds);
        OfficeDrawingImage picture = PublisherNativeTests.Elements(page.CreateDrawing()).OfType<OfficeDrawingImage>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-346") == true);
        OfficeImagePlacement bounds = picture.Projection.Placement;
        OfficeDrawingRichText[] regions = PublisherNativeTests.Elements(page.CreateDrawing()).OfType<OfficeDrawingRichText>()
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
