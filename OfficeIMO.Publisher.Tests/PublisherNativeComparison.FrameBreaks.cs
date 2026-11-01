using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed partial class PublisherNativeComparisonTests {
    [Fact]
    public void Exact_frame_ranges_force_the_first_line_and_require_full_later_allocations() {
        string fixture = PublisherNativeTests.Fixture("NativeComparison/TextFrameBreaks/native-textframebreaks.pub");
        byte[] source = File.ReadAllBytes(fixture);
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
        var fonts = new OfficeFontFaceCollection().Add("Arial", font).Add("Courier New", font);
        PublisherDocument document = PublisherDocument.Load(source, new PublisherReadOptions { Fonts = fonts });
        NativeCase contract = ReadContract("TextFrameBreaks");
        var profiles = contract.Expected.Shapes
            .Where(shape => shape.FrameBreak?.Kind == "FirstExact" || shape.FrameBreak?.Kind == "TightExact")
            .Select(shape => shape.Name).ToArray();
        int checkedFrames = 0;
        for (int pageIndex = 0; pageIndex < contract.NativePages.Length; pageIndex++) {
            foreach (NativeShape shape in contract.NativePages[pageIndex].Shapes.Where(shape => profiles.Contains(shape.Name))) {
                PublisherTextFrame frame = document.Pages[pageIndex].TextFrames.Single(frame =>
                    Near(frame.X, shape.Left) && Near(frame.Y, shape.Top));
                PublisherTextStory story = document.TextStories.Single(story => story.Id == frame.StoryId);
                string expected = shape.Text!.Replace("\r", "\n");
                Assert.Equal(expected, story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value));
                Assert.Equal(expected.Length < story.Text.Length, frame.HasOverflow);
                checkedFrames++;
            }
        }
        Assert.Equal(12, checkedFrames);
        Assert.Equal(source, document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        // Font substitution is deliberate here: these wide exact-line frames
        // establish allocation/transport, not native font-baseline agreement.
        PublisherPage shortPage = document.Pages[4];
        PublisherTextFrame shortFrame = Assert.Single(shortPage.TextFrames);
        Assert.Equal(12, shortFrame.Height, 3);
        OfficeRasterImage direct = OfficeDrawingRasterRenderer.Render(shortPage.CreateDrawing());
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        OfficeRasterImage reopened = OfficeDrawingRasterRenderer.Render(pdf.Pages[4].ToDrawing());
        Assert.True(InkBelowFrame(direct, shortFrame) > 0);
        Assert.True(InkBelowFrame(reopened, shortFrame) > 0);
    }

    private static int InkBelowFrame(OfficeRasterImage image, PublisherTextFrame frame) {
        int count = 0;
        for (int y = (int)Math.Ceiling(frame.Y + frame.Height) + 1; y < frame.Y + frame.Height + 12; y++)
            for (int x = (int)frame.X; x < frame.X + 100; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.A > 0 && color.R < 128 && color.G < 128 && color.B < 128) count++;
            }
        return count;
    }
}
