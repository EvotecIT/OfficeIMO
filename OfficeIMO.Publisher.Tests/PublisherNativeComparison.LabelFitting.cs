using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed partial class PublisherNativeComparisonTests {
    [Fact]
    public void Labelled_exact_frames_force_only_the_first_line_and_preserve_the_source() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextLabelFitting/native-textlabelfitting.pub"));
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
        PublisherDocument document = PublisherDocument.Load(source, new PublisherReadOptions {
            Fonts = new OfficeFontFaceCollection().Add("Arial", font)
        });
        NativeCase contract = ReadContract("TextLabelFitting");
        string[] profiles = contract.Expected.Shapes.Where(s => s.LabelFitting?.Kind == "Tight" && s.LabelFitting.Rule == 4)
            .Select(s => s.Name).ToArray();
        int checkedFrames = 0;
        for (int page = 0; page < contract.NativePages.Length; page++) {
            foreach (NativeShape native in contract.NativePages[page].Shapes.Where(s => profiles.Contains(s.Name))) {
                PublisherTextFrame frame = document.Pages[page].TextFrames.Single(f => Near(f.X, native.Left) && Near(f.Y, native.Top));
                PublisherTextStory story = document.TextStories.Single(s => s.Id == frame.StoryId);
                Assert.Equal(native.Text!.Replace("\r", "\n"), story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value));
                Assert.Equal("First gyjp\nSecond gyjp\nThird gyjp\n", story.Text);
                Assert.All(story.Paragraphs.Take(3), p => Assert.Equal("*", p.Label!.Run.Text));
                Assert.True(frame.HasOverflow);
                checkedFrames++;
            }
        }
        Assert.Equal(10, checkedFrames);
        Assert.Equal(source, document.ToPubBytes());
        // Exact line allocations and overhang are portable under substitution;
        // native Arial baseline and pixel agreement are separate corpus evidence.
        PublisherTextFrame first = document.Pages[0].TextFrames[0];
        Assert.Equal(6, first.Height);
        Assert.True(InkBelowFrame(OfficeDrawingRasterRenderer.Render(document.Pages[0].CreateDrawing()), first) > 0);
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.True(InkBelowFrame(OfficeDrawingRasterRenderer.Render(pdf.Pages[0].ToDrawing()), first) > 0);
    }
}
