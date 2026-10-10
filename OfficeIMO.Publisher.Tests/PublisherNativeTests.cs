using OfficeIMO.Drawing;
using System.Xml.Linq;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherNativeTests {
    [Theory]
    [InlineData("Simple.pub", 1, 2, 2)]
    [InlineData("Sample.pub", 2, 6, 7)]
    [InlineData("Sample_2010.pub", 2, 6, 7)]
    public void Native_publications_preserve_page_order_and_story_inventory(string file, int pages, int stories, int objects) {
        PublisherDocument document = PublisherDocument.Load(Fixture(file));
        Assert.Equal(pages, document.Pages.Count);
        Assert.Single(document.MasterPages);
        Assert.Equal(stories, document.TextStories.Count);
        Assert.Equal(objects, document.ReadReport.SourceObjectCount);
        Assert.Equal(objects, document.ReadReport.ProjectedObjectCount);
        Assert.Equal(595.27559, document.Pages[0].Width, 4);
        Assert.Equal(841.88976, document.Pages[0].Height, 4);
        Assert.All(document.Pages, page => Assert.Equal(document.MasterPages[0].Id, page.MasterPageId));
    }

    [Fact]
    public void Native_character_styles_and_frame_coordinates_match_independent_reader_evidence() {
        PublisherDocument document = PublisherDocument.Load(Fixture("Sample_2010.pub"));
        OfficeRichTextRun bold = document.TextStories.Single(story => story.Id == 3).Paragraphs[0].Runs[0];
        Assert.Equal("Arial", bold.FontFamily);
        Assert.Equal(20, bold.FontSize);
        Assert.True(bold.Bold);
        Assert.True(bold.Italic);
        OfficeRichTextRun normal = document.TextStories.Single(story => story.Id == 1).Paragraphs[0].Runs[0];
        Assert.Equal("Times New Roman", normal.FontFamily);
        Assert.Equal(10, normal.FontSize);
        Assert.False(normal.Bold);
        Assert.False(normal.Italic);
        PublisherTextFrame frame = document.Pages[0].TextFrames.Single(item => item.Id == 293);
        OfficeDrawingRichText column = Elements(document.Pages[0].Drawing).OfType<OfficeDrawingRichText>()
            .Single(item => item.SourceElementIds?.Contains("publisher-object-293") == true);
        Assert.Equal(85.03937, frame.X, 4);
        Assert.Equal(70.86614, frame.Y, 4);
        Assert.Equal(399.68504, frame.Width, 4);
        Assert.Equal(frame.X + 2.88, column.X, 4);
        Assert.Equal(frame.Y + 2.88, column.Y, 4);
    }

    [Fact]
    public void Table_text_maps_to_native_rows_and_columns_without_consuming_adjacent_cell_text() {
        PublisherDocument document = PublisherDocument.Load(Fixture("Sample.pub"));
        OfficeDrawingRichText[] cells = Elements(document.Pages[1].Drawing).OfType<OfficeDrawingRichText>()
            .Where(item => item.SourceElementIds?.Contains("publisher-object-299") == true).ToArray();
        Assert.Equal(new[] { "Table on page 2", "Top right", "P2 table left", "P2 table right", "Bottom Left", "Bottom Right" }, cells.Select(Text));
        Assert.Equal(107.71654, cells[0].X, 4);
        Assert.Equal(221.10236, cells[0].Y, 4);
        Assert.Equal(140.31496, cells[0].Width, 4);
        Assert.Equal(cells[0].X + cells[0].Width, cells[1].X, 6);
        Assert.Equal(cells[0].Y + cells[0].Height, cells[2].Y, 6);
    }

    [Theory]
    [InlineData("SampleBrochure.pub", 2, 20, 8)]
    [InlineData("SampleNewsletter.pub", 4, 44, 9)]
    public void Publication_corpus_recovers_grouped_artwork_and_embedded_images(string file, int pages, int stories, int images) {
        PublisherDocument document = PublisherDocument.Load(Fixture(file));
        Assert.Equal(pages, document.Pages.Count);
        Assert.Equal(stories, document.TextStories.Count);
        Assert.Equal(images, document.Images.Count);
        Assert.All(document.Images, image => Assert.True(image.ByteCount > 0));
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_GROUP_COORDINATES_UNRESOLVED");
        Assert.True(document.ReadReport.HasLoss);
        Assert.All(document.Pages, page => Assert.NotEmpty(Elements(page.Drawing).OfType<OfficeDrawingImage>()));
    }

    [Fact]
    public void Publication_text_overflow_is_reported_while_complete_stories_remain_available() {
        PublisherDocument document = PublisherDocument.Load(Fixture("SampleNewsletter.pub"));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_FRAME_OVERFLOW" && item.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINKED_TEXT_LAYOUT_APPROXIMATED");
        Assert.Contains(document.TextStories, story => story.Text.Contains("Living and Learning in"));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_WRAP_APPROXIMATED");
    }

    [Fact]
    public void Publisher_gif_envelope_is_recovered_as_gif_and_exported_with_its_actual_media_type() {
        PublisherDocument document = PublisherDocument.Load(Fixture("SampleBrochure.pub"));
        PublisherImage image = document.Images.Single(item => item.Id == 2);
        Assert.Equal("image/gif", image.ContentType);
        Assert.Equal("GIF89a", System.Text.Encoding.ASCII.GetString(image.GetBytes(), 0, 6));
        byte[] bytes = image.GetBytes(); bytes[0] = 0;
        Assert.Equal((byte)'G', image.GetBytes()[0]);
        string svg = document.ToSvg(1);
        Assert.Contains("data:image/gif;base64,", svg);
    }

    [Fact]
    public void Svg_export_preserves_physical_page_size_and_source_loss_evidence() {
        PublisherDocument document = PublisherDocument.Load(Fixture("Sample.pub"));
        var result = document.ToSvgResult(1);
        XElement svg = XElement.Parse(result.Value);
        Assert.EndsWith("pt", svg.Attribute("width")!.Value);
        Assert.Contains("Table on page 2", result.Value);
        Assert.Contains("Bottom Right", result.Value);
        Assert.Same(document.ReadReport, result.Report.ReadReport);
        Assert.Throws<OfficeConversionException>(() => result.RequireNoLoss());
        Assert.Throws<ArgumentOutOfRangeException>(() => document.ToSvg(-1));
    }

    [Theory]
    [InlineData("Sample2000.pub")]
    [InlineData("Sample98.pub")]
    public void Earlier_native_generations_are_rejected_instead_of_returning_text_salvage(string file) {
        Assert.Throws<NotSupportedException>(() => PublisherDocument.Load(Fixture(file)));
    }

    internal static string Fixture(string file) => Path.Combine(AppContext.BaseDirectory, "Fixtures", file);
    internal static IEnumerable<OfficeDrawingElement> Elements(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            yield return element;
            if (element is OfficeDrawingGroup group) foreach (OfficeDrawingElement child in Elements(group.InnerDrawing)) yield return child;
            if (element is OfficeDrawingEffectGroup transformed) foreach (OfficeDrawingElement child in Elements(transformed.InnerDrawing)) yield return child;
        }
    }
    internal static string Text(OfficeDrawingRichText frame) => string.Join("\n", frame.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
}
