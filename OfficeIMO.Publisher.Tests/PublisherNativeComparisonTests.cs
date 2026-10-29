using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using Xunit.Abstractions;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherNativeComparisonTests {
    private readonly ITestOutputHelper _output;
    public PublisherNativeComparisonTests(ITestOutputHelper output) => _output = output;

    [Theory]
    [InlineData("Blank")]
    [InlineData("BlankPages")]
    [InlineData("Text")]
    [InlineData("TextLong")]
    [InlineData("TextBlank")]
    [InlineData("Rectangle")]
    [InlineData("Image")]
    [InlineData("Newsletter")]
    [InlineData("TableLabel")]
    [InlineData("Brochure")]
    [InlineData("MultipleMasters")]
    [InlineData("TwoPageMaster")]
    public void Native_physical_pages_keep_their_order_size_and_master_reference(string caseId) {
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);

        uint[] nativePages = contract.NativePages.Select(page => NativeId(page.PageId)).ToArray();
        Assert.Equal(contract.Expected.Pages, contract.NativePages.Length);
        Assert.Equal(nativePages, publication.Pages.Select(page => page.Id));
        for (int index = 0; index < publication.Pages.Count; index++) {
            PublisherPage page = publication.Pages[index];
            AssertNear(contract.Expected.WidthPoints, page.Width);
            AssertNear(contract.Expected.HeightPoints, page.Height);
            if (contract.NativePages[index].Master != null)
                Assert.Equal(NativeId(contract.NativePages[index].Master!.PageId), page.MasterPageId);
        }
    }

    [Theory]
    [InlineData("Blank")]
    [InlineData("BlankPages")]
    [InlineData("Text")]
    [InlineData("TextLong")]
    [InlineData("TextBlank")]
    [InlineData("Rectangle")]
    [InlineData("Image")]
    [InlineData("Newsletter")]
    [InlineData("TableLabel")]
    [InlineData("Brochure")]
    [InlineData("MultipleMasters")]
    [InlineData("TwoPageMaster")]
    public void Pdf_export_keeps_every_native_document_page_including_blank_pages(string caseId) {
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        string referencePath = FilePath(caseId, contract.Files.Reference);
        AssertHash(referencePath, contract.Files.Reference.Sha256);

        PdfReadDocument nativePdf = PdfReadDocument.Open(File.ReadAllBytes(referencePath));
        PdfOptions options = new PdfOptions();
        // PDF's standard fonts intentionally use WinAnsi. The multilingual
        // control exercises the public Unicode-font configuration explicitly.
        if (caseId == "TableLabel") options.UseTextFallbacks();
        PdfDocumentConversionResult conversion = publication.ToPdfDocumentResult(options);
        PdfReadDocument exported = PdfReadDocument.Open(conversion.ToBytes());
        Assert.Equal(contract.Expected.Pages, nativePdf.Pages.Count);
        Assert.Equal(nativePdf.Pages.Count, exported.Pages.Count);
        Assert.Contains(publication.ReadReport, conversion.SourceConversionReports);
        if (caseId == "TableLabel")
            Assert.Contains(contract.Expected.Text[0], exported.ExtractText());
    }

    [Theory]
    [InlineData("Text")]
    [InlineData("TextLong")]
    [InlineData("TextBlank")]
    [InlineData("Newsletter")]
    [InlineData("Brochure")]
    [InlineData("MultipleMasters")]
    [InlineData("TwoPageMaster")]
    public void Complete_source_stories_survive_independently_of_native_frame_breaks(string caseId) {
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        // Publisher also creates empty stories for text-capable panel shapes.
        // The authored text contract concerns complete nonempty source stories.
        string[] stories = publication.TextStories.Select(story => NormalizeParagraphs(story.Text)).Where(text => text.Length > 0).ToArray();

        Assert.Equal(contract.Expected.Text.Length, stories.Length);
        foreach (string text in contract.Expected.Text)
            Assert.Contains(NormalizeParagraphs(text), stories);
    }

    [Theory]
    [InlineData("Text")]
    [InlineData("TextLong")]
    [InlineData("TextBlank")]
    [InlineData("Newsletter")]
    [InlineData("TableLabel")]
    [InlineData("Brochure")]
    [InlineData("MultipleMasters")]
    [InlineData("TwoPageMaster")]
    public void Text_frames_keep_native_page_ownership_and_geometry(string caseId) {
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        if (caseId == "Newsletter") WriteFrameInventory(publication);

        foreach (NativePage nativePage in contract.NativePages) {
            PublisherPage page = Assert.Single(publication.Pages, item => item.Id == NativeId(nativePage.PageId));
            AssertTextFrames(page, nativePage.Shapes);
        }
    }

    [Fact]
    public void Brochure_keeps_three_panel_markers_and_landscape_pdf_geometry_on_both_sides() {
        const string caseId = "Brochure";
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        string reference = FilePath(caseId, contract.Files.Reference);
        AssertHash(reference, contract.Files.Reference.Sha256);
        PdfReadDocument nativePdf = PdfReadDocument.Open(File.ReadAllBytes(reference));
        PdfReadDocument exported = PdfReadDocument.Open(publication.ToPdfBytes());

        Assert.NotNull(contract.Expected.PanelWidthPoints);
        AssertNear(contract.Expected.WidthPoints / 3, contract.Expected.PanelWidthPoints!.Value);
        for (int index = 0; index < contract.Expected.Pages; index++) {
            AssertNear(contract.Expected.WidthPoints, nativePdf.Pages[index].GetPageSize().Width);
            AssertNear(contract.Expected.HeightPoints, nativePdf.Pages[index].GetPageSize().Height);
            AssertNear(contract.Expected.WidthPoints, exported.Pages[index].GetPageSize().Width);
            AssertNear(contract.Expected.HeightPoints, exported.Pages[index].GetPageSize().Height);
            ExpectedShape[] markers = contract.Expected.Shapes.Where(shape => shape.Kind == "Rectangle" && shape.PageIndex == index + 1).ToArray();
            Assert.Equal(3, markers.Length);
            foreach (ExpectedShape marker in markers) {
                NativeShape native = Assert.Single(contract.NativePages[index].Shapes, shape => shape.Name == marker.Name);
                AssertNear(marker.Left, native.Left);
                AssertNear(marker.Top, native.Top);
                AssertNear(marker.Width, native.Width);
                AssertNear(marker.Height, native.Height);
                OfficeDrawingShape projected = Assert.Single(PublisherNativeTests.Elements(publication.Pages[index].Drawing).OfType<OfficeDrawingShape>(), shape => Near(marker.Left, shape.X) && Near(marker.Top, shape.Y));
                AssertNear(marker.Width, projected.Shape.Width);
                AssertNear(marker.Height, projected.Shape.Height);
            }
        }
    }

    [Theory]
    [InlineData("MultipleMasters")]
    [InlineData("TwoPageMaster")]
    public void Distinct_native_masters_keep_definitions_assignments_and_projected_markers(string caseId) {
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        WriteFrameInventory(publication);
        Assert.Equal(contract.NativeMasterPages.Select(master => NativeId(master.PageId)), publication.MasterPages.Select(master => master.Id));
        Assert.Equal(caseId == "TwoPageMaster" ? 2 : 0, contract.NativeMasterPages.Count(master => master.IsTwoPageMaster));
        foreach (NativeMaster nativeMaster in contract.NativeMasterPages) {
            PublisherPage master = Assert.Single(publication.MasterPages, item => item.Id == NativeId(nativeMaster.PageId));
            AssertTextFrames(master, nativeMaster.Shapes);
        }

        string reference = FilePath(caseId, contract.Files.Reference);
        AssertHash(reference, contract.Files.Reference.Sha256);
        PdfReadDocument nativePdf = PdfReadDocument.Open(File.ReadAllBytes(reference));
        PdfReadDocument exported = PdfReadDocument.Open(publication.ToPdfBytes());
        string[] allMasterMarkers = contract.NativeMasterPages.SelectMany(master => master.Shapes)
            .Where(shape => shape.Type == 17).Select(shape => NormalizeParagraphs(shape.Text!)).ToArray();
        Assert.Equal(contract.Expected.Pages, contract.Expected.MasterAssignments.Length);
        foreach (ExpectedMasterAssignment assignment in contract.Expected.MasterAssignments) {
            int index = assignment.PageIndex - 1;
            NativeMaster nativeMaster = Assert.Single(contract.NativeMasterPages, master => master.PageId == assignment.MasterPageId);
            PublisherPage page = publication.Pages[index];
            Assert.Equal(NativeId(assignment.MasterPageId), page.MasterPageId);
            string[] assignedMarkers = nativeMaster.Shapes.Where(shape => shape.Type == 17).Select(shape => NormalizeParagraphs(shape.Text!)).ToArray();
            OfficeDrawingRichText[] projectedFrames = PublisherNativeTests.Elements(page.Drawing).OfType<OfficeDrawingRichText>().ToArray();
            Assert.Equal(assignedMarkers.Length, projectedFrames.Count(frame => allMasterMarkers.Contains(NormalizeParagraphs(PublisherNativeTests.Text(frame)))));
            foreach (NativeShape marker in nativeMaster.Shapes.Where(shape => shape.Type == 17)) {
                string text = NormalizeParagraphs(marker.Text!);
                OfficeDrawingRichText projected = Assert.Single(projectedFrames, frame => NormalizeParagraphs(PublisherNativeTests.Text(frame)) == text);
                PdfTextSpan referenceSpan = Assert.Single(nativePdf.Pages[index].GetTextSpans(), span => span.IsVisible && NormalizeParagraphs(span.Text) == text);
                PdfTextSpan exportedSpan = Assert.Single(exported.Pages[index].GetTextSpans(), span => span.IsVisible && NormalizeParagraphs(span.Text) == text);
                _output.WriteLine($"Page {index + 1}, marker {text}, definition X {marker.Left}, native PDF {referenceSpan.X},{referenceSpan.Y}, drawing X {projected.X}, exported PDF {exportedSpan.X},{exportedSpan.Y}");
                // Applied positions come from the independent native PDF, which
                // can differ from the master definition's local coordinates.
                // Native text-origin rounding permits 0.1pt horizontally; the
                // qualified fonts differ by less than 1pt in baseline metrics.
                Assert.InRange(Math.Abs(referenceSpan.X - projected.X), 0, 0.1);
                Assert.InRange(Math.Abs(referenceSpan.X - exportedSpan.X), 0, 0.1);
                Assert.InRange(Math.Abs(referenceSpan.Y - exportedSpan.Y), 0, 1);
                AssertNear(0, referenceSpan.RotationDegrees);
                AssertNear(0, exportedSpan.RotationDegrees);
                AssertNear(marker.Top, projected.Y);
                AssertNear(marker.Width, projected.Width);
                AssertNear(marker.Height, projected.Height);
                Assert.Contains(text, nativePdf.Pages[index].ExtractText());
                Assert.Contains(text, exported.Pages[index].ExtractText());
            }
            foreach (string otherMarker in allMasterMarkers.Except(assignedMarkers)) {
                Assert.DoesNotContain(otherMarker, nativePdf.Pages[index].ExtractText());
                Assert.DoesNotContain(otherMarker, exported.Pages[index].ExtractText());
            }
        }
    }

    [Fact]
    public void Newsletter_links_one_complete_story_across_its_two_native_frames() {
        const string caseId = "Newsletter";
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        WriteFrameInventory(publication);
        PublisherTextFrame first = Assert.Single(publication.Pages[0].TextFrames, frame => frame.NextFrameId.HasValue);
        PublisherTextFrame last = Assert.Single(publication.Pages[1].TextFrames, frame => frame.PreviousFrameId.HasValue);

        Assert.Equal(first.Id, last.PreviousFrameId);
        Assert.Equal(last.Id, first.NextFrameId);
        Assert.Equal(first.StoryId, last.StoryId);
        PublisherTextStory story = Assert.Single(publication.TextStories, item => item.Id == first.StoryId);
        Assert.Equal(NormalizeParagraphs(contract.Expected.Text[1]), NormalizeParagraphs(story.Text));
        // Text ranges describe managed placement. This does not pin Publisher's
        // cached line/frame breaks or imply that every character fits visually.
        Assert.Equal(0, first.TextStart);
        Assert.NotNull(first.TextLength);
        Assert.Equal(first.TextLength, last.TextStart);
    }

    [Theory]
    [InlineData("Image")]
    [InlineData("Newsletter")]
    [InlineData("TableLabel")]
    public void Embedded_picture_keeps_source_pixels_and_native_placement(string caseId) {
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        Assert.NotNull(contract.Files.SourceImage);
        string sourcePath = FilePath(caseId, contract.Files.SourceImage!);
        AssertHash(sourcePath, contract.Files.SourceImage!.Sha256);
        PublisherImage asset = Assert.Single(publication.Images);
        Assert.True(OfficeRasterImageDecoder.TryDecode(File.ReadAllBytes(sourcePath), out OfficeRasterImage? expected));
        Assert.True(OfficeRasterImageDecoder.TryDecode(asset.GetBytes(), out OfficeRasterImage? actual));
        Assert.Equal(expected!.Width, actual!.Width);
        Assert.Equal(expected.Height, actual.Height);
        foreach (var point in new[] { (X: 0, Y: 0), (X: 31, Y: 0), (X: 0, Y: 31), (X: 31, Y: 31) })
            Assert.Equal(expected.GetPixel(point.X, point.Y), actual.GetPixel(point.X, point.Y));

        ExpectedShape geometry = Assert.Single(contract.Expected.Shapes, shape => shape.Kind == "Image");
        OfficeDrawingImage projected = Assert.Single(publication.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing)).OfType<OfficeDrawingImage>());
        AssertNear(geometry.Left, projected.Projection.X);
        AssertNear(geometry.Top, projected.Projection.Y);
        AssertNear(geometry.Width, projected.Projection.Width);
        AssertNear(geometry.Height, projected.Projection.Height);
    }

    [Fact]
    public void Custom_label_keeps_its_native_table_cells_and_multilingual_story() {
        const string caseId = "TableLabel";
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        PublisherPage page = Assert.Single(publication.Pages);
        ExpectedShape expected = Assert.Single(contract.Expected.Shapes, shape => shape.Kind == "Table");
        PublisherTable table = Assert.Single(page.Tables);

        Assert.True(table.HasPlacement);
        Assert.True(table.HasTextMapping);
        Assert.Equal(expected.Rows, table.RowCount);
        Assert.Equal(expected.Columns, table.ColumnCount);
        AssertNear(expected.Left, table.X);
        AssertNear(expected.Top, table.Y);
        AssertNear(expected.Width, table.Width);
        AssertNear(expected.Height, table.Height);
        Assert.Equal(expected.Cells, table.Cells.OrderBy(cell => cell.RowIndex).ThenBy(cell => cell.ColumnIndex).Select(cell => NormalizeParagraphs(cell.Text)));
        Assert.All(table.Cells, cell => { Assert.Equal(1, cell.RowSpan); Assert.Equal(1, cell.ColumnSpan); });
        Assert.Contains(publication.TextStories, story => NormalizeParagraphs(story.Text) == NormalizeParagraphs(contract.Expected.Text[0]));
    }

    [Fact]
    public void Solid_rectangle_keeps_native_position_size_and_fill() {
        const string caseId = "Rectangle";
        NativeCase contract = ReadContract(caseId);
        PublisherDocument publication = Load(caseId, contract);
        ExpectedShape expected = Assert.Single(contract.Expected.Shapes);
        OfficeDrawingShape shape = Assert.Single(
            PublisherNativeTests.Elements(Assert.Single(publication.Pages).Drawing).OfType<OfficeDrawingShape>(),
            item => Near(expected.Left, item.X) && Near(expected.Top, item.Y));

        AssertNear(expected.Left, shape.X);
        AssertNear(expected.Top, shape.Y);
        AssertNear(expected.Width, shape.Shape.Width);
        AssertNear(expected.Height, shape.Shape.Height);
        Assert.Equal(OfficeShapeKind.Rectangle, shape.Shape.Kind);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), shape.Shape.FillColor);
    }

    private static NativeCase ReadContract(string caseId) {
        using var input = File.OpenRead(Path.Combine(Folder(caseId), "native-" + caseId.ToLowerInvariant() + ".json"));
        return (NativeCase)new DataContractJsonSerializer(typeof(NativeCase)).ReadObject(input)!;
    }

    private static void AssertTextFrames(PublisherPage page, NativeShape[] shapes) {
        NativeShape[] nativeFrames = shapes.Where(shape => shape.Text != null).ToArray();
        Assert.Equal(nativeFrames.Length, page.TextFrames.Count);
        foreach (NativeShape expected in nativeFrames) {
            PublisherTextFrame frame = Assert.Single(page.TextFrames, item => Near(expected.Left, item.X) && Near(expected.Top, item.Y));
            Assert.Equal(page.Id, frame.PageId);
            AssertNear(expected.Width, frame.Width);
            AssertNear(expected.Height, frame.Height);
            Assert.Equal(expected.NextLink, frame.NextFrameId.HasValue);
            Assert.Equal(expected.PreviousLink, frame.PreviousFrameId.HasValue);
        }
    }

    private void WriteFrameInventory(PublisherDocument publication) {
        foreach (PublisherPage page in publication.Pages.Concat(publication.MasterPages)) {
            string role = publication.MasterPages.Contains(page) ? "Master" : "Page";
            _output.WriteLine($"{role} {page.Id}, master {page.MasterPageId}, {page.TextFrames.Count} frames");
            foreach (PublisherTextFrame frame in page.TextFrames)
                _output.WriteLine($"  frame {frame.Id}, story {frame.StoryId}, previous {frame.PreviousFrameId}, next {frame.NextFrameId}, rect {frame.X},{frame.Y},{frame.Width},{frame.Height}, range {frame.TextStart}+{frame.TextLength}, overflow {frame.HasOverflow}");
        }
        foreach (PublisherTextStory story in publication.TextStories)
            _output.WriteLine($"Story {story.Id}, complete UTF-16 length {story.Text.Length}");
        foreach (var diagnostic in publication.ReadReport.FidelityDiagnostics)
            _output.WriteLine($"Recovery {diagnostic.Code}: {diagnostic.Message}");
    }

    private static PublisherDocument Load(string caseId, NativeCase contract) {
        string path = FilePath(caseId, contract.Files.Publication);
        AssertHash(path, contract.Files.Publication.Sha256);
        return PublisherDocument.Load(path);
    }

    private static string Folder(string caseId) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "NativeComparison", caseId);
    private static string FilePath(string caseId, NativeFile file) => Path.Combine(Folder(caseId), file.Name);
    private static uint NativeId(long nativeId) => (uint)(nativeId & 0x00FFFFFF);
    private static string NormalizeParagraphs(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
    private static bool Near(double expected, double actual) => Math.Abs(expected - actual) <= 0.01;
    private static void AssertNear(double expected, double actual) => Assert.InRange(Math.Abs(expected - actual), 0, 0.01);

    private static void AssertHash(string path, string expected) {
        using var hash = SHA256.Create();
        using var input = File.OpenRead(path);
        Assert.Equal(expected, BitConverter.ToString(hash.ComputeHash(input)).Replace("-", "").ToLowerInvariant());
    }

    [DataContract]
    private sealed class NativeCase {
        [DataMember(Name = "expected")] public ExpectedPublication Expected { get; set; } = null!;
        [DataMember(Name = "nativePages")] public NativePage[] NativePages { get; set; } = null!;
        [DataMember(Name = "nativeMasterPages")] public NativeMaster[] NativeMasterPages { get; set; } = null!;
        [DataMember(Name = "files")] public NativeFiles Files { get; set; } = null!;
    }

    [DataContract]
    private sealed class ExpectedPublication {
        [DataMember(Name = "pages")] public int Pages { get; set; }
        [DataMember(Name = "widthPoints")] public double WidthPoints { get; set; }
        [DataMember(Name = "heightPoints")] public double HeightPoints { get; set; }
        [DataMember(Name = "text")] public string[] Text { get; set; } = null!;
        [DataMember(Name = "shapes")] public ExpectedShape[] Shapes { get; set; } = null!;
        [DataMember(Name = "panelWidthPoints")] public double? PanelWidthPoints { get; set; }
        [DataMember(Name = "masterAssignments")] public ExpectedMasterAssignment[] MasterAssignments { get; set; } = null!;
    }

    [DataContract]
    private sealed class ExpectedMasterAssignment {
        [DataMember(Name = "pageIndex")] public int PageIndex { get; set; }
        [DataMember(Name = "masterPageId")] public long MasterPageId { get; set; }
    }

    [DataContract]
    private sealed class ExpectedShape {
        [DataMember(Name = "name")] public string Name { get; set; } = null!;
        [DataMember(Name = "pageIndex")] public int PageIndex { get; set; }
        [DataMember(Name = "kind")] public string Kind { get; set; } = null!;
        [DataMember(Name = "left")] public double Left { get; set; }
        [DataMember(Name = "top")] public double Top { get; set; }
        [DataMember(Name = "width")] public double Width { get; set; }
        [DataMember(Name = "height")] public double Height { get; set; }
        [DataMember(Name = "rows")] public int Rows { get; set; }
        [DataMember(Name = "columns")] public int Columns { get; set; }
        [DataMember(Name = "cells")] public string[] Cells { get; set; } = null!;
    }

    [DataContract]
    private sealed class NativePage {
        [DataMember(Name = "pageId")] public long PageId { get; set; }
        [DataMember(Name = "master")] public NativeMaster? Master { get; set; }
        [DataMember(Name = "shapes")] public NativeShape[] Shapes { get; set; } = null!;
    }

    [DataContract]
    private sealed class NativeMaster {
        [DataMember(Name = "pageId")] public long PageId { get; set; }
        [DataMember(Name = "isTwoPageMaster")] public bool IsTwoPageMaster { get; set; }
        [DataMember(Name = "shapes")] public NativeShape[] Shapes { get; set; } = null!;
    }

    [DataContract]
    private sealed class NativeShape {
        [DataMember(Name = "name")] public string Name { get; set; } = null!;
        [DataMember(Name = "text")] public string? Text { get; set; }
        [DataMember(Name = "type")] public int Type { get; set; }
        [DataMember(Name = "left")] public double Left { get; set; }
        [DataMember(Name = "top")] public double Top { get; set; }
        [DataMember(Name = "width")] public double Width { get; set; }
        [DataMember(Name = "height")] public double Height { get; set; }
        [DataMember(Name = "nextLink")] public bool NextLink { get; set; }
        [DataMember(Name = "previousLink")] public bool PreviousLink { get; set; }
    }

    [DataContract]
    private sealed class NativeFiles {
        [DataMember(Name = "publication")] public NativeFile Publication { get; set; } = null!;
        [DataMember(Name = "reference")] public NativeFile Reference { get; set; } = null!;
        [DataMember(Name = "sourceImage")] public NativeFile? SourceImage { get; set; }
    }

    [DataContract]
    private sealed class NativeFile {
        [DataMember(Name = "name")] public string Name { get; set; } = null!;
        [DataMember(Name = "sha256")] public string Sha256 { get; set; } = null!;
    }
}
