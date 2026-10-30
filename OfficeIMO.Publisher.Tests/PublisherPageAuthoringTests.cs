using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Internal;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPageAuthoringTests {
    [Theory]
    [InlineData("TextBlank")]
    [InlineData("Brochure")]
    [InlineData("TableLabel")]
    public void Appended_pages_inherit_native_size_and_existing_content_keeps_its_source_provenance(string caseName) {
        PublisherDocument document = Load(caseName);
        PublisherPage first = document.Pages[0];
        var source = document.SourceReadReport;
        int records = document.ReadReport.InspectedRecordCount;
        string[] originalSvg = document.Pages.Select((_, index) => document.ToSvg(index)).ToArray();
        PublisherPage added = document.AddPage();
        Assert.True(document.HasNativeSource); Assert.True(first.HasNativeOrigin); Assert.False(added.HasNativeOrigin);
        Assert.Null(added.MasterPageId); Assert.Empty(added.TextFrames); Assert.Empty(added.Tables); Assert.Empty(added.Shapes);
        Assert.Equal(first.Width, added.Width); Assert.Equal(first.Height, added.Height);
        Assert.Equal(originalSvg.Length + 1, document.Pages.Count);
        for (int index = 0; index < originalSvg.Length; index++) Assert.Equal(originalSvg[index], document.ToSvg(index));
        PublisherTextFrame created = document.AddTextFrame(added.Id, 20, 30, 250, 50, PublisherAuthoringTests.Paragraphs("New physical page"));
        Assert.Same(created, Assert.Single(document.Pages[document.Pages.Count - 1].TextFrames));
        Assert.True(document.Pages[0].HasNativeOrigin); Assert.False(document.Pages[document.Pages.Count - 1].HasNativeOrigin);
        Assert.Same(source, document.SourceReadReport); Assert.Equal(source.SourceObjectCount, document.ReadReport.SourceObjectCount);
        Assert.Equal(records, document.ReadReport.InspectedRecordCount);
        if (caseName == "TableLabel") {
            PublisherTable table = Assert.Single(document.Pages.SelectMany(page => page.Tables));
            Assert.True(table.CanEditCellText, table.CellTextEditLimitation);
            document.ReplaceTableCellParagraphs(table.Id, 0, 0, PublisherAuthoringTests.Paragraphs("Retained table ownership"));
            Assert.Equal("Retained table ownership", Assert.Single(document.Pages.SelectMany(page => page.Tables)).Cells[0].Text);
        }
        var pdfOptions = new PdfOptions();
        if (caseName == "TableLabel") pdfOptions.UseTextFallbacks();
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes(pdfOptions));
        Assert.Equal(document.Pages.Count, pdf.Pages.Count);
        Assert.Equal(first.Width, pdf.Pages[pdf.Pages.Count - 1].GetPageSize().Width, 4);
        Assert.Equal(first.Height, pdf.Pages[pdf.Pages.Count - 1].GetPageSize().Height, 4);
        Assert.Contains("New physical page", pdf.Pages[pdf.Pages.Count - 1].ExtractText());
    }

    [Fact]
    public void Direct_master_reassignment_changes_only_inherited_paint_and_keeps_definitions_owned_once() {
        PublisherDocument document = Load("MultipleMasters");
        PublisherPage[] oldPages = document.Pages.ToArray();
        PublisherPage[] masters = document.MasterPages.ToArray();
        var source = document.SourceReadReport; var stories = document.TextStories.Select(story => (story.Id, story.Text)).ToArray();
        document.SetPageMaster(oldPages[0].Id, masters[1].Id);
        document.SetPageMaster(oldPages[2].Id, null);
        Assert.Contains("Master Beta", PublisherAuthoringTests.PaintedText(document.Pages[0]));
        Assert.DoesNotContain("Master Alpha", PublisherAuthoringTests.PaintedText(document.Pages[0]));
        Assert.Contains("Master Beta", PublisherAuthoringTests.PaintedText(document.Pages[1]));
        Assert.DoesNotContain("Master Alpha", PublisherAuthoringTests.PaintedText(document.Pages[2]));
        Assert.DoesNotContain("Master Beta", PublisherAuthoringTests.PaintedText(document.Pages[2]));
        Assert.Contains("Master Alpha", PublisherAuthoringTests.PaintedText(oldPages[0]));
        Assert.Equal(masters[0].Id, oldPages[2].MasterPageId); Assert.Null(document.Pages[2].MasterPageId);
        Assert.Equal(stories, document.TextStories.Select(story => (story.Id, story.Text)).ToArray());
        Assert.Equal(masters.Select(master => master.Id), document.MasterPages.Select(master => master.Id));
        Assert.All(document.Pages, page => Assert.Single(page.TextFrames));
        Assert.All(document.MasterPages, master => Assert.Single(master.TextFrames));
        Assert.All(document.Pages.Concat(document.MasterPages), page => Assert.True(page.HasNativeOrigin));
        Assert.Same(source, document.SourceReadReport);
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Contains("Master Beta", pdf.Pages[0].ExtractText()); Assert.DoesNotContain("Master Alpha", pdf.Pages[0].ExtractText());
        Assert.DoesNotContain("Master Alpha", pdf.Pages[2].ExtractText());
        Assert.Throws<NotSupportedException>(() => document.SetPageMaster(masters[0].Id, masters[1].Id));
    }

    [Fact]
    public void New_pages_keep_each_imported_paired_masters_qualified_applied_registration() {
        PublisherDocument document = Load("TwoPageMaster");
        PublisherPage[] originals = document.Pages.ToArray();
        for (int index = 1; index < originals.Length; index++) {
            PublisherPage original = originals[index];
            PublisherPage master = document.MasterPages.Single(item => item.Id == original.MasterPageId);
            PublisherTextFrame marker = Assert.Single(master.TextFrames);
            OfficeDrawingRichText expected = Marker(original, marker.Id);
            PublisherPage added = document.AddPage(master.Id);
            OfficeDrawingRichText actual = Marker(added, marker.Id);
            Assert.Equal(expected.X, actual.X, 6); Assert.Equal(expected.Y, actual.Y, 6);
            Assert.Equal(expected.Width, actual.Width, 6); Assert.Equal(expected.Height, actual.Height, 6);
            Assert.Equal(PublisherNativeTests.Text(expected), PublisherNativeTests.Text(actual));
            Assert.Empty(added.TextFrames); Assert.True(master.HasNativeOrigin); Assert.False(added.HasNativeOrigin);
        }
        PublisherPage created = document.AddMasterPage("Centered authored master");
        PublisherTextFrame centered = document.AddTextFrame(created.Id, 22, 32, 200, 30, PublisherAuthoringTests.Paragraphs("Centered marker"));
        document.SetPageMaster(originals[0].Id, created.Id);
        OfficeDrawingRichText paint = Marker(document.Pages[0], centered.Id);
        Assert.Equal(22, paint.X); Assert.Equal(32, paint.Y);
        Assert.False(document.MasterPages.Single(master => master.Id == created.Id).HasNativeOrigin);
        Assert.Throws<NotSupportedException>(() => document.AddLinkedTextFrame(centered.Id, originals[0].Id, 20, 70, 200, 30));
    }

    [Fact]
    public void Master_picture_assignment_repeats_paint_without_copying_assets_or_physical_placements() {
        PublisherDocument document = Load("PictureOwnership");
        PublisherPage master = Assert.Single(document.MasterPages);
        PublisherPicture inherited = Assert.Single(master.Pictures);
        PublisherPage original = document.Pages[0]; PublisherImage asset = Assert.Single(document.Images);
        int objects = document.ReadReport.CurrentObjectCount;
        document.SetPageMaster(original.Id, null);
        PublisherPage added = document.AddPage(master.Id);
        Assert.Equal(2, Images(document.Pages[0])); Assert.Equal(1, Images(document.Pages[1])); Assert.Equal(1, Images(added));
        Assert.Equal(3, Images(original));
        Assert.Equal(2, document.Pages[0].Pictures.Count); Assert.Empty(added.Pictures);
        Assert.Equal(inherited.Id, Assert.Single(document.MasterPages[0].Pictures).Id);
        Assert.Same(asset, Assert.Single(document.Images)); Assert.Equal(objects, document.ReadReport.CurrentObjectCount);
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(new[] { 2, 1, 1 }, pdf.Pages.Select(page => page.GetImagePlacements().Count));
    }

    [Theory]
    [InlineData("unqualified-master")]
    [InlineData("master-background")]
    [InlineData("nested-master")]
    [InlineData("page-background")]
    public void Unsupported_native_page_forms_remain_inspectable_but_reject_unsafe_master_assignment(string variant) {
        PublisherDocument document = PageVariant(variant);
        uint page = document.Pages[0].Id, master = document.MasterPages[0].Id;
        var pages = document.Pages; var masters = document.MasterPages; var source = document.SourceReadReport;
        string before = document.ToSvg();
        Assert.Throws<NotSupportedException>(() => document.SetPageMaster(page, master));
        if (variant != "page-background") Assert.Throws<NotSupportedException>(() => document.AddPage(master));
        if (variant == "unqualified-master")
            Assert.Throws<NotSupportedException>(() => document.AddTextFrame(master, 12, 12, 100, 30, PublisherAuthoringTests.Paragraphs("Unqualified")));
        Assert.Same(pages, document.Pages); Assert.Same(masters, document.MasterPages); Assert.Same(source, document.SourceReadReport);
        Assert.Equal(before, document.ToSvg()); Assert.NotEmpty(document.ReadReport.FidelityDiagnostics);
    }

    private static PublisherDocument Load(string name) => PublisherDocument.Load(PublisherNativeTests.Fixture(
        "NativeComparison/" + name + "/native-" + name.ToLowerInvariant() + ".pub"));
    private static OfficeDrawingRichText Marker(PublisherPage page, uint id) => Assert.Single(PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingRichText>(), text => text.SourceElementIds?.Contains("publisher-object-" + id) == true);
    private static int Images(PublisherPage page) => PublisherNativeTests.Elements(page.CreateDrawing()).OfType<OfficeDrawingImage>().Count();

    // Decoded variants exercise edit admission only; the unchanged controls above
    // carry independent Publisher-produced registration and rendering evidence.
    private static PublisherDocument PageVariant(string variant) {
        byte[] bytes = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextBlank/native-textblank.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        var context = new PublisherParseContext(new PublisherReadOptions(), default);
        var data = new PublisherBinaryData(compound!.Streams["Contents"], "Contents");
        PublisherContents source = new PublisherContentsReader(data, context).Read();
        PublisherTextFrameCatalogReader.Read(data, source, context);
        PublisherQuillText text = new PublisherQuillTextReader(new PublisherBinaryData(compound.Streams["Quill/QuillSub/CONTENTS"], "Quill"), context).Read(source.Palette);
        compound.Streams.TryGetValue("Escher/EscherDelayStm", out byte[]? delay);
        PublisherEscherData artwork = new PublisherEscherReader(new PublisherBinaryData(compound.Streams["Escher/EscherStm"], "Escher"),
            delay == null ? null : new PublisherBinaryData(delay, "EscherDelay"), context).Read();
        int index = source.Pages.FindIndex(page => page.IsMaster == (variant != "page-background"));
        PublisherSourcePage original = source.Pages[index];
        source.Pages[index] = new PublisherSourcePage(original.Id, original.Name,
            variant == "nested-master" ? original.Id : original.Master,
            variant is "master-background" or "page-background" ? 900U : original.Background,
            original.IsMaster, variant == "unqualified-master" ? 99U : original.NativeLayout);
        return new PublisherModelState(source, text, artwork, context, new PublisherNativeSource(bytes, compound)).Project(default);
    }
}
