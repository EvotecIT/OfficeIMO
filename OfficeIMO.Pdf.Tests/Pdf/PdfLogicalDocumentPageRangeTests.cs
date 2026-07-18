using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void LoadPageRanges_BuildsLogicalModelForSelectedSourcePagesInCallerOrder() {
        byte[] pdf = BuildThreePageLogicalPdf();

        PdfLogicalDocument logical = PdfLogicalDocument.LoadPageRanges(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }, PdfPageRange.ParseMany("3,1-2,3"));

        Assert.Equal(4, logical.PageCount);
        Assert.Equal(new[] { 3, 1, 2, 3 }, logical.Pages.Select(page => page.PageNumber).ToArray());
        Assert.Contains(logical.Pages[0].TextBlocks, block => block.Text.Contains("Third logical page", StringComparison.Ordinal));
        Assert.Contains(logical.Pages[1].TextBlocks, block => block.Text.Contains("First logical page", StringComparison.Ordinal));
        Assert.Contains(logical.Pages[2].TextBlocks, block => block.Text.Contains("Second logical page", StringComparison.Ordinal));
        Assert.Contains(logical.Pages[3].TextBlocks, block => block.Text.Contains("Third logical page", StringComparison.Ordinal));
        Assert.Equal(2, logical.Pages.Count(page => page.PageNumber == 3));
        Assert.True(logical.HasSourcePage(3));
        Assert.Equal(new[] { 3, 3 }, logical.PagesBySourcePageNumber[3].Select(page => page.PageNumber).ToArray());
        Assert.Same(logical.Pages[0], logical.GetPages(3)[0]);
        Assert.Same(logical.Pages[3], logical.GetPages(3)[1]);
        Assert.Equal(2, logical.TextBlocks.Count(block => block.PageNumber == 3 && block.Text.Contains("Third logical page", StringComparison.Ordinal)));
        Assert.Equal(2, logical.GetElements(3).OfType<PdfLogicalTextBlock>().Count(block => block.Text.Contains("Third logical page", StringComparison.Ordinal)));

        PdfReadDocument document = PdfReadDocument.Open(pdf);
        PdfLogicalDocument fromDocument = PdfLogicalDocument.FromPageRanges(document, PdfPageRange.From(2, 2));

        PdfLogicalPage selected = Assert.Single(fromDocument.Pages);
        Assert.Equal(2, selected.PageNumber);
        Assert.Contains(selected.TextBlocks, block => block.Text.Contains("Second logical page", StringComparison.Ordinal));
    }

    [Fact]
    public void LoadPageRanges_FiltersAcroFormFieldsToSelectedSourcePages() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .TextField("First.Page", width: 120, height: 20, value: "one")
            .PageBreak()
            .TextField("Second.Page", width: 120, height: 20, value: "two")
            .PageBreak()
            .TextField("Third.Page", width: 120, height: 20, value: "three")
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.LoadPageRanges(pdf, PdfPageRange.ParseMany("2,1,2"));

        Assert.Equal(new[] { 2, 1, 2 }, logical.Pages.Select(page => page.PageNumber).ToArray());
        Assert.Equal(2, logical.FormFields.Count);
        Assert.Contains("First.Page", logical.FormFieldNames);
        Assert.Contains("Second.Page", logical.FormFieldNames);
        Assert.DoesNotContain("Third.Page", logical.FormFieldNames);
        Assert.True(logical.TryGetFormField("Second.Page", out PdfFormField? secondField));
        Assert.Equal(new[] { 2 }, secondField!.PageNumbers);
        Assert.Equal(new[] { "two" }, secondField.Values);
        Assert.Same(secondField, Assert.Single(logical.GetFormFields(2)));
        Assert.Empty(logical.GetFormFields(3));
        Assert.Equal(2, logical.GetFormWidgets("Second.Page").Count);
        Assert.All(logical.GetFormWidgets("Second.Page"), widget => Assert.Equal(2, widget.PageNumber));
        Assert.Single(logical.GetFormWidgets("First.Page"));
        Assert.Empty(logical.GetFormWidgets("Third.Page"));
        Assert.Equal(3, logical.FormWidgets.Count);
        Assert.Equal(new[] { "Second.Page", "First.Page", "Second.Page" }, logical.FormWidgets.Select(widget => widget.FieldName).ToArray());
        Assert.Equal(3, logical.GetElements(PdfLogicalElementKind.FormWidget).Count);
        Assert.Contains(logical.Pages[0].Elements, element => element.Kind == PdfLogicalElementKind.FormWidget);
        Assert.Contains(logical.Pages[1].Elements, element => element.Kind == PdfLogicalElementKind.FormWidget);
        Assert.Contains(logical.Pages[2].Elements, element => element.Kind == PdfLogicalElementKind.FormWidget);

        PdfLogicalDocument full = PdfLogicalDocument.Load(pdf);
        Assert.Contains("Third.Page", full.FormFieldNames);
        Assert.Equal(3, full.FormFields.Count);
    }

    [Fact]
    public void LoadPageRanges_FiltersNavigationObjectsToSelectedSourcePages() {
        byte[] pdf = BuildThreePageNavigationPdf();

        PdfLogicalDocument logical = PdfLogicalDocument.LoadPageRanges(pdf, PdfPageRange.ParseMany("2,1,2"));

        Assert.Equal(new[] { 2, 1, 2 }, logical.Pages.Select(page => page.PageNumber).ToArray());
        Assert.Equal(new[] { "First", "Second" }, logical.NamedDestinations.Select(destination => destination.Name).OrderBy(name => name).ToArray());
        Assert.Equal(new[] { 1, 2 }, logical.NamedDestinations.Select(destination => destination.PageNumber!.Value).OrderBy(pageNumber => pageNumber).ToArray());
        Assert.Equal(new[] { "First outline", "Second outline" }, logical.Outlines.Select(outline => outline.Title).OrderBy(title => title).ToArray());
        Assert.Equal(new[] { 1, 2 }, logical.Outlines.Select(outline => outline.PageNumber!.Value).OrderBy(pageNumber => pageNumber).ToArray());
        PdfOutlineItem secondOutline = Assert.Single(logical.Outlines, outline => outline.Title == "Second outline");
        Assert.Equal(PdfOpenActionDestinationMode.FitRectangle, secondOutline.DestinationMode);
        Assert.Equal(10D, secondOutline.DestinationLeft);
        Assert.Equal(20D, secondOutline.DestinationBottom);
        Assert.Equal(90D, secondOutline.DestinationRight);
        Assert.Equal(144D, secondOutline.DestinationTop);
        Assert.False(logical.HasReadableOpenAction);
        Assert.Null(logical.OpenAction);

        PdfLogicalDocument thirdPage = PdfLogicalDocument.LoadPageRanges(pdf, PdfPageRange.From(3, 3));

        PdfNamedDestination thirdDestination = Assert.Single(thirdPage.NamedDestinations);
        Assert.Equal("Third", thirdDestination.Name);
        Assert.Equal(3, thirdDestination.PageNumber);
        PdfOutlineItem thirdOutline = Assert.Single(thirdPage.Outlines);
        Assert.Equal("Third outline", thirdOutline.Title);
        Assert.Equal(3, thirdOutline.PageNumber);
        Assert.True(thirdPage.HasReadableOpenAction);
        Assert.Equal(3, thirdPage.OpenAction!.PageNumber);

        PdfLogicalDocument full = PdfLogicalDocument.Load(pdf);

        Assert.Equal(3, full.NamedDestinations.Count);
        Assert.Equal(3, full.Outlines.Count);
        Assert.Equal(3, full.OpenAction!.PageNumber);
    }

    [Fact]
    public void LoadPageRanges_FiltersPageLabelsToSelectedSourcePages() {
        byte[] pdf = BuildThreePageLabelPdf();

        PdfLogicalDocument pageTwo = PdfLogicalDocument.LoadPageRanges(pdf, PdfPageRange.From(2, 2));

        PdfPageLabel inheritedLabel = Assert.Single(pageTwo.PageLabels);
        Assert.Equal(1, inheritedLabel.StartPageIndex);
        Assert.Equal(2, inheritedLabel.StartPageNumber);
        Assert.Equal("D", inheritedLabel.Style);
        Assert.Equal("A-", inheritedLabel.Prefix);
        Assert.Equal(11, inheritedLabel.StartNumber);

        PdfLogicalDocument selected = PdfLogicalDocument.LoadPageRanges(pdf, PdfPageRange.ParseMany("3,1,3"));

        Assert.Equal(new[] { 3, 1, 3 }, selected.Pages.Select(page => page.PageNumber).ToArray());
        Assert.Equal(2, selected.PageLabels.Count);
        Assert.Equal(new[] { 0, 2 }, selected.PageLabels.Select(label => label.StartPageIndex).ToArray());
        Assert.Equal(new[] { "A-", "B-" }, selected.PageLabels.Select(label => label.Prefix).ToArray());
        Assert.Equal(new[] { 10, 3 }, selected.PageLabels.Select(label => label.StartNumber!.Value).ToArray());

        PdfLogicalDocument full = PdfLogicalDocument.Load(pdf);

        Assert.Equal(2, full.PageLabels.Count);
        Assert.Equal(new[] { 0, 2 }, full.PageLabels.Select(label => label.StartPageIndex).ToArray());
        Assert.Equal(new[] { 10, 3 }, full.PageLabels.Select(label => label.StartNumber!.Value).ToArray());
    }

    [Fact]
    public void LoadPageRanges_ReadsPathAndStreamFromCurrentPosition() {
        byte[] pdf = BuildThreePageLogicalPdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-logical-ranges-" + Guid.NewGuid().ToString("N") + ".pdf");
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");

        try {
            File.WriteAllBytes(path, pdf);

            PdfLogicalDocument fromPath = PdfLogicalDocument.LoadPageRanges(path, PdfPageRange.From(2, 2));
            using var stream = new MemoryStream(prefix.Concat(pdf).ToArray());
            stream.Position = prefix.Length;
            PdfLogicalDocument fromStream = PdfLogicalDocument.LoadPageRanges(stream, PdfPageRange.From(1, 1));

            Assert.Equal(2, Assert.Single(fromPath.Pages).PageNumber);
            Assert.Contains(fromPath.TextBlocks, block => block.Text.Contains("Second logical page", StringComparison.Ordinal));
            Assert.Equal(1, Assert.Single(fromStream.Pages).PageNumber);
            Assert.Contains(fromStream.TextBlocks, block => block.Text.Contains("First logical page", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void LoadPageRanges_RejectsInvalidInputs() {
        byte[] pdf = BuildThreePageLogicalPdf();

        Assert.Throws<ArgumentNullException>(() => PdfLogicalDocument.LoadPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfLogicalDocument.LoadPageRanges((string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfLogicalDocument.LoadPageRanges(" ", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfLogicalDocument.LoadPageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfLogicalDocument.LoadPageRanges(pdf, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfLogicalDocument.LoadPageRanges(pdf));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfLogicalDocument.LoadPageRanges(pdf, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfLogicalDocument.LoadPageRanges(pdf, PdfPageRange.From(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfLogicalDocument.FromPageRanges((PdfReadDocument)null!, PdfPageRange.From(1, 1)));
    }
}
