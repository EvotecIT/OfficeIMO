using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false, false, false, 3)]
    [InlineData(true, true, false, 7)]
    [InlineData(true, true, true, 5)]
    [InlineData(false, false, true, 3)]
    public void Pages_reader_and_report_include_only_selected_header_footer_content(
        bool first, bool even, bool hideFirst, int reconstructedItems) {
        using MemoryStream package = CreatePagesPackageWithHeaderFooterVariants(first, even, hideFirst);
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(package);
        IWorkPagesSection section = Assert.Single(converted.Projection.Sections);
        Assert.Equal(3, section.HeaderContents.Count);
        Assert.Equal(3, section.FooterContents.Count);
        Assert.Equal(reconstructedItems, converted.Report.ReconstructedItemCount);
        var expected = new List<ulong> { 1, 2, 24, 25 };
        if (first && !hideFirst) expected.AddRange(new ulong[] { 20, 21 });
        if (even) expected.AddRange(new ulong[] { 22, 23 });
        Assert.Equal(expected.OrderBy(id => id), converted.Report.SourceUnits
            .Select(unit => unit.Identity.RecordIdentifier).OrderBy(id => id));
        package.Position = 0;
        OfficeDocumentReadResult read = IWorkReaderAdapter.ReadDocument(package, "selection.pages",
            new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);
        string markdown = Assert.IsType<string>(read.Markdown);
        Assert.Contains("Default header", markdown);
        Assert.Contains("Default footer", markdown);
        Assert.Equal(first && !hideFirst, markdown.Contains("First header"));
        Assert.Equal(first && !hideFirst, markdown.Contains("First footer"));
        Assert.Equal(even, markdown.Contains("Even header"));
        Assert.Equal(even, markdown.Contains("Even footer"));
        Assert.Equal((reconstructedItems - 1) / 2, section.SelectedHeaderContents.Count);
        Assert.Equal((reconstructedItems - 1) / 2, section.SelectedFooterContents.Count);
    }
}
