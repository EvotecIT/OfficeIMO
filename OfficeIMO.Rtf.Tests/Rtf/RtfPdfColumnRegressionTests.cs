using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfColumnRegressionTests {
    [Fact]
    public void ExplicitPageBreakDoesNotBalanceThePrecedingPageAcrossColumns() {
        RtfDocument document = RtfDocument.Read(@"{\rtf1\ansi\paperw12240\paperh15840\margl1440\margr1440
\sectd\cols2\colsx480\pard Before01\line Before02\line Before03\line Before04\line Before05\line Before06\line Before07\line Before08\page After marker\par}").Document;
        using var read = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(PortableOptions()));
        Assert.Equal(2, read.NumberOfPages);
        var before = read.GetPage(1).GetWords().Where(word => word.Text.StartsWith("Before", StringComparison.Ordinal)).ToArray();
        Assert.Equal(8, before.Length);
        Assert.All(before, word => Assert.Equal(72, word.BoundingBox.Left, 1));
        Assert.Contains("After marker", read.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ColumnStartingSectionsReportTheNewPageFallback() {
        RtfDocument document = RtfDocument.Create();
        RtfSection first = document.AddSection();
        first.ColumnCount = 2;
        first.AddParagraph("First section marker");
        RtfSection second = document.AddSection(RtfSectionBreakKind.Column);
        second.ColumnCount = 2;
        second.AddParagraph("Column section marker");
        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(PortableOptions());
        var warning = Assert.Single(result.Warnings, item => item.Code == "ColumnSectionBreakFlattened");
        Assert.Equal(nameof(RtfConversionAction.Flattened), warning.Details["RtfAction"]);
        var read = PdfCore.PdfReadDocument.Open(result.ToBytes());
        Assert.Equal(2, read.Pages.Count);
        Assert.Contains("First section marker", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Column section marker", read.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void EqualSectionColumnsHonorInlineColumnAndPageBreaks() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPaperSize(12240, 15840);
        document.PageSetup.SetMargins(1440, 1440, 1440, 1440);
        RtfSection section = document.AddSection();
        section.ColumnCount = 2;
        section.ColumnSpaceTwips = 480;
        RtfParagraph paragraph = section.AddParagraph("Left marker");
        paragraph.AddColumnBreak();
        paragraph.AddText("Right marker");
        paragraph.AddPageBreak();
        paragraph.AddText("Next left marker");
        paragraph.AddColumnBreak();
        paragraph.AddText("Next right marker");
        byte[] bytes = document.ToPdfBytes(PortableOptions());
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, read.NumberOfPages);
        foreach (int pageNumber in new[] { 1, 2 }) {
            var page = read.GetPage(pageNumber);
            string left = pageNumber == 1 ? "Left" : "Next left";
            string right = pageNumber == 1 ? "Right" : "Next right";
            var words = page.GetWords().ToArray();
            Assert.Equal(72, words.First(word => word.Text == (pageNumber == 1 ? "Left" : "Next")).BoundingBox.Left, 1);
            Assert.Equal(318, words.First(word => word.BoundingBox.Left > 300).BoundingBox.Left, 1);
            Assert.Contains(left, page.Text, StringComparison.Ordinal);
            Assert.Contains(right, page.Text, StringComparison.Ordinal);
        }
        Assert.Throws<InvalidDataException>(() => document.ToPdfBytes(new RtfToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new PdfCore.PdfOptions { MaxGeneratedPages = 1 }
        }));
    }

    [Fact]
    public void ContinuousColumnsFlowOnTheCurrentPageAndReturnToSingleColumn() {
        RtfDocument document = RtfDocument.Create();
        document.AddSection().AddParagraph("Full width before");
        RtfSection columns = document.AddSection(RtfSectionBreakKind.Continuous);
        columns.ColumnCount = 2;
        columns.ColumnSpaceTwips = 480;
        columns.AddParagraph("Left body").AddColumnBreak();
        columns.AddParagraph("Right body");
        document.AddSection(RtfSectionBreakKind.Continuous).AddParagraph("Full width after");
        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(PortableOptions());
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
        byte[] bytes = result.ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        var words = read.GetPage(1).GetWords().ToArray();
        Assert.Equal(318, words.Single(word => word.Text == "Right").BoundingBox.Left, 1);
        Assert.All(words.Where(word => word.Text == "Full"), word => Assert.Equal(72, word.BoundingBox.Left, 1));
        Assert.True(words.Last(word => word.Text == "Full").BoundingBox.Bottom < words.Single(word => word.Text == "Right").BoundingBox.Bottom);
    }

    [Fact]
    public void ColumnBreakInASingleColumnSectionAdvancesToTheNextPage() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before marker");
        paragraph.AddColumnBreak();
        paragraph.AddText("After marker");
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes(PortableOptions()));
        Assert.Equal(2, read.Pages.Count);
        Assert.Contains("Before marker", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("After marker", read.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void UnsupportedColumnGeometryReportsFlatteningAndKeepsTheBody() {
        RtfDocument document = RtfDocument.Create();
        RtfSection section = document.AddSection();
        section.ColumnCount = 2;
        section.AddColumn(3000, 480);
        section.AddColumn(5880, 0);
        section.AddParagraph("Unequal body marker");
        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(PortableOptions());
        var warning = Assert.Single(result.Warnings, item => item.Code == "SectionColumnsFlattened");
        Assert.Equal(nameof(RtfConversionAction.Flattened), warning.Details["RtfAction"]);
        Assert.Contains("Unequal body marker", PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void ContinuousPageGeometryAndStoriesReportTheSettingsThatAreNotApplied() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPaperSize(12240, 15840);
        document.AddSection().AddParagraph("Before body");
        RtfSection section = document.AddSection(RtfSectionBreakKind.Continuous);
        section.PageSetup.SetLandscape();
        section.AddHeader().AddParagraph("Deferred header");
        section.AddParagraph("Continuous body");
        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(PortableOptions());
        var warning = Assert.Single(result.Warnings, item => item.Code == "ContinuousSectionPageSettingsFlattened");
        Assert.Equal(nameof(RtfConversionAction.Flattened), warning.Details["RtfAction"]);
        Assert.Contains("PaperSize", warning.Details["Properties"], StringComparison.Ordinal);
        Assert.Contains("Headers", warning.Details["Properties"], StringComparison.Ordinal);
        var read = PdfCore.PdfReadDocument.Open(result.ToBytes());
        Assert.Single(read.Pages);
        Assert.Contains("Continuous body", read.ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Deferred header", read.ExtractText(), StringComparison.Ordinal);
    }

    private static RtfToPdfOptions PortableOptions() => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
    };
}
