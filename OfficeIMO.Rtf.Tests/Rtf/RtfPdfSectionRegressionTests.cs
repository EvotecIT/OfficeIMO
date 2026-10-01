using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfSectionRegressionTests {
    [Fact]
    public void AbsentRtfStoriesPreserveCallerRunningContentAcrossFacingPageSections() {
        RtfDocument document = RtfDocument.Create();
        document.Settings.FacingPages = true;
        document.AddSection().AddParagraph("First body");
        document.AddSection().AddParagraph("Second body");
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes(new RtfToPdfOptions {
            PdfOptions = new PdfCore.PdfOptions { ShowPageNumbers = true, FooterFormat = "Caller page {page}" }
        }));
        Assert.Equal(2, read.Pages.Count);
        Assert.Contains("Caller page 1", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Caller page 2", read.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void SectionsWithoutEffectiveStoriesRetainCallerFirstAndEvenVariants() {
        RtfDocument document = RtfDocument.Create();
        document.Settings.FacingPages = true;
        RtfSection first = document.AddSection();
        first.AddParagraph("First body").AddPageBreak();
        first.AddParagraph("Second body");
        RtfSection second = document.AddSection();
        second.AddParagraph("Third body").AddPageBreak();
        second.AddParagraph("Fourth body");
        RtfSection third = document.AddSection();
        third.AddHeader().AddParagraph("RTF third header");
        third.AddParagraph("Fifth body");
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes(new RtfToPdfOptions {
            PdfOptions = new PdfCore.PdfOptions {
                ShowHeader = true,
                ShowPageNumbers = true,
                HeaderFormat = "Caller default header",
                FooterFormat = "Caller default footer {page}",
                DifferentFirstPageHeaderFooter = true,
                FirstPageHeaderFormat = "Caller first header",
                FirstPageFooterFormat = "Caller first footer {page}",
                DifferentOddAndEvenPagesHeaderFooter = true,
                EvenPageHeaderFormat = "Caller even header",
                EvenPageFooterFormat = "Caller even footer {page}"
            }
        }));
        Assert.Equal(5, read.Pages.Count);
        foreach (int index in new[] { 0, 2 }) {
            Assert.Contains("Caller first header", read.Pages[index].ExtractText(), StringComparison.Ordinal);
            Assert.Contains("Caller first footer " + (index + 1), read.Pages[index].ExtractText(), StringComparison.Ordinal);
        }
        foreach (int index in new[] { 1, 3 }) {
            Assert.Contains("Caller even header", read.Pages[index].ExtractText(), StringComparison.Ordinal);
            Assert.Contains("Caller even footer " + (index + 1), read.Pages[index].ExtractText(), StringComparison.Ordinal);
        }
        Assert.Contains("RTF third header", read.Pages[4].ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(RtfSectionBreakKind.OddPage, 3)]
    [InlineData(RtfSectionBreakKind.EvenPage, 2)]
    public void SectionParityBreaksInsertOnlyTheRequiredBlankPage(RtfSectionBreakKind kind, int pageCount) {
        RtfDocument document = RtfDocument.Create();
        RtfSection first = document.AddSection();
        first.AddHeader().AddParagraph("First header");
        first.AddParagraph("First body");
        RtfSection second = document.AddSection(kind);
        second.AddHeader().AddParagraph("Second header");
        second.AddParagraph("Second body");

        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(pageCount, read.Pages.Count);
        Assert.Contains("Second header", read.Pages[pageCount - 1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Second body", read.Pages[pageCount - 1].ExtractText(), StringComparison.Ordinal);
        if (pageCount == 3) {
            Assert.True(string.IsNullOrWhiteSpace(read.Pages[1].ExtractText()));
            Assert.Throws<InvalidDataException>(() => document.ToPdfBytes(new RtfToPdfOptions {
                PdfOptions = new PdfCore.PdfOptions { MaxGeneratedPages = 2 }
            }));
        }
    }

    [Fact]
    public void DocumentPageNumberStartDoesNotRestartEachSubsequentSection() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPageNumbering(start: 7);
        document.AddSection().AddParagraph("First body");
        document.AddSection().AddParagraph("Second body");
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes(new RtfToPdfOptions {
            PdfOptions = new PdfCore.PdfOptions { ShowPageNumbers = true, FooterFormat = "{page}/{pages}" }
        }));
        Assert.Equal(2, read.Pages.Count);
        Assert.Contains("7/8", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("8/8", read.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void EvenHeadersFollowVisiblePageNumbersWhileFirstHeadersFollowSectionStarts() {
        RtfDocument document = RtfDocument.Create();
        document.Settings.FacingPages = true;
        RtfSection first = document.AddSection();
        first.AddHeader(RtfHeaderFooterKind.RightHeader).AddParagraph("Odd A");
        first.AddHeader(RtfHeaderFooterKind.LeftHeader).AddParagraph("Even A");
        first.AddParagraph("First body");
        RtfSection second = document.AddSection();
        second.AddHeader(RtfHeaderFooterKind.RightHeader).AddParagraph("Odd B");
        second.AddHeader(RtfHeaderFooterKind.LeftHeader).AddParagraph("Even B");
        second.AddParagraph("Second body");
        RtfSection third = document.AddSection();
        third.PageSetup.SetPageNumbering(start: 10, restart: true);
        third.AddHeader(RtfHeaderFooterKind.RightHeader).AddParagraph("Odd C");
        third.AddHeader(RtfHeaderFooterKind.LeftHeader).AddParagraph("Even C");
        third.AddParagraph("Third body");
        RtfSection fourth = document.AddSection();
        fourth.PageSetup.SetDifferentFirstPageHeaderFooter();
        fourth.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("Title D");
        fourth.AddHeader(RtfHeaderFooterKind.RightHeader).AddParagraph("Odd D");
        fourth.AddHeader(RtfHeaderFooterKind.LeftHeader).AddParagraph("Even D");
        fourth.AddParagraph("Fourth body").AddPageBreak();
        fourth.AddParagraph("Fifth body");

        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(6, read.Pages.Count);
        Assert.Contains("Odd A", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Even B", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.True(string.IsNullOrWhiteSpace(read.Pages[2].ExtractText()));
        Assert.Contains("Even C", read.Pages[3].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Title D", read.Pages[4].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Even D", read.Pages[5].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void LaterSectionsUseDocumentDefaultsAndExplicitPortraitWithoutChangingCallerOptions() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPaperSize(12240, 15840);
        document.PageSetup.SetMargins(1440, 1440, 1440, 1440);
        RtfSection first = document.AddSection();
        first.PageSetup.SetLandscape();
        first.PageSetup.SetMargins(720, 720, 720, 720);
        first.AddParagraph("Landscape body");
        RtfSection second = document.AddSection();
        second.AddParagraph("Default portrait body");
        RtfSection third = document.AddSection();
        third.PageSetup.SetLandscape(false);
        third.AddParagraph("Explicit portrait body");
        var caller = new PdfCore.PdfOptions { PageWidth = 400, PageHeight = 500, MarginLeft = 17 };

        byte[] bytes = document.ToPdfBytes(new RtfToPdfOptions {
            PdfOptions = caller,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(792, info.Pages[0].Width, 2);
        Assert.Equal(612, info.Pages[0].Height, 2);
        Assert.All(info.Pages.Skip(1), page => {
            Assert.Equal(612, page.Width, 2);
            Assert.Equal(792, page.Height, 2);
        });
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(36, read.GetPage(1).Letters[0].StartBaseLine.X, 1);
        Assert.Equal(72, read.GetPage(2).Letters[0].StartBaseLine.X, 1);
        Assert.Equal(400, caller.PageWidth);
        Assert.Equal(500, caller.PageHeight);
        Assert.Equal(17, caller.MarginLeft);
    }

    [Fact]
    public void DeclaredFirstAndEvenStoriesRemainInactiveUntilTheirSelectionFlagsAreEnabled() {
        RtfDocument document = RtfDocument.Create();
        document.AddHeader().AddParagraph("Default header");
        document.AddFooter().AddParagraph("Default footer");
        document.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("Inactive first header");
        document.AddHeader(RtfHeaderFooterKind.LeftHeader).AddParagraph("Inactive even header");
        document.AddFooter(RtfHeaderFooterKind.FirstFooter).AddParagraph("Inactive first footer");
        document.AddFooter(RtfHeaderFooterKind.LeftFooter).AddParagraph("Inactive even footer");
        document.AddParagraph("First body").AddPageBreak();
        document.AddParagraph("Second body");

        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(2, read.Pages.Count);
        Assert.All(read.Pages, page => {
            string text = page.ExtractText();
            Assert.Contains("Default header", text, StringComparison.Ordinal);
            Assert.Contains("Default footer", text, StringComparison.Ordinal);
            Assert.DoesNotContain("Inactive", text, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void FirstPageSelectionIsScopedToEachSectionAndCanBeExplicitlyDisabled() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetDifferentFirstPageHeaderFooter();
        RtfSection first = document.AddSection();
        first.AddHeader().AddParagraph("First default header");
        first.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("First title header");
        first.AddParagraph("First body");
        RtfSection second = document.AddSection();
        second.PageSetup.SetDifferentFirstPageHeaderFooter(false);
        second.AddHeader().AddParagraph("Second default header");
        second.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("Second inactive title");
        second.AddParagraph("Second body");
        RtfSection third = document.AddSection();
        third.AddHeader().AddParagraph("Third default header");
        third.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("Third title header");
        third.AddParagraph("Third body").AddPageBreak();
        third.AddParagraph("Fourth body");

        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(4, read.Pages.Count);
        Assert.Contains("First title header", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Second default header", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Second inactive title", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Third title header", read.Pages[2].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Third default header", read.Pages[3].ExtractText(), StringComparison.Ordinal);
    }
}
