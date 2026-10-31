using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "default", "next")]
    [InlineData(false, "first", "next")]
    [InlineData(false, "even", "next")]
    [InlineData(true, "default", "next")]
    [InlineData(true, "first", "next")]
    [InlineData(true, "even", "next")]
    [InlineData(false, "default", "fields")]
    [InlineData(false, "first", "fields")]
    [InlineData(false, "even", "fields")]
    [InlineData(true, "default", "fields")]
    [InlineData(true, "first", "fields")]
    [InlineData(true, "even", "fields")]
    [InlineData(false, "default", "blank")]
    [InlineData(false, "first", "blank")]
    [InlineData(false, "even", "blank")]
    [InlineData(true, "default", "blank")]
    [InlineData(true, "first", "blank")]
    [InlineData(true, "even", "blank")]
    [InlineData(false, "default", "continuous")]
    [InlineData(false, "first", "continuous")]
    [InlineData(false, "even", "continuous")]
    [InlineData(true, "default", "continuous")]
    [InlineData(true, "first", "continuous")]
    [InlineData(true, "even", "continuous")]
    [InlineData(false, "default", "override")]
    [InlineData(false, "first", "override")]
    [InlineData(false, "even", "override")]
    [InlineData(true, "default", "override")]
    [InlineData(true, "first", "override")]
    [InlineData(true, "even", "override")]
    public void SaveAsPdf_InheritedRunningStoriesPreserveRolesFieldsAndContinuousPages(bool footer, string kind, string mode) {
        using WordDocument document = CreateInheritedRunningStoryControl(footer, kind, mode);
        WordSection later = document.Sections[1];
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        WordPdfExportOptions options = EmptyRunningOptions();
        options.IncludePageNumbers = false;
        if (mode == "override") options.PdfOptions = new PdfCreateOptions {
            HeaderFont = PdfStandardFont.Courier, HeaderFontSize = global::OfficeIMO.OfficeLength.Points(18D), HeaderAlign = PdfAlign.Left,
            FooterFont = PdfStandardFont.Courier, FooterFontSize = global::OfficeIMO.OfficeLength.Points(18D), FooterAlign = PdfAlign.Left
        };
        using WordDocument? explicitDocument = mode == "override" ? CreateExplicitRunningStoryControl(document) : null;
        using var explicitPdf = explicitDocument == null ? null : PdfPigDocument.Open(explicitDocument.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        for (int repeat = 0; repeat < 2; repeat++) {
            using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
            Assert.Equal(mode == "continuous" ? 1 : 3, pdf.NumberOfPages);
            Assert.Contains("S1", pdf.GetPage(1).Text);
            Assert.Contains("S2A", pdf.GetPage(mode == "continuous" ? 1 : 2).Text);
            if (mode != "continuous") Assert.Contains("S2B", pdf.GetPage(3).Text);
            for (int page = 1; page <= pdf.NumberOfPages; page++) {
                bool marker = kind == "default" || (kind == "first" ? page <= 2 : page == 2);
                if (mode == "blank" && page > 1) marker = false;
                string text = pdf.GetPage(page).Text;
                Assert.Equal(marker, text.Contains("MARK"));
                if (marker && mode == "fields") Assert.Contains($"MARK P{page} S{(page == 1 ? 1 : 2)}", text);
                if (explicitPdf != null)
                    Assert.Equal(explicitPdf.GetPage(page).Letters.Select(letter =>
                            (letter.Value, letter.StartBaseLine, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())),
                        pdf.GetPage(page).Letters.Select(letter =>
                            (letter.Value, letter.StartBaseLine, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())));
            }
        }
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        // Resolving inherited content must preserve authored linkage slots.
        if (mode != "blank") {
            Assert.Null(later.Header?.Default); Assert.Null(later.Footer?.Default);
            Assert.Null(later.Header?.First); Assert.Null(later.Footer?.First);
            Assert.Null(later.Header?.Even); Assert.Null(later.Footer?.Even);
        }
    }

    private static WordDocument CreateExplicitRunningStoryControl(WordDocument source) {
        using WordDocument explicitSource = WordDocument.Load(new MemoryStream(source.ToBytes()));
        foreach (var reference in explicitSource.Sections[0]._sectionProperties.ChildElements
            .Where(element => element is W.HeaderReference || element is W.FooterReference))
            explicitSource.Sections[1]._sectionProperties.AddChild(reference.CloneNode(true), true);
        return WordDocument.Load(new MemoryStream(explicitSource.ToBytes()));
    }

    [Theory]
    [InlineData(false, 5)]
    [InlineData(false, 6)]
    [InlineData(true, 5)]
    [InlineData(true, 6)]
    public void SaveAsPdf_InheritedEvenRunningStoryUsesRestartedPageNumber(bool footer, int start) {
        using WordDocument document = CreateInheritedRunningStoryControl(footer, "even", "fields");
        document.Sections[1]._sectionProperties.AddChild(new W.PageNumberType { Start = start }, true);
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(start % 2 == 0 ? 3 : 4, pdf.NumberOfPages);
        int secondPage = pdf.NumberOfPages - 1;
        for (int index = 0; index < 2; index++) {
            string text = pdf.GetPage(secondPage + index).Text;
            bool hasMarker = (start + index) % 2 == 0;
            Assert.Equal(hasMarker, text.Contains("MARK"));
            if (hasMarker) Assert.Contains($"MARK P{start + index} S2", text);
        }
    }

    private static WordDocument CreateInheritedRunningStoryControl(bool footer, string kind, string mode) {
        using WordDocument source = WordDocument.Create();
        WordHeaderFooter[] headers = { source.HeaderDefaultOrCreate, source.HeaderFirstOrCreate, source.HeaderEvenOrCreate };
        WordHeaderFooter[] footers = { source.FooterDefaultOrCreate, source.FooterFirstOrCreate, source.FooterEvenOrCreate };
        foreach (WordHeaderFooter story in headers.Concat(footers))
            foreach (var child in story.ChildElements.ToArray()) child.Remove();
        source.Sections[0].DifferentFirstPage = kind == "first";
        source.Sections[0].DifferentOddAndEvenPages = kind == "even";
        int slot = Array.IndexOf(new[] { "default", "first", "even" }, kind);
        WordParagraph marker = (footer ? footers : headers)[slot].AddParagraph("MARK");
        SetInheritedRunningControlStyle(marker);
        if (mode == "fields") marker.AddText(" P").AddField(WordFieldType.Page).Paragraph.AddText(" S").AddField(WordFieldType.SectionPages);
        SetInheritedRunningControlStyle(source.AddParagraph("S1"));
        WordSection later = source.AddSection(mode == "continuous" ? WordSectionBreakType.Continuous : WordSectionBreakType.NextPage);
        later.DifferentFirstPage = kind == "first";
        if (mode == "blank") {
            later.AddHeadersAndFooters();
            if (kind == "even") later.DifferentOddAndEvenPages = true;
            WordHeaderFooter? blank = footer
                ? kind == "first" ? later.Footer?.First : kind == "even" ? later.Footer?.Even : later.Footer?.Default
                : kind == "first" ? later.Header?.First : kind == "even" ? later.Header?.Even : later.Header?.Default;
            Assert.NotNull(blank);
            foreach (var child in blank!.ChildElements.ToArray()) child.Remove();
            blank.AddParagraph();
        }
        WordParagraph body = source.AddParagraph("S2A"); SetInheritedRunningControlStyle(body);
        if (mode != "continuous") {
            body.AddBreak(WordBreakType.Page); SetInheritedRunningControlStyle(source.AddParagraph("S2B"));
        }
        W.HeaderFooterValues role = kind == "first" ? W.HeaderFooterValues.First
            : kind == "even" ? W.HeaderFooterValues.Even : W.HeaderFooterValues.Default;
        foreach (WordSection section in source.Sections) {
            foreach (W.HeaderReference reference in section._sectionProperties.Elements<W.HeaderReference>().ToArray())
                if (footer || reference.Type?.Value != role || (section == later && mode != "blank")) reference.Remove();
            foreach (W.FooterReference reference in section._sectionProperties.Elements<W.FooterReference>().ToArray())
                if (!footer || reference.Type?.Value != role || (section == later && mode != "blank")) reference.Remove();
        }
        return WordDocument.Load(new MemoryStream(source.ToBytes()));
    }

    private static void SetInheritedRunningControlStyle(WordParagraph paragraph) {
        paragraph.FontFamily = "Arial"; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(240) * 635)));
        paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
    }
}
