using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using OfficeIMO.Rtf.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfSectionHeaderFooterRegressionTests {
    private const string Input = @"{\rtf1\ansi\sectd{\header H1\par}{\footer F1\par}First\par\sect\sectd{\header H2\par}Second\par\sect\sectd Third\par\sect\sectd{\header}Fourth\par}";

    [Fact]
    public void Word_Bridge_Does_Not_Create_Empty_First_Or_Even_Stories_That_Clear_Inheritance() {
        const string input = @"{\rtf1\ansi\facingp\sectd\titlepg{\headerf FirstHeader\par}{\headerl EvenHeader\par}{\footerf FirstFooter\par}{\footerl EvenFooter\par}One\par\sect\sectd\titlepg Two\par}";
        RtfDocument source = RtfDocument.Read(input).Document;
        using WordDocument word = source.ToWordDocument();
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        RtfDocument result = reopened.ToRtfDocument();
        Assert.Equal(2, result.Sections.Count);
        Assert.Equal(4, result.Sections[0].HeaderFooters.Count);
        Assert.Empty(result.Sections[1].HeaderFooters);
        Assert.True(result.Sections[1].PageSetup.DifferentFirstPageHeaderFooter);
        Assert.True(result.Settings.FacingPages);
        Assert.Equal(new[] { "FirstHeader", "EvenHeader", "FirstFooter", "EvenFooter" }.OrderBy(text => text),
            result.GetEffectiveHeaderFooters(result.Sections[1]).Select(item => item.ToPlainText()).OrderBy(text => text));
    }

    [Fact]
    public void Pdf_Pages_Use_Each_Sections_Effective_Header_And_Inherited_Footer() {
        RtfDocument source = RtfDocument.Read(Input).Document;
        using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(source.ToPdfBytes());
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("H1", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("H2", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("H2", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("H2", pdf.GetPage(3).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("H1", pdf.GetPage(4).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("H2", pdf.GetPage(4).Text, StringComparison.Ordinal);
        Assert.All(pdf.GetPages(), page => Assert.Contains("F1", page.Text, StringComparison.Ordinal));
    }

    [Fact]
    public void Word_Bridge_Preserves_Default_Header_Overrides_And_Footer_Inheritance() {
        RtfDocument source = RtfDocument.Read(Input).Document;
        using WordDocument word = source.ToWordDocument();
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        AssertSections(reopened.ToRtfDocument());
    }

    [Fact]
    public void Parse_Save_And_Clone_Retain_Section_Overrides_Inheritance_And_Explicit_Empty_Headers() {
        RtfDocument source = RtfDocument.Read(Input).Document;
        AssertSections(source);
        AssertSections(RtfDocument.Read(source.ToRtf()).Document);
        RtfDocument clone = source.Clone();
        AssertSections(clone);
        Assert.Same(clone.HeaderFooters[2], clone.Sections[1].HeaderFooters[0]);
        clone.Sections[1].HeaderFooters[0].Paragraphs[0].Runs[0].Text = "Changed";
        Assert.Equal("H2", source.Sections[1].HeaderFooters[0].ToPlainText());
    }

    [Fact]
    public void Section_Builders_And_Document_AddHeader_Attach_To_The_Selected_Section() {
        RtfDocument document = RtfDocument.Create();
        RtfSection first = document.AddSection();
        first.AddHeader().AddParagraph("H1");
        first.AddParagraph("One");
        RtfSection second = document.AddSection();
        document.AddHeader().AddParagraph("H2");
        second.AddFooter().AddParagraph("F2");
        second.AddParagraph("Two");
        Assert.Equal("H1", Assert.Single(first.HeaderFooters).ToPlainText());
        Assert.Equal(2, second.HeaderFooters.Count);
        RtfDocument reopened = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal("H2", reopened.GetEffectiveHeaderFooters(reopened.Sections[1]).Single(item => item.Kind == RtfHeaderFooterKind.Header).ToPlainText());
    }

    [Fact]
    public void Html_Metadata_RoundTrip_Preserves_Section_Ownership_And_Empty_Overrides() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        string html = document.ToHtml(new RtfToHtmlOptions { IncludeRoundTripMetadata = true, FragmentOnly = false });
        AssertSections(HtmlConversionDocument.Parse(html).ToRtfDocument());
    }

    private static void AssertSections(RtfDocument document) {
        Assert.Equal(4, document.Sections.Count);
        Assert.Equal(new[] { 2, 1, 0, 1 }, document.Sections.Select(section => section.HeaderFooters.Count));
        Assert.Equal(new[] { "H1", "H2", "H2", "" }, document.Sections.Select(section =>
            document.GetEffectiveHeaderFooters(section).Single(item => item.Kind == RtfHeaderFooterKind.Header).ToPlainText()));
        Assert.All(document.Sections, section => Assert.Equal("F1",
            document.GetEffectiveHeaderFooters(section).Single(item => item.Kind == RtfHeaderFooterKind.Footer).ToPlainText()));
        Assert.Equal(new[] { "First", "Second", "Third", "Fourth" }, document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
    }
}
