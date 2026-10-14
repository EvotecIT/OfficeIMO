using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.IWork;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesHeaderFooterWritingTests {
    [Fact]
    public void Default_page_text_is_deterministic_and_retains_styles_and_distances_in_strict_saved_Docx() {
        IWorkPagesDocument document = Example();
        byte[] bytes = document.SaveBytes(); Assert.Equal(bytes, Example().SaveBytes());
        Assert.Equal(File.ReadAllBytes(Fixture("pages-created-headers-v15.4.pages")), bytes);
        IWorkPagesProjection projection = IWorkSourceDocument.Open(new MemoryStream(bytes)).ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        IWorkPagesSection section = Assert.Single(projection.Sections);
        Assert.False(section.DifferentFirstPage); Assert.False(section.DifferentOddAndEvenPages);
        Assert.Equal("OfficeIMO — Pages header", Assert.Single(section.DefaultPageHeaderContents).PlainText);
        Assert.Equal("Zażółć café — footer", Assert.Single(section.DefaultPageFooterContents).PlainText);
        Assert.Equal(30, projection.PageLayout!.HeaderMarginPoints); Assert.Equal(36, projection.PageLayout.FooterMarginPoints);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
        W.Paragraph header = Assert.Single(Assert.Single(saved.MainDocumentPart!.HeaderParts).Header!.Elements<W.Paragraph>());
        W.Paragraph footer = Assert.Single(Assert.Single(saved.MainDocumentPart.FooterParts).Footer!.Elements<W.Paragraph>());
        Assert.Equal(document.Header!.Text, header.InnerText); Assert.Equal(document.Footer!.Text, footer.InnerText);
        W.Run headerRun = Assert.Single(header.Elements<W.Run>(), run => run.InnerText.Length > 0);
        W.Run footerRun = Assert.Single(footer.Elements<W.Run>(), run => run.InnerText.Length > 0);
        Assert.True(headerRun.RunProperties!.Bold!.Val?.Value ?? true);
        Assert.Equal("28", headerRun.RunProperties.FontSize!.Val!.Value); Assert.Equal("003366", headerRun.RunProperties.Color!.Val!.Value);
        Assert.True(footerRun.RunProperties!.Italic!.Val?.Value ?? true);
        W.PageMargin margin = saved.MainDocumentPart.Document!.Body!.Elements<W.SectionProperties>().Single().GetFirstChild<W.PageMargin>()!;
        Assert.Equal(600u, margin.Header!.Value); Assert.Equal(720u, margin.Footer!.Value);
    }

    [Fact]
    public void Native_saved_page_text_retains_edits_and_exports_the_same_text_and_distances_as_Apple() {
        string source = Fixture("pages-created-headers-edited-v15.4.pages");
        IWorkPagesProjection projection = IWorkSourceDocument.Open(source).ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        IWorkPagesSection section = Assert.Single(projection.Sections);
        Assert.Equal("OfficeIMO — Native header", Assert.Single(section.DefaultPageHeaderContents).PlainText);
        Assert.Equal("Zażółć café — native footer", Assert.Single(section.DefaultPageFooterContents).PlainText);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source, conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        using var native = WordprocessingDocument.Open(Fixture("pages-created-headers-edited-v15.4.docx"), false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        Assert.Equal(Assert.Single(native.MainDocumentPart!.HeaderParts).Header!.InnerText,
            Assert.Single(saved.MainDocumentPart!.HeaderParts).Header!.InnerText);
        Assert.Equal(Assert.Single(native.MainDocumentPart.FooterParts).Footer!.InnerText,
            Assert.Single(saved.MainDocumentPart.FooterParts).Footer!.InnerText);
        Assert.Equal(native.MainDocumentPart.Document!.Body!.Elements<W.Paragraph>().Select(paragraph => paragraph.InnerText),
            saved.MainDocumentPart.Document!.Body!.Elements<W.Paragraph>().Take(73).Select(paragraph => paragraph.InnerText));
        W.Run header = Assert.Single(saved.MainDocumentPart.HeaderParts.Single().Header!.Descendants<W.Run>(), run => run.InnerText.Length > 0);
        W.Run footer = Assert.Single(saved.MainDocumentPart.FooterParts.Single().Footer!.Descendants<W.Run>(), run => run.InnerText.Length > 0);
        Assert.True(header.RunProperties!.Bold!.Val?.Value ?? true); Assert.True(footer.RunProperties!.Italic!.Val?.Value ?? true);
        W.PageMargin nativeMargin = native.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Single().GetFirstChild<W.PageMargin>()!;
        W.PageMargin savedMargin = saved.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Single().GetFirstChild<W.PageMargin>()!;
        Assert.Equal(nativeMargin.Header!.Value, savedMargin.Header!.Value); Assert.Equal(nativeMargin.Footer!.Value, savedMargin.Footer!.Value);
    }

    [Theory]
    [InlineData(true, false, "Header")]
    [InlineData(false, true, "Footer")]
    [InlineData(true, true, "Both")]
    public void A_single_enabled_area_or_both_areas_have_complete_readback(bool header, bool footer, string text) {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Body");
        if (header) document.SetHeader(text, distancePoints: 0);
        if (footer) document.SetFooter(text, distancePoints: 0);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(document.SaveBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        IWorkPagesProjection projection = IWorkSourceDocument.Open(new MemoryStream(document.SaveBytes())).ReadPages();
        Assert.Equal(header ? text : "", string.Concat(projection.Sections.Single().SelectedHeaderContents.Select(content => content.PlainText)));
        Assert.Equal(footer ? text : "", string.Concat(projection.Sections.Single().SelectedFooterContents.Select(content => content.PlainText)));
    }

    [Fact]
    public void Invalid_page_text_updates_are_atomic_and_clearing_restores_the_original_creation_graph() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Body"); byte[] original = document.SaveBytes();
        document.SetHeader("Header"); document.SetFooter("Footer"); byte[] authored = document.SaveBytes();
        foreach (float distance in new[] { -1f, 72f, float.NaN, float.PositiveInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => document.SetHeader("Rejected", distancePoints: distance));
            Assert.Throws<ArgumentOutOfRangeException>(() => document.SetFooter("Rejected", distancePoints: distance));
        }
        Assert.Throws<ArgumentNullException>(() => document.SetHeader((IWorkTextRunData)null!));
        foreach (string text in new[] { "", "two\nlines", "two\r\nlines", "two\u2028lines", "two\u2029lines", "a\tb" }) {
            Assert.Throws<ArgumentException>(() => document.SetHeader(text));
            Assert.Throws<ArgumentException>(() => document.SetFooter(text));
        }
        var emptyStyledRun = new IWorkTextRunData("", "Times New Roman", 24, "003366", bold: true);
        Assert.Throws<ArgumentException>(() => document.SetHeader(emptyStyledRun));
        Assert.Throws<ArgumentException>(() => document.SetFooter(emptyStyledRun));
        Assert.Equal(authored, document.SaveBytes());
        document.ClearHeader(); document.ClearFooter();
        Assert.Null(document.Header); Assert.Null(document.Footer);
        Assert.Equal(0, document.HeaderDistancePoints); Assert.Equal(0, document.FooterDistancePoints);
        Assert.Equal(original, document.SaveBytes());
    }

    [Theory]
    [InlineData("paragraphs")]
    [InlineData("runs")]
    [InlineData("characters")]
    public void Page_text_shares_creation_budgets_before_mutating_output(string budget) {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("B"); document.SetHeader("H"); document.SetFooter("F");
        var limits = new IWorkPagesWriteOptions { MaximumParagraphs = 3, MaximumTextRuns = 3, MaximumTextCharacters = 4 };
        byte[] expected = document.SaveBytes(limits);
        switch (budget) {
            case "paragraphs": limits.MaximumParagraphs = 2; break;
            case "runs": limits.MaximumTextRuns = 2; break;
            case "characters": limits.MaximumTextCharacters = 3; break;
        }
        using var output = new MemoryStream(new byte[] { 9, 8, 7 });
        Assert.Throws<InvalidDataException>(() => document.Save(output, limits)); Assert.Equal(new byte[] { 9, 8, 7 }, output.ToArray());
        Assert.Equal(expected, document.SaveBytes());
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.SetHeader(new IWorkTextRunData("OfficeIMO — Pages header", "Arial", 14, "003366", bold: true), 30);
        document.SetFooter(new IWorkTextRunData("Zażółć café — footer", "Arial", 12, "555555", italic: true), 36);
        document.AddParagraph("Native header and footer creation", fontSizePoints: 20);
        for (int index = 1; index <= 72; index++) document.AddParagraph($"Body line {index:D2} — header/footer creation");
        return document;
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
