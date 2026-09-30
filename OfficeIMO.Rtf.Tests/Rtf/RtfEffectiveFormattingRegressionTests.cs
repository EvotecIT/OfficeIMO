using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfEffectiveFormattingRegressionTests {
    private const string Input = @"{\rtf1\ansi{\stylesheet{\s1\b\qc\li720\keepn Base;}{\s2\sbasedon1\i Derived;}{\*\cs1\ul Character;}}\pard\s2 Inherited {\b0\i0\ulnone Cleared} {\cs1 Character}\par\pard\s2\ql\keepn0\li0 Explicit\par}";

    [Fact]
    public void Word_Style_Colors_Are_Schema_Valid_And_Preserve_Effective_Formatting() {
        RtfDocument source = RtfDocument.Create();
        RtfStyle style = source.AddStyle(1, "Colored");
        style.ForegroundColorIndex = source.AddColor(0x12, 0x34, 0x56);
        style.HighlightColorIndex = source.AddColor(0xFF, 0xFF, 0x00);
        source.AddParagraph("Inherited").StyleId = style.Id;
        using WordDocument word = source.ToWordDocument();
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word._wordprocessingDocument));
        RtfDocument reopened = word.ToRtfDocument();
        RtfParagraph paragraph = Assert.Single(reopened.Paragraphs);
        RtfRun run = reopened.ResolveRunFormatting(paragraph, Assert.Single(paragraph.Runs));
        Assert.Equal("#123456", reopened.GetColor(run.ForegroundColorIndex!.Value)!.ToString());
        Assert.Equal("#FFFF00", reopened.GetColor(run.HighlightColorIndex!.Value)!.ToString());
        RtfStyle imported = reopened.Styles.Single(item => item.Id == paragraph.StyleId);
        Assert.Equal("#123456", reopened.GetColor(imported.ForegroundColorIndex!.Value)!.ToString());
        Assert.Equal("#FFFF00", reopened.GetColor(imported.HighlightColorIndex!.Value)!.ToString());
    }

    [Fact]
    public void Docx_Contains_Explicit_Off_Values_Underneath_Inherited_Styles() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        using WordDocument word = document.ToWordDocument();
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using var package = new System.IO.Compression.ZipArchive(bytes, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
        using Stream xml = package.GetEntry("word/document.xml")!.Open();
        System.Xml.Linq.XDocument documentXml = System.Xml.Linq.XDocument.Load(xml);
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        System.Xml.Linq.XElement cleared = documentXml.Descendants(w + "r").Single(run => run.Element(w + "t")?.Value == "Cleared");
        Assert.True(cleared.Element(w + "rPr")?.Element(w + "b")?.Attribute(w + "val")?.Value is "false" or "0" or "off");
        Assert.True(cleared.Element(w + "rPr")?.Element(w + "i")?.Attribute(w + "val")?.Value is "false" or "0" or "off");
        Assert.Equal("none", cleared.Element(w + "rPr")?.Element(w + "u")?.Attribute(w + "val")?.Value);
        System.Xml.Linq.XElement paragraph = documentXml.Descendants(w + "p").Single(item => item.Descendants(w + "t").Any(text => text.Value == "Explicit"));
        Assert.True(paragraph.Element(w + "pPr")?.Element(w + "keepNext")?.Attribute(w + "val")?.Value is "false" or "0" or "off");
    }

    [Fact]
    public void Default_Style_Zero_And_Explicit_Plain_Color_And_Border_Resets_Survive_Normalization() {
        const string input = @"{\rtf1\ansi{\stylesheet{\s0\b\fs36\cf1\qc\brdrt\brdrs Normal;}}{\colortbl;\red255\green0\blue0;}\pard Inherited {\b0\cf0 Cleared} {\plain Default}\par\pard\brdrt\brdrnil NoBorder\par}";
        RtfDocument source = RtfDocument.Read(input).Document;
        AssertDefaults(source);
        AssertDefaults(RtfDocument.Read(source.ToRtf()).Document);
        string html = source.ToHtml(new RtfToHtmlOptions { IncludeRoundTripMetadata = true, FragmentOnly = false });
        AssertDefaults(HtmlConversionDocument.Parse(html).ToRtfDocument());

        static void AssertDefaults(RtfDocument document) {
            RtfParagraph paragraph = document.Paragraphs[0];
            Assert.Equal(RtfTextAlignment.Center, document.ResolveParagraphFormatting(paragraph).Alignment);
            Assert.True(document.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Inherited ")).Bold);
            RtfRun cleared = document.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Cleared"));
            Assert.False(cleared.Bold);
            Assert.Equal(0, cleared.ForegroundColorIndex);
            RtfRun plain = document.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Default"));
            Assert.False(plain.Bold);
            Assert.Equal(12, plain.FontSize);
            Assert.Equal(0, plain.ForegroundColorIndex);
            Assert.Equal(RtfParagraphBorderStyle.Single, document.ResolveParagraphFormatting(paragraph).TopBorder.Style);
            Assert.Equal(RtfParagraphBorderStyle.None, document.ResolveParagraphFormatting(document.Paragraphs[1]).TopBorder.Style);
        }
    }

    [Fact]
    public void Resolver_Applies_Kind_Specific_Bases_And_Explicit_Default_Overrides() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        RtfParagraph first = document.Paragraphs[0];
        Assert.Null(first.DirectAlignment);
        RtfParagraph effective = document.ResolveParagraphFormatting(first);
        Assert.Equal(RtfTextAlignment.Center, effective.Alignment);
        Assert.True(effective.KeepWithNext);
        Assert.Equal(720, effective.LeftIndentTwips);
        RtfRun inherited = document.ResolveRunFormatting(first, first.Runs.Single(run => run.Text == "Inherited "));
        Assert.True(inherited.Bold);
        Assert.True(inherited.Italic);
        Assert.False(inherited.Underline);
        RtfRun cleared = document.ResolveRunFormatting(first, first.Runs.Single(run => run.Text == "Cleared"));
        Assert.False(cleared.Bold);
        Assert.False(cleared.Italic);
        Assert.False(cleared.Underline);
        Assert.True(document.ResolveRunFormatting(first, first.Runs.Single(run => run.Text == "Character")).Underline);
        RtfParagraph explicitParagraph = document.ResolveParagraphFormatting(document.Paragraphs[1]);
        Assert.Equal(RtfTextAlignment.Left, explicitParagraph.Alignment);
        Assert.False(explicitParagraph.KeepWithNext);
        Assert.Equal(0, explicitParagraph.LeftIndentTwips);
    }

    [Fact]
    public void Normalized_Save_Retains_Inheritance_For_Subsequent_Style_Edits() {
        RtfDocument original = RtfDocument.Read(Input).Document;
        RtfDocument document = RtfDocument.Read(original.ToRtf(new RtfWriteOptions { MaterializeStyleFormatting = false })).Document;
        Assert.Null(document.Paragraphs[0].DirectAlignment);
        RtfRun inherited = document.Paragraphs[0].Runs.Single(run => run.Text == "Inherited ");
        Assert.Null(inherited.DirectBold);
        Assert.True(document.ResolveRunFormatting(document.Paragraphs[0], inherited).Bold);
        document.Styles.Single(style => style.Kind == RtfStyleKind.Paragraph && style.Id == 1).Bold = false;
        Assert.False(document.ResolveRunFormatting(document.Paragraphs[0], inherited).Bold);
        Assert.False(document.ResolveRunFormatting(document.Paragraphs[0], document.Paragraphs[0].Runs.Single(run => run.Text == "Cleared")).Bold);
    }

    [Fact]
    public void Html_And_Markdown_Render_Resolved_Formatting_While_Html_Metadata_Retains_Authored_Overrides() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        string html = document.ToHtml(new RtfToHtmlOptions { FragmentOnly = false, IncludeRoundTripMetadata = true });
        Assert.Contains("text-align:center", html, StringComparison.Ordinal);
        Assert.Contains("<strong><em>Inherited ", html, StringComparison.Ordinal);
        Assert.Contains("***Inherited", document.ToMarkdown(), StringComparison.Ordinal);
        RtfDocument reopened = HtmlConversionDocument.Parse(html).ToRtfDocument();
        Assert.Null(reopened.Paragraphs[0].DirectAlignment);
        Assert.Null(reopened.Paragraphs[0].Runs.Single(run => run.Text == "Inherited ").DirectBold);
        Assert.False(reopened.Paragraphs[0].Runs.Single(run => run.Text == "Cleared").DirectBold);
        Assert.True(reopened.ResolveRunFormatting(reopened.Paragraphs[0], reopened.Paragraphs[0].Runs.Single(run => run.Text == "Inherited ")).Bold);
    }

    [Fact]
    public void Pdf_And_Word_Use_Resolved_Style_Formatting() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        OfficeIMO.Pdf.PdfDocument pdf = document.ToPdfDocument();
        OfficeIMO.Pdf.RichParagraphBlock paragraph = Assert.IsType<OfficeIMO.Pdf.RichParagraphBlock>(pdf.Blocks.First());
        Assert.True(paragraph.Runs.First(run => run.Text == "Inherited ").Bold);
        Assert.False(paragraph.Runs.First(run => run.Text == "Cleared").Bold);
        using WordDocument word = document.ToWordDocument();
        Assert.True(word.Paragraphs.First(run => run.Text == "Inherited ").Bold);
        Assert.False(word.Paragraphs.First(run => run.Text == "Cleared").Bold);
    }

    [Fact]
    public void Cyclic_Style_Bases_Are_Bounded_And_Resolve_The_Selected_Styles_Overrides_First() {
        RtfDocument document = RtfDocument.Create();
        document.AddStyle(1, "One").BasedOnStyleId = 2;
        RtfStyle two = document.AddStyle(2, "Two");
        two.BasedOnStyleId = 1;
        two.Bold = true;
        RtfParagraph paragraph = document.AddParagraph("Bounded");
        paragraph.StyleId = 1;
        Assert.True(document.ResolveRunFormatting(paragraph, paragraph.Runs[0]).Bold);
    }
}
