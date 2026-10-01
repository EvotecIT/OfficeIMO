using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Word.Rtf;
using OfficeIMO.ContentSafety;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfStyleVisibilityRegressionTests {
    private const string Input = @"{\rtf1\ansi{\stylesheet{\s0 Normal;}{\s1\v HiddenText;}{\s2\shidden HiddenStyleUI;}}
\pard\s1 Secret marker {\v0 Visible marker} {\plain Plain marker}\par
\pard\s2 UI style marker\par}";

    [Fact]
    public void LegacyHtmlRunMetadataRetainsCssHiddenText() {
        string legacy = Convert.ToBase64String(Encoding.UTF8.GetBytes("version=MQ==\nplain=ZmFsc2U=\n"));
        string html = "<p><span data-officeimo-rtf-direct-run=\"" + legacy + "\"><span style=\"visibility:hidden\">Secret</span></span><span>Visible</span></p>";
        RtfDocument imported = HtmlConversionDocument.Parse(html).ToRtfDocument();
        RtfParagraph paragraph = Assert.Single(imported.Paragraphs);
        Assert.True(imported.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Secret")).Hidden);
        Assert.False(imported.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Visible")).Hidden);
    }

    [Fact]
    public void ContentSafetyFindsInheritedHiddenTextAndCleansOnlyTheSelectedRun() {
        byte[] input = Encoding.ASCII.GetBytes(Input);
        var inspected = RtfDocument.InspectContentSafety(input);
        var hidden = Assert.Single(inspected.Findings, item => item.Kind == OfficeContentConcealmentKind.HiddenByProperty);
        Assert.Contains("Secret marker", hidden.TextPreview, StringComparison.Ordinal);
        var cleaned = RtfDocument.RemoveSelectedContent(input, new OfficeContentCleanupSelection(new[] { hidden.Id }));
        string text = string.Join("\n", RtfDocument.Load(cleaned.Output).Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        Assert.DoesNotContain("Secret marker", text, StringComparison.Ordinal);
        Assert.Contains("Visible marker", text, StringComparison.Ordinal);
        Assert.Contains("Plain marker", text, StringComparison.Ordinal);
        Assert.Contains("UI style marker", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentSafetyUsesInheritedHighlightAndHonorsDirectAutomaticColor() {
        const string input = @"{\rtf1\ansi{\colortbl;\red0\green0\blue0;}{\stylesheet{\s1\cf1\highlight1 Concealed;}}\pard\s1 Concealed {\highlight0 Visible}\par}";
        var inspected = RtfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(input));
        var concealed = Assert.Single(inspected.Findings, item => item.Kind == OfficeContentConcealmentKind.LowContrastText);
        Assert.Contains("Concealed", concealed.TextPreview, StringComparison.Ordinal);
        Assert.DoesNotContain(inspected.Findings, item => item.TextPreview.Contains("Visible"));
    }

    [Fact]
    public void StylesheetHiddenTextResolvesWithoutConfusingStyleGalleryVisibility() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        AssertVisibility(document);
        AssertVisibility(document.Clone());
        AssertVisibility(RtfDocument.Read(document.ToRtf()).Document);

        static void AssertVisibility(RtfDocument document) {
            RtfParagraph paragraph = document.Paragraphs[0];
            Assert.True(document.ResolveRunFormatting(paragraph, paragraph.Runs.First()).Hidden);
            Assert.False(document.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Visible marker")).Hidden);
            Assert.False(document.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Plain marker")).Hidden);
            Assert.False(document.ResolveRunFormatting(document.Paragraphs[1], document.Paragraphs[1].Runs.First()).Hidden);
        }
    }

    [Fact]
    public void PdfOmitsInheritedHiddenTextAndRetainsExplicitVisibleOverrides() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes(new RtfToPdfOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        })).ExtractText();
        Assert.DoesNotContain("Secret marker", text, StringComparison.Ordinal);
        Assert.Contains("Visible marker", text, StringComparison.Ordinal);
        Assert.Contains("Plain marker", text, StringComparison.Ordinal);
        Assert.Contains("UI style marker", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlMetadataRetainsVisibilityInheritanceForLaterStyleEdits() {
        RtfDocument source = RtfDocument.Read(Input).Document;
        string html = source.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());
        Assert.Contains("visibility:hidden", html, StringComparison.Ordinal);
        RtfDocument document = HtmlConversionDocument.Parse(html).ToRtfDocument();
        RtfParagraph paragraph = document.Paragraphs[0];
        RtfRun inherited = paragraph.Runs.First();
        Assert.Null(inherited.DirectHidden);
        Assert.True(document.ResolveRunFormatting(paragraph, inherited).Hidden);
        Assert.False(paragraph.Runs.Single(run => run.Text == "Visible marker").DirectHidden);
        document.Styles.Single(style => style.Id == 1).TextHidden = false;
        Assert.False(document.ResolveRunFormatting(paragraph, inherited).Hidden);
    }

    [Fact]
    public void WordStylesAndExplicitVisibleOverridesSurviveSemanticReimport() {
        RtfDocument document = RtfDocument.Read(Input).Document;
        using var word = document.ToWordDocument();
        RtfDocument reopened = word.ToRtfDocument();
        RtfParagraph paragraph = reopened.Paragraphs[0];
        Assert.True(reopened.Styles.Single(style => style.Id == paragraph.StyleId && style.Kind == RtfStyleKind.Paragraph).TextHidden);
        Assert.True(reopened.ResolveRunFormatting(paragraph, paragraph.Runs.First()).Hidden);
        Assert.False(reopened.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Visible marker")).Hidden);
        Assert.False(reopened.ResolveRunFormatting(paragraph, paragraph.Runs.Single(run => run.Text == "Plain marker")).Hidden);
    }
}
