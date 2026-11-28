using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

/// <summary>
/// Word content keeps its HTML export when the markup around it changes without changing what it is: an XML comment
/// between paragraph children, a reference or text the package holds untyped, text moved into a selected alternate
/// branch, and a control wrapped in a run-level conflict mark.
/// </summary>
public sealed class HtmlWordMarkupPlacementTests {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string MathMarkup =
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";

    [Fact]
    public void AnXmlCommentBeforeAnInlineEquationKeepsTheEquationInPlace() {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Before");
        paragraph.AddEquation(MathMarkup);
        paragraph.AddText("After");
        string expected = document.ToHtml();

        W.Paragraph sdkParagraph = paragraph._paragraph;
        OpenXmlElement math = sdkParagraph.ChildElements.First(child => child.LocalName == "oMath");
        sdkParagraph.InsertBefore(new OpenXmlMiscNode(XmlNodeType.Comment, "<!--note-->"), math);
        string actual = document.ToHtml();

        Assert.Equal(expected, actual);
        int before = actual.IndexOf("Before", StringComparison.Ordinal);
        int equation = actual.IndexOf("<math", StringComparison.Ordinal);
        int after = actual.IndexOf("After", StringComparison.Ordinal);
        Assert.True(before >= 0 && before < equation && equation < after, actual);
    }

    [Theory]
    [InlineData("footnoteReference")]
    [InlineData("endnoteReference")]
    public void AnUntypedNoteReferenceInsideANoteIsSkipped(string referenceName) {
        using var document = WordDocument.Create();
        WordParagraph body = document.AddParagraph("Body");
        body.AddFootNote("Foot");
        body.AddEndNote("End");
        string expected = document.ToHtml();

        W.Run noteRun = NoteRuns(document, referenceName == "footnoteReference").Last();
        noteRun.AppendChild(new OpenXmlUnknownElement("w", referenceName, WordNamespace));
        string actual = document.ToHtml();

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void DeletedTextInTheSelectedFallbackStaysInTheOriginalView() {
        using var document = WordDocument.Create();
        document.AddParagraph("Prefix ").AddDeletedText("old wording", "Reviewer");
        var options = new WordHtmlExportOptions { TrackedChangePolicy = WordTrackedChangeExportPolicy.Original };
        string expected = document.ToHtml(options);

        W.DeletedText deleted = document._wordprocessingDocument!.MainDocumentPart!.Document!.Body!
            .Descendants<W.DeletedText>().Single();
        OpenXmlElement run = deleted.Parent!;
        deleted.Remove();
        var fallback = new AlternateContentFallback();
        fallback.AppendChild(deleted);
        var alternate = new AlternateContent();
        alternate.AppendChild(new AlternateContentChoice { Requires = "unsupported" });
        alternate.AppendChild(fallback);
        run.AppendChild(alternate);
        string actual = document.ToHtml(options);

        Assert.Equal(expected, actual);
        Assert.Contains("old wording", actual, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("conflictIns")]
    [InlineData("conflictDel")]
    public void ACheckBoxInsideAConflictMarkStillRendersBesideAnEquation(string conflictName) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Before ");
        paragraph.AddEquation(MathMarkup);
        paragraph.AddCheckBox(true, "Approval", "Approval");
        paragraph.AddHyperLink("Link", new Uri("https://example.test"));

        W.Paragraph sdkParagraph = paragraph._paragraph;
        W.Hyperlink hyperlink = sdkParagraph.Elements<W.Hyperlink>().Single();
        OpenXmlElement math = sdkParagraph.ChildElements.First(child => child.LocalName == "oMath");
        W.SdtRun control = sdkParagraph.Elements<W.SdtRun>().Single();
        math.Remove();
        control.Remove();
        hyperlink.AppendChild(math);
        hyperlink.AppendChild(control);
        string expected = document.ToHtml();

        control.Remove();
        OpenXmlCompositeElement conflict = conflictName == "conflictIns" ? new RunConflictInsertion() : new RunConflictDeletion();
        conflict.AppendChild(control);
        hyperlink.AppendChild(conflict);
        string actual = document.ToHtml();

        Assert.Equal(expected, actual);
        Assert.Contains("type=\"checkbox\"", actual, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUntypedCommentReferenceLeavesTheRunTextAndCommentAnchorAlone() {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Seen text");
        WordComment comment = paragraph.AddComment("Reviewer", "R", "Remark");
        var options = new WordHtmlExportOptions { ExportComments = true };
        string expected = document.ToHtml(options);

        W.Run textRun = paragraph._paragraph.Descendants<W.Run>().First(run => run.InnerText == "Seen text");
        textRun.AppendChild(UntypedElement(
            $"<w:commentReference xmlns:w=\"{WordNamespace}\" w:id=\"{comment.Id}\"/>"));
        string actual = document.ToHtml(options);

        Assert.Equal(expected, actual);
        Assert.Contains("Seen text", actual, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("HTMLPreformatted")]
    public void UntypedTextInARunIsNotExported(string? styleId) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Shown");
        if (styleId != null) paragraph.SetStyleId(styleId);
        string expected = document.ToHtml();

        paragraph._paragraph.Elements<W.Run>().Single().AppendChild(UntypedElement(
            $"<w:t xmlns:w=\"{WordNamespace}\">Hidden</w:t>"));
        string actual = document.ToHtml();

        Assert.Equal(expected, actual);
        Assert.DoesNotContain("Hidden", actual, StringComparison.Ordinal);
    }

    // Word markup parsed inside an unknown wrapper stays untyped; detached, it keeps that.
    private static OpenXmlElement UntypedElement(string outerXml) {
        var wrapper = new OpenXmlUnknownElement("x", "wrapper", "urn:officeimo:tests:vendor") { InnerXml = outerXml };
        OpenXmlElement element = wrapper.FirstChild!;
        element.Remove();
        return element;
    }

    private static IEnumerable<W.Run> NoteRuns(WordDocument document, bool footnotes) {
        var main = document._wordprocessingDocument!.MainDocumentPart!;
        OpenXmlElement root = footnotes ? main.FootnotesPart!.Footnotes! : main.EndnotesPart!.Endnotes!;
        OpenXmlElement note = root.ChildElements.Last(element => element.InnerText.Contains(footnotes ? "Foot" : "End"));
        return note.Descendants<W.Run>();
    }
}
