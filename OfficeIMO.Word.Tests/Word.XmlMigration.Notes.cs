using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void WordMl2003_IndependentNotesRetainAnchorsTextFormattingAndEditableDocx() {
        string path = GetFixtureDoc(Path.Combine("WordMl2003Notes", "libreoffice-26.8.1-notes.xml"));
        byte[] original = File.ReadAllBytes(path);
        using WordDocument document = WordDocument.Load(path, new WordLoadOptions(), out OfficeConversionReport report);
        Assert.DoesNotContain(report.Diagnostics.Where(diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission), diagnostic => diagnostic.Location?.EndsWith("}footnote", StringComparison.Ordinal) == true
            || diagnostic.Location?.EndsWith("}endnote", StringComparison.Ordinal) == true
            || diagnostic.Location?.EndsWith("}footnoteRef", StringComparison.Ordinal) == true
            || diagnostic.Location?.EndsWith("}suppressRef", StringComparison.Ordinal) == true);
        OfficeDiagnostic[] referenceLosses = report.Diagnostics.Where(diagnostic => diagnostic.Code == "WORDML2003_NOTE_REFERENCE_APPROXIMATION").ToArray();
        Assert.Equal(new[] { "{http://schemas.microsoft.com/office/word/2003/wordml}footnote", "{http://schemas.microsoft.com/office/word/2003/wordml}endnote" }, referenceLosses.Select(diagnostic => diagnostic.Location));
        Assert.All(referenceLosses, diagnostic => Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind));
        AssertWordMl2003Notes(document, edited: false);
        document.FootNotes[0].Paragraphs!.Single(run => run.Text == "bold").SetText("edited bold");
        document.EndNotes[1].Paragraphs!.Single(run => run.Text == "E2 independent text.").SetText("Edited E2 text.");
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        AssertWordMl2003Notes(reopened, edited: true);
        Assert.Empty(reopened.ValidateDocument());
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    private static void AssertWordMl2003Notes(WordDocument document, bool edited) {
        Assert.Equal(3, document.FootNotes.Count);
        Assert.Equal(3, document.EndNotes.Count);
        Assert.Equal(new[] { "\tF1 first " + (edited ? "edited bold" : "bold") + " text.", "", "F1 secondtab italic." }, document.FootNotes[0].TextParagraphs);
        Assert.Equal(new[] { "\tF2 independent text." }, document.FootNotes[1].TextParagraphs);
        Assert.Equal(new[] { "*\tCustom footnote text." }, document.FootNotes[2].TextParagraphs);
        Assert.Equal(new[] { "\tE1 first paragraph.", "E1 second paragraph." }, document.EndNotes[0].TextParagraphs);
        Assert.Equal(new[] { "\t" + (edited ? "Edited E2 text." : "E2 independent text.") }, document.EndNotes[1].TextParagraphs);
        Assert.Equal(new[] { "!\tCustom endnote text." }, document.EndNotes[2].TextParagraphs);
        W.Body body = document.OpenXmlDocument.MainDocumentPart!.Document.Body!;
        string[] text = body.Elements<W.Paragraph>().Select(paragraph => string.Concat(paragraph.Descendants().Select(element => element switch {
            W.Text literal => literal.Text,
            W.FootnoteReference note => "[F" + note.Id!.Value + "]",
            W.EndnoteReference note => "[E" + note.Id!.Value + "]",
            _ => ""
        }))).ToArray();
        Assert.Equal(new[] {
            "Before footnote[F1] between endnote[E1] after.",
            "Another footnote[F2] then endnote[E2] done.",
            "Custom footnote[F3] and custom endnote[E3] end."
        }, text);
        Assert.Equal(new[] { false, false, true }, body.Descendants<W.FootnoteReference>().Select(note => note.CustomMarkFollows?.Value ?? false));
        Assert.Equal(new[] { false, false, true }, body.Descendants<W.EndnoteReference>().Select(note => note.CustomMarkFollows?.Value ?? false));
        var main = document.OpenXmlDocument.MainDocumentPart;
        Assert.Equal(2, main.FootnotesPart!.Footnotes.Descendants<W.FootnoteReferenceMark>().Count());
        Assert.Equal(2, main.EndnotesPart!.Endnotes.Descendants<W.EndnoteReferenceMark>().Count());
        Assert.Empty(main.EndnotesPart.Endnotes.Descendants<W.FootnoteReferenceMark>());
        WordParagraph bold = document.FootNotes[0].Paragraphs!.Single(run => run.Text == (edited ? "edited bold" : "bold"));
        Assert.Contains(WordRunFormatting.EnumerateRunStyleProperties(bold), properties => properties.GetFirstChild<W.Bold>() is { } value && (value.Val?.Value ?? true));
        WordParagraph italic = document.FootNotes[0].Paragraphs!.Single(run => run.Text == "italic");
        Assert.Contains(WordRunFormatting.EnumerateRunStyleProperties(italic), properties => properties.GetFirstChild<W.Italic>() is { } value && (value.Val?.Value ?? true));
        WordRunInfo[] inspected = document.Inspect().Sections.SelectMany(section => section.Elements).OfType<WordParagraphInfo>()
            .SelectMany(paragraph => paragraph.Runs).ToArray();
        Assert.Equal(document.FootNotes.Select(note => note.ReferenceId), inspected.Where(run => run.Footnote != null).Select(run => (long?)run.Footnote!.ReferenceId));
        Assert.Equal(document.EndNotes.Select(note => note.ReferenceId), inspected.Where(run => run.Endnote != null).Select(run => (long?)run.Endnote!.ReferenceId));
    }

    [Theory]
    [InlineData("nested")]
    [InlineData("header")]
    [InlineData("special-type")]
    [InlineData("invalid-suppress")]
    [InlineData("orphan-mark")]
    [InlineData("suppressed-reference")]
    public void WordMl2003_UnqualifiedNotesRemainExplicitStrictLoss(string profile) {
        string note = "<w:footnote><w:p><w:r><w:t>Note content</w:t></w:r></w:p></w:footnote>";
        string content = profile switch {
            "nested" => "<w:p><w:r><w:footnote><w:p><w:r><w:t>Outer retained</w:t>" + note + "</w:r></w:p></w:footnote></w:r></w:p>",
            "header" => "<w:sectPr><w:hdr w:type=\"odd\"><w:p><w:r>" + note + "</w:r></w:p></w:hdr></w:sectPr>",
            "special-type" => "<w:p><w:r>" + note.Replace("<w:footnote>", "<w:footnote w:type=\"separator\">") + "</w:r></w:p>",
            "suppressed-reference" => "<w:p><w:r>" + note.Replace("<w:footnote>", "<w:footnote w:suppressRef=\"on\">") + "</w:r></w:p>",
            "invalid-suppress" => "<w:p><w:r>" + note.Replace("<w:footnote>", "<w:footnote w:suppressRef=\"unknown\">") + "</w:r></w:p>",
            _ => "<w:p><w:r><w:footnoteRef/></w:r></w:p>"
        };
        byte[] input = Encoding.UTF8.GetBytes(WordMlNotesXml(content));
        using WordDocument document = WordDocument.Load(new MemoryStream(input), new WordLoadOptions(), out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.LossKind == (profile == "suppressed-reference" ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.Omission)
            && diagnostic.Message.IndexOf("note", StringComparison.Ordinal) >= 0);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(new MemoryStream(input), new WordLoadOptions { FailOnLoss = true }));
        if (profile == "nested") Assert.Equal(new[] { "Outer retained" }, Assert.Single(document.FootNotes).TextParagraphs);
        else if (profile == "suppressed-reference") Assert.Equal(new[] { "Note content" }, Assert.Single(document.FootNotes).TextParagraphs);
        else Assert.Empty(document.FootNotes);
        using var output = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData("footnote")]
    [InlineData("endnote")]
    public void WordMl2003_NotePartsRespectReconstructionLimits(string kind) {
        const string body = "<w:p><w:r><w:t>Note</w:t></w:r></w:p>";
        XElement source = XDocument.Parse(WordMlNotesXml("<w:p><w:r><w:" + kind + ">" + body + "</w:" + kind + "></w:r></w:p>")).Root!;
        foreach (bool parts in new[] { true, false }) {
            var policy = new OfficePackageSecurityOptions { MaxPartCount = parts ? 4 : 10, MaxRelationships = parts ? 10 : 1 };
            using var rejected = new MemoryStream();
            var error = Assert.Throws<InvalidDataException>(() => WordMl2003Projection.WritePackage(source, rejected, policy, default));
            Assert.Contains("relationship or part limit", error.Message);
            if (parts) policy.MaxPartCount++;
            else policy.MaxRelationships++;
            using var accepted = new MemoryStream();
            Assert.Empty(WordMl2003Projection.WritePackage(source, accepted, policy, default));
            accepted.Position = 0;
            using WordDocument document = WordDocument.Load(accepted);
            Assert.Equal(1, document.FootNotes.Count + document.EndNotes.Count);
        }
    }

    [Theory]
    [InlineData("footnote")]
    [InlineData("endnote")]
    public void WordMl2003_NoteTablesPicturesAndLinksKeepTheirPartOwnership(string kind) {
        byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=");
        string picture = "<w:pict><v:shape id=\"pixel\" style=\"width:1pt;height:1pt\"><v:imagedata src=\"pixel.png\"/></v:shape></w:pict>";
        string note = "<w:" + kind + "><w:p><w:r><w:footnoteRef/>" + picture + picture.Replace("id=\"pixel\"", "id=\"pixel2\"") + "</w:r>" +
            "<w:hlink w:dest=\"https://example.com/note\"><w:r><w:t>Note link</w:t></w:r></w:hlink></w:p>" +
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"1440\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Editable cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:" + kind + ">";
        string resource = "<w:p><w:r><w:pict><w:binData w:name=\"pixel.png\">" + Convert.ToBase64String(pixel) + "</w:binData></w:pict></w:r></w:p>";
        string xml = WordMlNotesXml(resource + "<w:p><w:r>" + note + picture + "<w:t>Body after note</w:t></w:r></w:p>")
            .Replace("<w:wordDocument ", "<w:wordDocument xmlns:v=\"urn:schemas-microsoft-com:vml\" ");
        using WordDocument document = WordDocument.Load(new MemoryStream(Encoding.UTF8.GetBytes(xml)),
            new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var saved = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(saved, false);
        MainDocumentPart main = package.MainDocumentPart!;
        OpenXmlPart notePart = kind == "footnote" ? main.FootnotesPart! : main.EndnotesPart!;
        Assert.Single(main.ImageParts);
        Assert.Single(notePart.Parts);
        Assert.Single(notePart.HyperlinkRelationships);
        Assert.Empty(main.HyperlinkRelationships);
        ImagePart noteImage = Assert.IsType<ImagePart>(Assert.Single(notePart.Parts).OpenXmlPart);
        using var imageBytes = new MemoryStream();
        using (Stream image = noteImage.GetStream()) image.CopyTo(imageBytes);
        Assert.Equal(pixel, imageBytes.ToArray());
        OpenXmlElement noteRoot = kind == "footnote" ? main.FootnotesPart!.Footnotes : main.EndnotesPart!.Endnotes;
        Assert.Equal("Editable cell", Assert.Single(noteRoot.Descendants<W.Table>()).InnerText);
        Assert.Equal(2, noteRoot.Descendants<W.Picture>().Count());
        Assert.Single(main.Document.Descendants<W.Picture>());
        Assert.Empty(document.ValidateDocument());
    }

    private static string WordMlNotesXml(string body) =>
        "<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\"><w:body>" + body + "</w:body></w:wordDocument>";
}
