using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlNativeNoteModeTests {
    [Theory]
    [InlineData("", false, @"\fet0", true)]
    [InlineData(@"\fet0", false, @"\fet0", true)]
    [InlineData(@"\fet1", false, @"\fet1", false)]
    [InlineData(@"\fet2", false, @"\fet2", false)]
    [InlineData("", true, @"\fet2", false)]
    public void RoundTrip_Metadata_Retains_Native_Note_Policy_And_All_Paragraphs(
        string controls, bool authored, string expectedHeader, bool endnote) {
        RtfDocument source;
        if (authored) {
            source = RtfDocument.Create();
            RtfNote note = source.AddNote(RtfNoteKind.Footnote);
            note.AddParagraph("First note paragraph");
            note.AddParagraph("Second note paragraph");
            source.AddParagraph("Body").AddNoteReference(note);
        } else source = RtfDocument.Read(RtfNativeNoteModeTests.Source(controls)).Document;
        string html = source.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());
        RtfDocument returned = HtmlConversionDocument.Parse(html).ToRtfDocument();
        RtfNote raw = Assert.Single(returned.Paragraphs.SelectMany(paragraph => paragraph.Inlines)
            .Select(inline => inline is RtfRun run ? run.Note : (inline as RtfGeneratedText)?.Note).OfType<RtfNote>());
        Assert.Equal(RtfNoteKind.Footnote, raw.Kind);
        Assert.Equal(new[] { "First note paragraph", "Second note paragraph" }, raw.Paragraphs.Select(p => p.ToPlainText()));
        Assert.Contains(expectedHeader, returned.ToRtf());
        using WordDocument word = returned.ToWordDocument();
        using var stream = new MemoryStream();
        word.Save(stream);
        stream.Position = 0;
        using WordprocessingDocument persisted = WordprocessingDocument.Open(stream, false);
        MainDocumentPart main = persisted.MainDocumentPart!;
        Assert.Equal(endnote ? 1 : 0, main.Document!.Descendants<EndnoteReference>().Count());
        Assert.Equal(endnote ? 0 : 1, main.Document.Descendants<FootnoteReference>().Count());
        Assert.Equal(new[] { "First note paragraph", "Second note paragraph" }, endnote
            ? Assert.Single(main.EndnotesPart!.Endnotes!.Elements<Endnote>(), n => n.Id!.Value > 0).Elements<Paragraph>().Select(p => p.InnerText)
            : Assert.Single(main.FootnotesPart!.Footnotes!.Elements<Footnote>(), n => n.Id!.Value > 0).Elements<Paragraph>().Select(p => p.InnerText));
    }
}
