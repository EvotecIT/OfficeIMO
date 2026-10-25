using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfNativeNoteModeTests {
    [Theory]
    [InlineData("", false, true, true)]
    [InlineData(@"\fet0", false, true, true)]
    [InlineData(@"\ftnbj", false, false, false)]
    [InlineData(@"\endnotes", false, true, true)]
    [InlineData(@"\fet1\endnotes\aendnotes", false, false, false)]
    [InlineData(@"\fet2\ftnbj", false, false, false)]
    [InlineData(@"\fet2\ftnbj", true, true, false)]
    [InlineData(@"\fet2", false, false, false)]
    [InlineData(@"\fet2\endnotes", false, false, false)]
    public void Parsed_Mode_Projects_Native_Reference_And_All_Note_Paragraphs(
        string controls, bool alternate, bool endnote, bool sectionEnd) {
        RtfReadResult read = RtfDocument.Read(Source(controls, alternate));
        Assert.False(read.HasOriginalBytes); // Character input: policy cannot depend on a retained byte buffer.
        RtfNote raw = Assert.Single(read.Document.Notes);
        Assert.Equal(alternate ? RtfNoteKind.Endnote : RtfNoteKind.Footnote, raw.Kind);
        Assert.Equal(new[] { "First note paragraph", "Second note paragraph" }, raw.Paragraphs.Select(p => p.ToPlainText()));

        using WordDocument word = read.Document.ToWordDocument();
        using var stream = new MemoryStream();
        word.Save(stream);
        stream.Position = 0;
        using WordprocessingDocument persisted = WordprocessingDocument.Open(stream, false);
        MainDocumentPart main = persisted.MainDocumentPart!;
        Assert.Equal(endnote ? 1 : 0, main.Document!.Descendants<EndnoteReference>().Count());
        Assert.Equal(endnote ? 0 : 1, main.Document.Descendants<FootnoteReference>().Count());
        var note = endnote
            ? (DocumentFormat.OpenXml.OpenXmlCompositeElement)Assert.Single(main.EndnotesPart!.Endnotes!.Elements<Endnote>(), n => n.Id!.Value > 0)
            : Assert.Single(main.FootnotesPart!.Footnotes!.Elements<Footnote>(), n => n.Id!.Value > 0);
        Assert.Equal(new[] { "First note paragraph", "Second note paragraph" }, note.Elements<Paragraph>().Select(p => p.InnerText));
        EndnotePositionValues placement = main.DocumentSettingsPart!.Settings!
            .GetFirstChild<EndnoteDocumentWideProperties>()?.GetFirstChild<EndnotePosition>()?.Val?.Value
            ?? EndnotePositionValues.DocumentEnd;
        if (endnote) Assert.Equal(sectionEnd ? EndnotePositionValues.SectionEnd : EndnotePositionValues.DocumentEnd, placement);
        if (!endnote) Assert.NotEqual(FootnotePositionValues.SectionEnd,
            main.Document.Body!.Descendants<SectionProperties>().Last().GetFirstChild<FootnoteProperties>()?.FootnotePosition?.Val?.Value);
        Assert.Equal(alternate ? RtfNoteKind.Endnote : RtfNoteKind.Footnote, raw.Kind);
    }

    [Theory]
    [InlineData("", @"\fet0", true)]
    [InlineData(@"\fet0", @"\fet0", true)]
    [InlineData(@"\fet1", @"\fet1", false)]
    [InlineData(@"\fet2", @"\fet2", false)]
    public void Clone_Metadata_Edit_And_Normalization_Retain_Parsed_Mode(
        string controls, string normalizedMode, bool endnote) {
        RtfReadResult read = RtfDocument.Read(Source(controls));
        Assert.Equal(Source(controls), read.ToRtfLossless());
        RtfDocument copy = read.Document.Clone();
        copy.Info.Title = "Edited metadata";
        string written = copy.ToRtf();
        Assert.Contains(normalizedMode, written);
        if (normalizedMode != @"\fet2") Assert.DoesNotContain(@"\fet2", written);
        using WordDocument direct = copy.ToWordDocument();
        using WordDocument normalized = RtfDocument.Read(written).Document.ToWordDocument();
        Assert.Equal(endnote ? 1 : 0, direct.EndNotes.Count);
        Assert.Equal(endnote ? 1 : 0, normalized.EndNotes.Count);
        Assert.Equal(endnote ? 0 : 1, normalized.FootNotes.Count);
        Assert.Null(read.Document.Info.Title);
    }

    [Fact]
    public void Empty_Append_Adopts_Parsed_Note_Defaults_And_Declared_Settings() {
        RtfDocument source = RtfDocument.Read(Source(@"\endnotes\ftnstart3\ftnnrlc")).Document;
        RtfDocument destination = RtfDocument.Create();
        RtfDocumentMergeResult result = destination.AppendDocument(source);
        result.Report.RequireNoLoss();
        using WordDocument word = destination.ToWordDocument();
        Assert.Single(word.EndNotes);
        Assert.Empty(word.FootNotes);
        Assert.Equal(3, word.EndnoteSettings.StartNumber);
        Assert.Equal(WordNumberFormat.LowerRoman, word.EndnoteSettings.NumberingFormat);
        Assert.Equal(WordEndnotePosition.SectionEnd, word.EndnoteSettings.Position);
        Assert.Contains(@"\fet0", destination.ToRtf());
        Assert.Equal(RtfNoteKind.Footnote, Assert.Single(source.Notes).Kind);
    }

    [Fact]
    public void Incompatible_Nonempty_Append_Reports_Document_Note_Policy_Loss() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddParagraph("Existing modern body");
        RtfDocumentMergeResult result = destination.AppendDocument(RtfDocument.Read(Source("")).Document);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "RtfMergeDocumentSettingsFlattened" && d.Detail!.Contains("NoteMode"));
        Assert.Throws<RtfConversionLossException>(result.Report.RequireNoLoss);
        Assert.Contains(@"\fet2", destination.ToRtf());
    }

    [Fact]
    public void Fresh_Authored_Footnote_Remains_Modern_Across_Write_And_Read() {
        RtfDocument source = RtfDocument.Create();
        RtfNote note = source.AddNote(RtfNoteKind.Footnote);
        note.AddParagraph("Authored note");
        source.AddParagraph("Body").AddNoteReference(note);
        string written = source.ToRtf();
        Assert.Contains(@"\fet2", written);
        using WordDocument word = RtfDocument.Read(written).Document.ToWordDocument();
        Assert.Single(word.FootNotes);
        Assert.Empty(word.EndNotes);
    }

    [Fact]
    public void Projected_Ordinary_And_Explicit_Endnote_Group_Collision_Is_Reported() {
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi\pard Body{\footnote First}{\footnote\ftnalt Second}\par}").Document;
        RtfConversionResult<WordDocument> result = source.ToWordDocumentResult();
        using WordDocument word = result.Value;
        Assert.Equal(2, word.EndNotes.Count);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "RtfWordNoteGroupsFlattened");
        Assert.Throws<RtfConversionLossException>(result.Report.RequireNoLoss);
    }

    [Theory]
    [InlineData(@"\ftnbj", @"\endnotes", true)]
    [InlineData(@"\endnotes", @"\ftnbj", false)]
    public void Parsed_Section_Overrides_Resolve_Table_Cell_Note_References(
        string documentPlacement, string sectionPlacement, bool endnote) {
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi" + documentPlacement + @"\sectd" + sectionPlacement +
            @"\trowd\cellx2400\pard\intbl Cell{\footnote\pard First\par Second}\cell\row}").Document;
        using WordDocument word = source.ToWordDocument();
        Assert.Equal(endnote ? 1 : 0, word.EndNotes.Count);
        Assert.Equal(endnote ? 0 : 1, word.FootNotes.Count);
        Assert.Equal(2, endnote ? Assert.Single(word.EndNotes).Paragraphs!.Select(p => p._paragraph).Distinct().Count()
            : Assert.Single(word.FootNotes).Paragraphs!.Select(p => p._paragraph).Distinct().Count());
    }

    [Fact]
    public void Unsupported_Mode_Is_Lossless_In_Syntax_And_Diagnosed_During_Normalization() {
        string text = Source(@"\fet9");
        RtfReadResult read = RtfDocument.Read(text);
        Assert.Equal(text, read.ToRtfLossless());
        Assert.Contains(read.Diagnostics, d => d.Code == "RTF108");
        RtfConversionResult<string> normalized = read.Document.ToRtfResult();
        Assert.True(normalized.Report.HasLoss);
        Assert.Throws<RtfConversionLossException>(normalized.Report.RequireNoLoss);
    }

    internal static string Source(string controls, bool alternate = false) =>
        @"{\rtf1\ansi" + controls + @"\pard Body{\footnote" + (alternate ? @"\ftnalt" : "") +
        @"\pard First note paragraph\par Second note paragraph}\par}";
}
