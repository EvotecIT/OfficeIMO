using System.Text.RegularExpressions;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordPdfNoteNumberingTests {
    [Theory]
    [InlineData(WordNumberFormat.Decimal, false, "0", "1")]
    [InlineData(WordNumberFormat.UpperLetter, false, "", "A")]
    [InlineData(WordNumberFormat.LowerLetter, false, "", "a")]
    [InlineData(WordNumberFormat.UpperRoman, false, "", "I")]
    [InlineData(WordNumberFormat.LowerRoman, false, "", "i")]
    [InlineData(WordNumberFormat.Decimal, true, "0", "1")]
    [InlineData(WordNumberFormat.UpperLetter, true, "", "A")]
    [InlineData(WordNumberFormat.LowerLetter, true, "", "a")]
    [InlineData(WordNumberFormat.UpperRoman, true, "", "I")]
    [InlineData(WordNumberFormat.LowerRoman, true, "", "i")]
    public void ZeroStartPreservesBothNoteKindsAndContinuesTheirNumbering(
        WordNumberFormat format, bool documentWide, string firstLabel, string secondLabel) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].AddFootnoteProperties(format, WordFootnotePosition.PageBottom,
            WordNoteNumberRestart.Continuous, startNumber: 0);
        source.Sections[0].AddEndnoteProperties(format, WordEndnotePosition.DocumentEnd,
            WordNoteNumberRestart.Continuous, startNumber: 0);
        if (documentWide) {
            var section = source.Sections[0]._sectionProperties;
            var foot = section.GetFirstChild<W.FootnoteProperties>()!;
            var end = section.GetFirstChild<W.EndnoteProperties>()!;
            var settings = source._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            settings.AddChild(new W.FootnoteDocumentWideProperties(foot.ChildElements.Select(x => x.CloneNode(true))), true);
            settings.AddChild(new W.EndnoteDocumentWideProperties(end.ChildElements.Select(x => x.CloneNode(true))), true);
            foot.Remove();
            end.Remove();
        }
        source.AddParagraph("FIRSTREF").AddFootNote("FIRSTBODY");
        source.AddParagraph("SECONDREF").AddFootNote("SECONDBODY");
        source.AddParagraph("ENDREF").AddEndNote("ENDBODY");
        source.AddParagraph("NEXTENDREF").AddEndNote("NEXTENDBODY");

        string text = ReadPdfText(source, false);
        Assert.Contains("FIRSTREF" + firstLabel + "SECONDREF" + secondLabel, text);
        Assert.Contains("ENDREF" + firstLabel + "NEXTENDREF" + secondLabel, text);
        Assert.Contains(firstLabel + "FIRSTBODY" + secondLabel + "SECONDBODY", text);
        Assert.Contains(firstLabel + "ENDBODY" + secondLabel + "NEXTENDBODY", text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ConfiguredNoteLabelsReachReferencesAndNoteBodies(bool native, bool table) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].AddFootnoteProperties(WordNumberFormat.UpperLetter,
            WordFootnotePosition.PageBottom, WordNoteNumberRestart.Continuous, startNumber: 3);
        source.Sections[0].AddEndnoteProperties(WordNumberFormat.LowerLetter,
            WordEndnotePosition.DocumentEnd, WordNoteNumberRestart.Continuous, startNumber: 9);
        WordParagraph foot = table ? source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : source.AddParagraph();
        foot.Text = "FOOTREF";
        foot.AddFootNote("FOOTBODY");
        WordParagraph end = table ? source.Tables[0].Rows[0].Cells[0].AddParagraph() : source.AddParagraph();
        end.Text = "ENDREF";
        end.AddEndNote("ENDBODY");

        string text = ReadPdfText(source, native);
        Assert.Contains("FOOTREFC", text);
        Assert.Contains("ENDREFi", text);
        Assert.Contains("CFOOTBODY", text);
        Assert.Contains("iENDBODY", text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultFootnoteAndEndnoteLabelsUseTheirDistinctFormats(bool native) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("FOOTREF").AddFootNote("FOOTBODY");
        source.AddParagraph("ENDREF").AddEndNote("ENDBODY");

        string text = ReadPdfText(source, native);
        Assert.Contains("FOOTREF1", text);
        Assert.Contains("ENDREFi", text);
        Assert.Contains("1FOOTBODY", text);
        Assert.Contains("iENDBODY", text);
    }

    private static string ReadPdfText(WordDocument source, bool native) {
        using WordDocument restored = WordDocument.Load(new MemoryStream(
            source.ToBytes(native ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = PdfPigDocument.Open(restored.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Arial"
        }));
        return Regex.Replace(string.Concat(pdf.GetPages().Select(page => page.Text)), @"\s+", "");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SectionRestartUsesRepeatedLabelsWithoutLosingReferenceIdentity(bool restart) {
        using WordDocument source = WordDocument.Create();
        WordNoteNumberRestart policy = restart ? WordNoteNumberRestart.EachSection : WordNoteNumberRestart.Continuous;
        source.Sections[0].AddFootnoteProperties(WordNumberFormat.UpperLetter,
            WordFootnotePosition.PageBottom, policy, startNumber: 3);
        source.AddParagraph("FIRSTREF").AddFootNote("FIRSTBODY");
        WordSection second = source.AddSection();
        second.AddFootnoteProperties(WordNumberFormat.UpperLetter,
            WordFootnotePosition.PageBottom, policy, startNumber: 3);
        second.AddParagraph("SECONDREF").AddFootNote("SECONDBODY");

        string text = ReadPdfText(source, false);
        Assert.Contains("FIRSTREFC", text);
        Assert.Contains("CFIRSTBODY", text);
        Assert.Contains("SECONDREF" + (restart ? "C" : "D"), text);
        Assert.Contains((restart ? "C" : "D") + "SECONDBODY", text);
    }

    [Fact]
    public void FirstNoteAfterAnEmptySectionUsesTheContainingSectionsStart() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("INTRODUCTION");
        WordSection section = source.AddSection();
        section.AddFootnoteProperties(WordNumberFormat.UpperLetter,
            WordFootnotePosition.PageBottom, WordNoteNumberRestart.Continuous, startNumber: 3);
        section.AddParagraph("FOOTREF").AddFootNote("FOOTBODY");
        string text = ReadPdfText(source, false);
        Assert.Contains("FOOTREFC", text);
        Assert.Contains("CFOOTBODY", text);
    }

    [Fact]
    public void PageRestartApproximationIsReportedWithTheRenderedResult() {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].AddFootnoteProperties(WordNumberFormat.UpperLetter,
            WordFootnotePosition.PageBottom, WordNoteNumberRestart.EachPage, startNumber: 3);
        source.AddParagraph("FOOTREF").AddFootNote("FOOTBODY");
        var result = source.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeNotePageRestartApproximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Contains("FOOTREFC", Regex.Replace(pdf.GetPage(1).Text, @"\s+", ""));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LetterNotesAfterZUseWordsRepeatedLetterSequence(bool native) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].AddFootnoteProperties(WordNumberFormat.LowerLetter,
            WordFootnotePosition.PageBottom, WordNoteNumberRestart.Continuous, startNumber: 28);
        source.Sections[0].AddEndnoteProperties(WordNumberFormat.UpperLetter,
            WordEndnotePosition.DocumentEnd, WordNoteNumberRestart.Continuous, startNumber: 28);
        source.AddParagraph("FOOTREF").AddFootNote("FOOTBODY");
        source.AddParagraph("ENDREF").AddEndNote("ENDBODY");
        string text = ReadPdfText(source, native);
        Assert.Contains("FOOTREFbb", text);
        Assert.Contains("ENDREFBB", text);
        Assert.Contains("bbFOOTBODY", text);
        Assert.Contains("BBENDBODY", text);
    }
}
