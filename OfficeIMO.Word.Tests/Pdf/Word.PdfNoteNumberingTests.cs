using System.Text.RegularExpressions;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordPdfNoteNumberingTests {
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
}
