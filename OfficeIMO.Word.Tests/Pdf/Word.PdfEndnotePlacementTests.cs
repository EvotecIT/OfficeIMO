using System.Text.RegularExpressions;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordPdfEndnotePlacementTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void EndnotePositionControlsWhichSectionBodyPrecedesTheNote(bool native, bool sectionEnd) {
        using WordDocument source = CreateTwoSections(sectionEnd ? WordEndnotePosition.SectionEnd : WordEndnotePosition.DocumentEnd);
        string text = ReadPdfText(source, native);
        AssertOrder(text, sectionEnd ? "FIRSTENDNOTEBODY" : "FINALBODY", sectionEnd ? "FINALBODY" : "FIRSTENDNOTEBODY");
        AssertOrder(text, "FOOTNOTEBODY", "FINALBODY");
        AssertOrder(text, "FINALBODY", "LASTENDNOTEBODY");
        Assert.Contains("FIRSTREFERENCEi", text);
        Assert.Contains("LASTREFERENCEii", text);
        Assert.Contains("iFIRSTENDNOTEBODY", text);
        Assert.Contains("iiLASTENDNOTEBODY", text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultEndnotesFollowTheWholeDocument(bool native) {
        using WordDocument source = CreateTwoSections(null);
        string text = ReadPdfText(source, native);
        AssertOrder(text, "FINALBODY", "FIRSTENDNOTEBODY");
        AssertOrder(text, "FIRSTENDNOTEBODY", "LASTENDNOTEBODY");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void DocumentEndnotesFollowContinuousAndColumnFlows(bool native, bool columns) {
        using WordDocument source = CreateTwoSections(WordEndnotePosition.DocumentEnd, continuous: !columns, columns: columns);
        string text = ReadPdfText(source, native);
        AssertOrder(text, "FINALBODY", "FIRSTENDNOTEBODY");
        Assert.Contains("LASTENDNOTEBODY", text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DocumentWidePlacementAppliesWhenSectionsOmitTheirPosition(bool sectionEnd) {
        using WordDocument source = CreateTwoSections(null);
        W.Settings settings = source._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        W.EndnoteDocumentWideProperties? properties = settings.GetFirstChild<W.EndnoteDocumentWideProperties>();
        if (properties == null) { properties = new W.EndnoteDocumentWideProperties(); settings.AddChild(properties, true); }
        properties.RemoveAllChildren<W.EndnotePosition>();
        properties.AddChild(new W.EndnotePosition { Val = sectionEnd ? W.EndnotePositionValues.SectionEnd : W.EndnotePositionValues.DocumentEnd }, true);
        string text = ReadPdfText(source, false);
        AssertOrder(text, sectionEnd ? "FIRSTENDNOTEBODY" : "FINALBODY", sectionEnd ? "FINALBODY" : "FIRSTENDNOTEBODY");
    }

    private static WordDocument CreateTwoSections(WordEndnotePosition? position, bool continuous = false, bool columns = false) {
        WordDocument source = WordDocument.Create();
        void Configure(WordSection section) {
            if (position.HasValue) section.AddEndnoteProperties(WordNumberFormat.LowerRoman, position,
                WordNoteNumberRestart.Continuous, startNumber: 1);
        }
        Configure(source.Sections[0]);
        if (columns) source.Sections[0].ColumnCount = 2;
        source.AddParagraph("FIRSTREFERENCE").AddEndNote("FIRSTENDNOTEBODY");
        source.AddParagraph("FOOTREFERENCE").AddFootNote("FOOTNOTEBODY");
        WordSection last = continuous ? source.AddSection(WordSectionBreakType.Continuous) : source.AddSection();
        Configure(last);
        if (columns) last.ColumnCount = 2;
        last.AddParagraph("FINALBODY");
        last.AddParagraph("LASTREFERENCE").AddEndNote("LASTENDNOTEBODY");
        return source;
    }

    private static string ReadPdfText(WordDocument source, bool native) {
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(native ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = PdfPigDocument.Open(restored.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false, FontFamily = "Arial" }));
        return Regex.Replace(string.Concat(pdf.GetPages().Select(page => page.Text)), @"\s+", "");
    }

    private static void AssertOrder(string text, string first, string second) {
        Assert.Contains(first, text);
        Assert.Contains(second, text);
        Assert.True(text.IndexOf(first, StringComparison.Ordinal) < text.IndexOf(second, StringComparison.Ordinal));
    }
}
