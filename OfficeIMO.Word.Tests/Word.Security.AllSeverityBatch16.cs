using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordAllSeverityBatch16SecurityTests {
    [Fact]
    public void InvalidCellMarginWidthsDoNotEscapeAsParseExceptions() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.AddTableCellProperties();
        cell._tableCellProperties!.TableCellMargin = new TableCellMargin(
            new TopMargin { Width = "not-a-number" },
            new BottomMargin { Width = "999999999999" });

        Assert.Null(cell.MarginTopWidth);
        Assert.Null(cell.MarginBottomWidth);
    }

    [Fact]
    public void FluentCustomPropertyTreatsNullAsEmptyText() {
        using WordDocument document = WordDocument.Create();

        document.AsFluent().Info(info => info.Custom("Optional", null));

        Assert.Equal(string.Empty, document.CustomDocumentProperties["Optional"].Value);
    }

    [Fact]
    public void DuplicateCaseInsensitiveStyleIdsDoNotCrashReferenceFieldUpdates() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.AddItem("Target").AddBookmark("Target");
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(
            new Style { Type = StyleValues.Paragraph, StyleId = "DuplicateStyle" },
            new Style { Type = StyleValues.Paragraph, StyleId = "duplicatestyle" });
        document.AddParagraph("Reference: ")._paragraph.Append(BuildSimpleField(" REF Target \\n ", "stale"));

        Exception? exception = Record.Exception(() => document.UpdateFieldsAndGetReport());

        Assert.Null(exception);
    }

    [Fact]
    public void DuplicateNumberingIdsAndLevelsDoNotCrashReferenceFieldUpdates() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.AddItem("Target").AddBookmark("Target");
        Numbering numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        AbstractNum original = numbering.Elements<AbstractNum>().First();
        int abstractId = original.AbstractNumberId!.Value;
        original.Append(new Level { LevelIndex = 0 });
        numbering.Append(new AbstractNum(new Level { LevelIndex = 0 }) { AbstractNumberId = abstractId });
        document.AddParagraph("Reference: ")._paragraph.Append(BuildSimpleField(" REF Target \\n ", "stale"));

        Exception? exception = Record.Exception(() => document.UpdateFieldsAndGetReport());

        Assert.Null(exception);
    }

    [Fact]
    public void OversizedListPlaceholderIsPreservedWithoutThrowing() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Numbered);
        list.AddItem("Item");
        LevelText levelText = document._wordprocessingDocument.MainDocumentPart!
            .NumberingDefinitionsPart!.Numbering!.Descendants<LevelText>().First();
        levelText.Val = "%999999999999999999999";

        Dictionary<WordParagraph, (int Level, string Marker)> markers = DocumentTraversal.BuildListMarkers(document);

        Assert.Equal("%999999999999999999999", Assert.Single(markers).Value.Marker);
    }

    [Fact]
    public void TypedFieldsAcceptEmptyOptionalParameters() {
        using WordDocument document = WordDocument.Create();

        Exception? exception = Record.Exception(() => {
            document.AddField(new FileNameField());
            document.AddField(new AskField());
        });

        Assert.Null(exception);
        Assert.Equal(2, document.Fields.Count);
    }

    [Fact]
    public void RepeatingSectionMetadataCannotInjectMarkup() {
        const string hostile = "value' /><w:tag w:val='injected";
        using WordDocument document = WordDocument.Create();

        WordRepeatingSection section = document.AddParagraph().AddRepeatingSection(hostile, hostile, hostile);
        XElement xml = XElement.Parse(section._sdtRun.OuterXml);

        Assert.Equal(hostile, section.Alias);
        Assert.Equal(hostile, section.Tag);
        Assert.Single(xml.Descendants(), element => element.Name.LocalName == "tag");
        Assert.Equal(hostile, xml.Descendants().Single(element => element.Name.LocalName == "repeatingSection")
            .Attributes().Single(attribute => attribute.Name.LocalName == "sectionTitle").Value);
    }

    [Fact]
    public void VerticalMergeLeavesUnevenRowsUntouched() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2);
        table.Rows[1].Cells[1].Remove();

        Exception? exception = Record.Exception(() => table.Rows[0].MergeVertically(1, 1));

        Assert.Null(exception);
        Assert.Null(table.Rows[0].Cells[1].VerticalMerge);
    }

    [Fact]
    public void EqualParagraphsHaveEqualHashCodes() {
        using WordDocument firstDocument = WordDocument.Create();
        using WordDocument secondDocument = WordDocument.Create();
        WordParagraph first = firstDocument.AddParagraph("Same text");
        WordParagraph second = secondDocument.AddParagraph("Same text");

        Assert.Equal(first, second);
        Assert.Equal(first.GetHashCode(), second.GetHashCode());
    }

    [Fact]
    public void ExternalImagesAcceptWebUrisAndRejectLocalResourceSchemes() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();

        paragraph.AddImage(new Uri("https://example.test/image.png"), 10, 10);
        Assert.Null(Record.Exception(() => document.ToPdf()));
        Assert.Throws<ArgumentException>(() =>
            paragraph.AddImage(new Uri("file:///private/secret.png"), 10, 10));
        Assert.Throws<ArgumentException>(() =>
            paragraph.AddImage(new Uri("smb://server/share/image.png"), 10, 10));
        Assert.Throws<ArgumentNullException>(() => paragraph.AddImage(null!, 10, 10));
    }

    [Fact]
    public void EquationMarkupRejectsDocumentTypeDeclarationsBeforeOpenXmlParsing() {
        using WordDocument document = WordDocument.Create();
        const string omml = "<!DOCTYPE x [<!ENTITY boom 'expanded'>]><m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:r><m:t>&boom;</m:t></m:r></m:oMath>";

        Assert.Throws<System.Xml.XmlException>(() => document.AddParagraph().AddEquation(omml));
    }

    [Fact]
    public void PdfRenderDoesNotReuseRowsRemovedAfterEarlierRender() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 1);
        table.Rows[0].Cells[0].Paragraphs[0].SetText("visible");
        table.Rows[1].Cells[0].Paragraphs[0].SetText("removed-secret");
        _ = document.ToPdf();

        table.Rows[1].Remove();
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));

        Assert.Contains("visible", text, StringComparison.Ordinal);
        Assert.DoesNotContain("removed-secret", text, StringComparison.Ordinal);
    }

    private static SimpleField BuildSimpleField(string instruction, string resultText) =>
        new SimpleField(new Run(new Text(resultText))) { Instruction = instruction };
}
