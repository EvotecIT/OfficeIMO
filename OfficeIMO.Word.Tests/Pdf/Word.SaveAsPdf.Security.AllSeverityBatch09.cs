using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_ToleratesDuplicateNumberingAndStyleIds() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDuplicateNumberingAndStyles.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDuplicateNumberingAndStyles.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(new StyleName { Val = "First paragraph style" }) { Type = StyleValues.Paragraph, StyleId = "DuplicateParagraphStyle", CustomStyle = true },
                new Style(new StyleName { Val = "Second paragraph style" }) { Type = StyleValues.Paragraph, StyleId = "DuplicateParagraphStyle", CustomStyle = true },
                new Style(new StyleName { Val = "First character style" }) { Type = StyleValues.Character, StyleId = "DuplicateCharacterStyle", CustomStyle = true },
                new Style(new StyleName { Val = "Second character style" }) { Type = StyleValues.Character, StyleId = "DuplicateCharacterStyle", CustomStyle = true },
                new Style(new StyleName { Val = "First table style" }) { Type = StyleValues.Table, StyleId = "DuplicateTableStyle", CustomStyle = true },
                new Style(new StyleName { Val = "Second table style" }) { Type = StyleValues.Table, StyleId = "DuplicateTableStyle", CustomStyle = true });

            document.AddParagraph("Duplicate paragraph style survives").SetStyleId("DuplicateParagraphStyle");
            document.AddParagraph().AddText("Duplicate character style survives").SetCharacterStyleId("DuplicateCharacterStyle");
            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = new TableStyle { Val = "DuplicateTableStyle" };
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Duplicate table style survives";

            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering ??= new Numbering();
            numberingPart.Numbering.Append(
                new AbstractNum(
                    new Level(new StartNumberingValue { Val = 1 }) { LevelIndex = 0 },
                    new Level(new StartNumberingValue { Val = 2 }) { LevelIndex = 0 }) { AbstractNumberId = 42 },
                new AbstractNum(new Level { LevelIndex = 0 }) { AbstractNumberId = 42 },
                new NumberingInstance(
                    new AbstractNumId { Val = 42 },
                    new LevelOverride(new StartOverrideNumberingValue { Val = 3 }) { LevelIndex = 0 },
                    new LevelOverride(new StartOverrideNumberingValue { Val = 4 }) { LevelIndex = 0 }) { NumberID = 7 });

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions { IncludePageNumbers = false });
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_IgnoresMalformedLineHundredths() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMalformedLineHundredths.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMalformedLineHundredths.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph paragraph = document.AddParagraph("Malformed line spacing survives");
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            var spacing = new SpacingBetweenLines();
            spacing.SetAttribute(new OpenXmlAttribute(
                "w",
                "beforeLines",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "not-an-integer"));
            paragraph._paragraph.ParagraphProperties.Append(spacing);

            document.SaveAsPdf(pdfPath, new PdfSaveOptions { IncludePageNumbers = false });
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_ClampsExtremeTableCellParagraphIndents() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeExtremeTableCellIndents.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeExtremeTableCellIndents.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "Extreme cell indentation survives";
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.Indentation = new Indentation {
                Left = "2147483647",
                Right = "2147483647",
                FirstLine = "2147483647"
            };

            document.SaveAsPdf(pdfPath, new PdfSaveOptions { IncludePageNumbers = false });
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_AllowsConditionalRowFormattingAcrossVerticalMerge() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLastRowVerticalMerge.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLastRowVerticalMerge.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Merged across final row";
            table.Rows[0].Cells[0].MergeVertically(1);
            table.ConditionalFormattingFirstRow = true;
            table.ConditionalFormattingLastRow = true;

            document.SaveAsPdf(pdfPath, new PdfSaveOptions { IncludePageNumbers = false });
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_RejectsAggregateStyleInheritanceAmplification() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStyleInheritanceBudget.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStyleInheritanceBudget.pdf");

        using WordDocument document = WordDocument.Create(docPath);
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        string? parentStyleId = null;
        for (int index = 0; index < 128; index++) {
            string styleId = "SecurityStyle" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var style = new Style(new StyleName { Val = styleId }) {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            };
            if (parentStyleId != null) {
                style.Append(new BasedOn { Val = parentStyleId });
            }

            styles.Append(style);
            document.AddParagraph("Bounded style " + index.ToString(System.Globalization.CultureInfo.InvariantCulture)).SetStyleId(styleId);
            parentStyleId = styleId;
        }

        document.Save();
        Assert.Throws<InvalidDataException>(() => document.SaveAsPdf(pdfPath, new PdfSaveOptions { IncludePageNumbers = false }));
    }
}
