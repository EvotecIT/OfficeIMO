using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_NestedTable_PreservesDocumentOrderAndReportsFlattenedLayout() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTableCell outerCell = outer.Rows[0].Cells[0];
        outerCell.Paragraphs[0].Text = "Before nested table";
        WordTable nested = outerCell.AddTable(1, 1, WordTableStyle.TableGrid);
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "Nested table evidence";
        outerCell.AddParagraph("After nested table");

        PdfDocumentConversionResult result = document.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Join(" ", pdf.GetPages().SelectMany(static page => page.GetWords()).Select(static word => word.Text));
        int beforeIndex = text.IndexOf("Before nested table", StringComparison.Ordinal);
        int nestedIndex = text.IndexOf("Nested table evidence", StringComparison.Ordinal);
        int afterIndex = text.IndexOf("After nested table", StringComparison.Ordinal);
        Assert.True(beforeIndex >= 0, "Expected text before the nested table.");
        Assert.True(nestedIndex > beforeIndex, "Expected nested table text after the preceding paragraph.");
        Assert.True(afterIndex > nestedIndex, "Expected the following paragraph after the nested table text.");
        Assert.Contains(
            result.Warnings,
            static warning => warning.Code == "NativeNestedTableLayoutApproximated");
    }

    [Fact]
    public void SaveAsPdf_NestedTable_PreservesImagesAndFormControls() {
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableGrid);
        WordParagraph paragraph = nested.Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "Nested mixed content";
        paragraph.AddPictureControl(imagePath, 48, 48, "Nested Logo", "NestedLogo");
        paragraph.AddCheckBox(true, "Nested Approval", "NestedApproval");
        paragraph.AddDropDownList(new[] { "One", "Two" }, "Nested Choice", "NestedChoice");

        PdfDocumentConversionResult result = document.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();

        string pdfContent = PdfOperatorSearchText.From(bytes);
        Assert.Contains("/Subtype /Image", pdfContent, StringComparison.Ordinal);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Contains(info.FormFields, static field => field.Name == "NestedApproval" && field.IsCheckBox && field.Value == "Yes");
        Assert.Contains(info.FormFields, static field => field.Name == "NestedChoice" && field.IsChoiceField && field.Value == "One");
        Assert.Contains(
            result.Warnings,
            static warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.DoesNotContain(
            result.Warnings,
            static warning => warning.Code == "NativeBodyContentControlUnsupported");
    }

    [Fact]
    public void SaveAsPdf_ContentControlWrappedNestedTable_PreservesText() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTableCell outerCell = outer.Rows[0].Cells[0];
        outerCell.Paragraphs[0].Text = "Before wrapped table";
        WordTable nested = outerCell.AddTable(1, 1, WordTableStyle.TableGrid);
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "Wrapped nested evidence";

        nested._table.Remove();
        Paragraph trailingParagraph = outerCell._tableCell.ChildElements.OfType<Paragraph>().Last();
        outerCell._tableCell.InsertBefore(
            new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Nested table host" }),
                new SdtContentBlock(nested._table)),
            trailingParagraph);
        outerCell.AddParagraph("After wrapped table");

        byte[] bytes = document.ToPdfBytes();

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Join(" ", pdf.GetPages().SelectMany(static page => page.GetWords()).Select(static word => word.Text));
        int beforeIndex = text.IndexOf("Before wrapped table", StringComparison.Ordinal);
        int nestedIndex = text.IndexOf("Wrapped nested evidence", StringComparison.Ordinal);
        int afterIndex = text.IndexOf("After wrapped table", StringComparison.Ordinal);
        Assert.True(beforeIndex >= 0, "Expected text before the content-control-wrapped table.");
        Assert.True(nestedIndex > beforeIndex, "Expected wrapped nested-table text after the preceding paragraph.");
        Assert.True(afterIndex > nestedIndex, "Expected the following paragraph after wrapped nested-table text.");
    }

    [Fact]
    public void SaveAsPdf_HeaderNestedTable_PreservesTextAndImage() {
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        using WordDocument document = WordDocument.Create();
        WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
        WordTable outer = header.AddTable(1, 1, WordTableStyle.TableGrid);
        outer.Rows[0].Cells[0].Paragraphs[0].Text = "Header outer content";
        WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableGrid);
        WordParagraph nestedParagraph = nested.Rows[0].Cells[0].Paragraphs[0];
        nestedParagraph.Text = "Header nested evidence";
        nestedParagraph.AddPictureControl(imagePath, 32, 32, "Header Nested Logo", "HeaderNestedLogo");
        document.AddParagraph("Body evidence");

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false
        });
        byte[] bytes = result.ToBytes();

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Header nested evidence", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Image", PdfOperatorSearchText.From(bytes), StringComparison.Ordinal);
        Assert.Contains(
            result.Warnings,
            static warning => warning.Code == "NativeNestedTableLayoutApproximated");
    }
}
