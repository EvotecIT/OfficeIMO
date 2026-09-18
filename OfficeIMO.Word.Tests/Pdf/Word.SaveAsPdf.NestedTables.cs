using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using System.Text;
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
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
        AlternativeFormatImportPart embeddedPart = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<html><body>unsupported nested fragment</body></html>"))) {
            embeddedPart.FeedData(stream);
        }
        nested.Rows[0].Cells[0]._tableCell.Append(new AltChunk { Id = mainPart.GetIdOfPart(embeddedPart) });

        nested._table.Remove();
        Paragraph trailingParagraph = outerCell._tableCell.ChildElements.OfType<Paragraph>().Last();
        outerCell._tableCell.InsertBefore(
            new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Nested table host" }),
                new SdtContentBlock(nested._table)),
            trailingParagraph);
        outerCell.AddParagraph("After wrapped table");

        PdfDocumentConversionResult result = document.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Join(" ", pdf.GetPages().SelectMany(static page => page.GetWords()).Select(static word => word.Text));
        int beforeIndex = text.IndexOf("Before wrapped table", StringComparison.Ordinal);
        int nestedIndex = text.IndexOf("Wrapped nested evidence", StringComparison.Ordinal);
        int afterIndex = text.IndexOf("After wrapped table", StringComparison.Ordinal);
        Assert.True(beforeIndex >= 0, "Expected text before the content-control-wrapped table.");
        Assert.True(nestedIndex > beforeIndex, "Expected wrapped nested-table text after the preceding paragraph.");
        Assert.True(afterIndex > nestedIndex, "Expected the following paragraph after wrapped nested-table text.");
        Assert.Contains(result.Warnings, static warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.Contains(result.Warnings, static warning => warning.Code == "NativeBodyEmbeddedDocumentUnsupported");
    }

    [Fact]
    public void SaveAsPdf_NestedTable_UsesItsOwnStyleAndConditionalDefaults() {
        using WordDocument document = WordDocument.Create();
        const string styleId = "NestedTableRichTextStyle";
        const string sourceFamily = "OfficeIMO Nested Table Source";
        const string targetFamily = "OfficeIMO Nested Table Portable";
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Nested Table Rich Text Style" },
            new StyleRunProperties(
                new RunFonts { Ascii = sourceFamily, HighAnsi = sourceFamily },
                new Color { Val = "2255AA" },
                new FontSize { Val = "34" }),
            new TableStyleProperties(
                new RunPropertiesBaseStyle(new Bold())) {
                Type = TableStyleOverrideValues.FirstRow
            }) {
            Type = StyleValues.Table,
            StyleId = styleId,
            CustomStyle = true
        });

        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1);
        nested._tableProperties!.TableStyle = new TableStyle { Val = styleId };
        nested.ConditionalFormattingFirstRow = true;
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "AB";
        nested._table.Remove();
        WordTableCell outerCell = outer.Rows[0].Cells[0];
        Paragraph trailingParagraph = outerCell._tableCell.ChildElements.OfType<Paragraph>().Last();
        outerCell._tableCell.InsertBefore(
            new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Styled nested table host" }),
                new SdtContentBlock(nested._table)),
            trailingParagraph);

        var configured = new PdfOptions();
        configured
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                targetFamily,
                ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B')))
            .RegisterFontFamilySubstitution(
                sourceFamily,
                targetFamily,
                PdfFontFamilySubstitutionImpact.Compatible);

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false,
            FontFamily = "Helvetica",
            PdfOptions = configured,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        byte[] bytes = result.ToBytes();

        string raw = PdfOperatorSearchText.From(bytes);
        Assert.Contains("17 Tf", raw, StringComparison.Ordinal);
        Assert.Contains("0.133 0.333 0.667 rg", raw, StringComparison.Ordinal);
        PdfConversionWarning diagnostic = Assert.Single(result.Warnings, warning =>
            warning.Code == "NativeFontFamilySubstituted" &&
            warning.Details.TryGetValue("fontFamily", out string? family) &&
            family == sourceFamily);
        Assert.Equal(targetFamily, diagnostic.Details["resolvedFontFamily"]);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.All(
            pdf.GetPage(1).Letters.Where(static letter => letter.Value is "A" or "B"),
            letter => {
                Assert.Contains("OfficeIMO", letter.FontName, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase);
            });
    }

    [Fact]
    public void SaveAsPdf_NestedTable_SkipsMergedContinuationCellContent() {
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTable nested = outer.Rows[0].Cells[0].AddTable(3, 3, WordTableStyle.TableGrid);

        nested.Rows[0].Cells[0].Paragraphs[0].Text = "Horizontal anchor";
        nested.Rows[0].Cells[0].HorizontalMerge = WordCellMerge.Restart;
        WordParagraph hiddenHorizontal = nested.Rows[0].Cells[1].Paragraphs[0];
        hiddenHorizontal.Text = "HIDDEN-HORIZONTAL-CONTINUATION";
        hiddenHorizontal.AddPictureControl(imagePath, 24, 24, "Hidden continuation logo", "HiddenContinuationLogo");
        nested.Rows[0].Cells[1].HorizontalMerge = WordCellMerge.Continue;
        nested.Rows[0].Cells[2].Paragraphs[0].Text = "Horizontal tail";

        nested.Rows[1].Cells[0].Paragraphs[0].Text = "Vertical anchor";
        nested.Rows[1].Cells[0].VerticalMerge = WordCellMerge.Restart;
        nested.Rows[1].Cells[1].Paragraphs[0].Text = "Upper peer";
        WordParagraph hiddenVertical = nested.Rows[2].Cells[0].Paragraphs[0];
        hiddenVertical.Text = "HIDDEN-VERTICAL-CONTINUATION";
        hiddenVertical.AddCheckBox(true, "Hidden Continuation Check", "HiddenContinuationCheck");
        nested.Rows[2].Cells[0].VerticalMerge = WordCellMerge.Continue;
        nested.Rows[2].Cells[1].Paragraphs[0].Text = "Lower peer";

        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Join(" ", pdf.GetPages().SelectMany(static page => page.GetWords()).Select(static word => word.Text));
        Assert.Contains("Horizontal anchor", text, StringComparison.Ordinal);
        Assert.Contains("Horizontal tail", text, StringComparison.Ordinal);
        Assert.Contains("Vertical anchor", text, StringComparison.Ordinal);
        Assert.Contains("Upper peer", text, StringComparison.Ordinal);
        Assert.Contains("Lower peer", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-HORIZONTAL-CONTINUATION", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-VERTICAL-CONTINUATION", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Image", PdfOperatorSearchText.From(bytes), StringComparison.Ordinal);
        Assert.DoesNotContain(
            PdfInspector.Inspect(bytes).FormFields,
            static field => field.Name == "HiddenContinuationCheck");
    }

    [Fact]
    public void SaveAsPdf_HeaderNestedTable_PreservesTextAndImage() {
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        using WordDocument document = WordDocument.Create();
        WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
        WordTable outer = header.AddTable(1, 1, WordTableStyle.TableGrid);
        outer.Rows[0].Cells[0].Paragraphs[0].Text = "Header outer content";
        WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 2, WordTableStyle.TableGrid);
        WordParagraph nestedParagraph = nested.Rows[0].Cells[0].Paragraphs[0];
        nestedParagraph.Text = "Header nested evidence";
        nestedParagraph.AddPictureControl(imagePath, 32, 32, "Header Nested Logo", "HeaderNestedLogo");
        nested.Rows[0].Cells[0].HorizontalMerge = WordCellMerge.Restart;
        nested.Rows[0].Cells[1].Paragraphs[0].Text = "HIDDEN-HEADER-CONTINUATION";
        nested.Rows[0].Cells[1].HorizontalMerge = WordCellMerge.Continue;
        nested._table.Remove();
        WordTableCell outerCell = outer.Rows[0].Cells[0];
        Paragraph trailingParagraph = outerCell._tableCell.ChildElements.OfType<Paragraph>().Last();
        outerCell._tableCell.InsertBefore(
            new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Header nested table host" }),
                new SdtContentBlock(nested._table)),
            trailingParagraph);
        document.AddParagraph("Body evidence");

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false
        });
        byte[] bytes = result.ToBytes();

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Header nested evidence", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-HEADER-CONTINUATION", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Image", PdfOperatorSearchText.From(bytes), StringComparison.Ordinal);
        Assert.Contains(
            result.Warnings,
            static warning => warning.Code == "NativeNestedTableLayoutApproximated");
    }
}
