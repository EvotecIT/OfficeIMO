using DocumentFormat.OpenXml.Wordprocessing;
using System;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Row_Height() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableRowHeight.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableRowHeight.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            table.Rows[0].Height = 1600;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "TallRow";
            document.AddParagraph("AfterTall");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 500),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var tableWord = Assert.Single(words, word => word.Text == "TallRow");
        var followingWord = Assert.Single(words, word => word.Text == "AfterTall");

        Assert.True(tableWord.BoundingBox.Bottom > followingWord.BoundingBox.Top + 45D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_NonUniform_Row_Heights() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableNonUniformRowHeights.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableNonUniformRowHeights.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(3, 1);
            table.Rows[0].Height = 400;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "ShortA";
            table.Rows[1].Height = 1200;
            table.Rows[1].Cells[0].Paragraphs[0].Text = "TallB";
            table.Rows[2].Height = 400;
            table.Rows[2].Cells[0].Paragraphs[0].Text = "ShortC";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(320, 260),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var shortA = Assert.Single(words, word => word.Text == "ShortA");
        var tallB = Assert.Single(words, word => word.Text == "TallB");
        var shortC = Assert.Single(words, word => word.Text == "ShortC");

        double firstGap = shortA.BoundingBox.Bottom - tallB.BoundingBox.Bottom;
        double secondGap = tallB.BoundingBox.Bottom - shortC.BoundingBox.Bottom;
        Assert.True(secondGap > firstGap + 35D, $"Expected non-uniform Word row height to push the third row down. ShortA/TallB gap: {firstGap}; TallB/ShortC gap: {secondGap}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_SpacingOnly_Empty_Paragraph_After_Table() {
        double compactGap = RenderNativeAfterTableSpacingGap("PdfNativeAfterTableNoSpacingParagraph", includeSpacingParagraph: false);
        double spacedGap = RenderNativeAfterTableSpacingGap("PdfNativeAfterTableSpacingParagraph", includeSpacingParagraph: true);

        Assert.True(spacedGap > compactGap + 4D, $"Expected a spacing-only empty Word paragraph after a table to move following content down. Gap without spacer: {compactGap:0.##}; gap with spacer: {spacedGap:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Table_Cell_NoWrap_Text() {
        double wrappedGap = RenderNativeTableCellWrapTextGap("PdfNativeTableCellWrapText", wrapText: true);
        double noWrapGap = RenderNativeTableCellWrapTextGap("PdfNativeTableCellNoWrapText", wrapText: false);

        Assert.True(wrappedGap > noWrapGap + 16D, $"Expected Word no-wrap table cell text to avoid vertical wrapping in native PDF output. Wrapped gap: {wrappedGap:0.##}; no-wrap gap: {noWrapGap:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Expands_Percentage_Width_Table_To_Content_Frame() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativePercentageWidthTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativePercentageWidthTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 2);
            table.WidthType = TableWidthUnitValues.Pct;
            table.Width = 5000;
            table.Rows[0].Cells[0].Width = 1440;
            table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
            table.Rows[0].Cells[1].Width = 1440;
            table.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "LeftColumn";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "RightColumn";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 300),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var left = Assert.Single(words, word => word.Text == "LeftColumn");
        var right = Assert.Single(words, word => word.Text == "RightColumn");

        Assert.InRange(left.BoundingBox.Left, 42D, 55D);
        Assert.True(right.BoundingBox.Left > 195D, $"Expected a 100% Word table to expand to the content frame. RightColumn left: {right.BoundingBox.Left}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Autofit_For_Percentage_Width_Word_Table() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAutofitPercentageTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAutofitPercentageTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 7);
            table.WidthType = TableWidthUnitValues.Pct;
            table.Width = 5000;
            table.LayoutType = TableLayoutValues.Autofit;
            string[] headers = {
                "Item",
                "IssuedOn",
                "Category",
                "Reference",
                "ClausePath",
                "UpdatedOn",
                "Status"
            };
            string[] values = {
                "Invoice | Review 10",
                "05/26/2023 09:00:56",
                "commercial.contract",
                "253cfd36-2f82-4672-b8e3-31b7a8ebaaf4",
                "Section=Revenue,Article=LateFee,Clause={253CFD36-2F82-4672-B8E3-31B7A8EBAAF4},Page=12,Paragraph=4,Region=Global",
                "05/26/2023 09:00:56",
                "Allow"
            };

            for (int column = 0; column < headers.Length; column++) {
                table.Rows[0].Cells[column].Width = 2400;
                table.Rows[0].Cells[column].WidthType = TableWidthUnitValues.Dxa;
                table.Rows[0].Cells[column].Paragraphs[0].Text = headers[column];
                table.Rows[1].Cells[column].Width = 2400;
                table.Rows[1].Cells[column].WidthType = TableWidthUnitValues.Dxa;
                table.Rows[1].Cells[column].Paragraphs[0].Text = values[column];
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 300),
                Margins = new PdfCore.PageMargins(72, 72, 40, 40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var clausePath = Assert.Single(words, word => word.Text == "ClausePath");
        var status = Assert.Single(words, word => word.Text == "Status");

        Assert.True(clausePath.BoundingBox.Left < 345D, $"Expected Word autofit to move the structured path column left of equal-grid placement based on value shape. Left: {clausePath.BoundingBox.Left}.");
        Assert.True(status.BoundingBox.Left > 455D, $"Expected Word autofit to reserve a wide structured path column based on value shape. Status left: {status.BoundingBox.Left}.");
    }

    private double RenderNativeAfterTableSpacingGap(string fileNamePrefix, bool includeSpacingParagraph) {
        string tableMarker = fileNamePrefix + "Table";
        string afterMarker = fileNamePrefix + "After";
        string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
        string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = tableMarker;

            if (includeSpacingParagraph) {
                WordParagraph blank = document.AddParagraph();
                blank.LineSpacingAfterPoints = 6;
            }

            WordParagraph after = document.AddParagraph(afterMarker);
            after.LineSpacingAfterPoints = 0;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(36),
                FontFamily = "Helvetica"
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        double tableY = Assert.Single(words, word => word.Text == tableMarker).BoundingBox.Bottom;
        double afterY = Assert.Single(words, word => word.Text == afterMarker).BoundingBox.Bottom;
        return tableY - afterY;
    }

    private double RenderNativeTableCellWrapTextGap(string fileNamePrefix, bool wrapText) {
        const string tableMarker = "Start";
        const string afterMarker = "After";
        string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
        string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            table.Width = 900;
            table.WidthType = TableWidthUnitValues.Dxa;
            table.LayoutType = TableLayoutValues.Fixed;
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Width = 900;
            cell.WidthType = TableWidthUnitValues.Dxa;
            cell.WrapText = wrapText;
            cell.Paragraphs[0].Text = tableMarker + " Alpha Beta Gamma Delta Epsilon";

            WordParagraph after = document.AddParagraph(afterMarker);
            after.LineSpacingAfterPoints = 0;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(36),
                FontFamily = "Helvetica"
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        double tableY = Assert.Single(words, word => word.Text == tableMarker).BoundingBox.Bottom;
        double afterY = Assert.Single(words, word => word.Text == afterMarker).BoundingBox.Bottom;
        return tableY - afterY;
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_DocDefaults_For_Unstyled_Table_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableDocDefaults.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableDocDefaults.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            DocDefaults docDefaults = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!;
            RunPropertiesBaseStyle runDefaults = docDefaults.GetFirstChild<RunPropertiesDefault>()!.GetFirstChild<RunPropertiesBaseStyle>()!;
            runDefaults.GetFirstChild<FontSize>()!.Val = "28";
            runDefaults.GetFirstChild<FontSizeComplexScript>()!.Val = "28";

            WordTable table = document.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "InheritedTableSize";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        string content = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Matches(@"/F\d+\s+14\s+Tf", content);

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Contains("InheritedTableSize", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Table_Style_Run_Properties_For_Cell_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRunProperties.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRunProperties.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeTableRunProperties";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Table Run Properties" },
                new StyleRunProperties(
                    new Color { Val = "C00000" },
                    new FontSize { Val = "28" }))
            {
                Type = StyleValues.Table,
                StyleId = styleId,
                CustomStyle = true
            });

            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.Rows[0].Cells[0].Paragraphs[0].Text = "StyledTableRun";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(40),
                FontFamily = "Helvetica"
            });
        }

        string content = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Matches(@"/F\d+\s+14\s+Tf", content);
        Assert.Contains("0.753 0 0 rg", content);

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Contains("StyledTableRun", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Table_Style_Shading_For_Cell_Fills() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleShading.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleShading.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeTableStyleShading";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Table Style Shading" },
                new StyleTableProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "D9EAD3" }))
            {
                Type = StyleValues.Table,
                StyleId = styleId,
                CustomStyle = true
            });

            WordTable table = document.AddTable(1, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.Rows[0].Cells[0].Paragraphs[0].Text = "InheritedFill";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "DirectFill";
            table.Rows[0].Cells[1].ShadingFillColorHex = "F4CCCC";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 220),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        string content = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.851 0.918 0.827 rg", content);
        Assert.Contains("0.957 0.8 0.8 rg", content);

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("InheritedFill", text);
        Assert.Contains("DirectFill", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_DocDefaults_For_Table_Cell_Paragraph_Spacing() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellParagraphSpacing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellParagraphSpacing.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            DocDefaults docDefaults = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!;
            ParagraphPropertiesBaseStyle paragraphDefaults = docDefaults.GetFirstChild<ParagraphPropertiesDefault>()!.GetFirstChild<ParagraphPropertiesBaseStyle>()!;
            paragraphDefaults.GetFirstChild<SpacingBetweenLines>()!.After = "600";

            WordTable table = document.AddTable(2, 1);
            table._tableProperties!.TableStyle?.Remove();
            table.Rows[0].Cells[0].Paragraphs[0].Text = "CellSpacingOne";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "CellSpacingTwo";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        var first = Assert.Single(words, word => word.Text == "CellSpacingOne");
        var second = Assert.Single(words, word => word.Text == "CellSpacingTwo");

        double rowBaselineGap = first.BoundingBox.Bottom - second.BoundingBox.Bottom;
        Assert.True(rowBaselineGap > 40D, $"Expected Word doc-default paragraph spacing to increase native PDF table row height. Gap: {rowBaselineGap}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Paragraph_Style_For_Table_Cell_Paragraph_Spacing() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellParagraphStyleSpacing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellParagraphStyleSpacing.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "CellParagraphSpacingStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Cell Paragraph Spacing Style" },
                new StyleParagraphProperties(new SpacingBetweenLines { After = "600" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordTable table = document.AddTable(2, 1);
            table._tableProperties!.TableStyle?.Remove();
            table.Rows[0].Cells[0].Paragraphs[0].Text = "StyledCellSpacingOne";
            table.Rows[0].Cells[0].Paragraphs[0].SetStyleId(styleId);
            table.Rows[1].Cells[0].Paragraphs[0].Text = "StyledCellSpacingTwo";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        var first = Assert.Single(words, word => word.Text == "StyledCellSpacingOne");
        var second = Assert.Single(words, word => word.Text == "StyledCellSpacingTwo");

        double rowBaselineGap = first.BoundingBox.Bottom - second.BoundingBox.Bottom;
        Assert.True(rowBaselineGap > 40D, $"Expected Word paragraph style spacing to increase native PDF table row height. Gap: {rowBaselineGap}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Collapses_Table_Cell_Adjacent_Paragraph_Spacing() {
        double compactGap = RenderNativeTableCellParagraphSpacingGap("PdfNativeTableCellCompactParagraphSpacing", firstSpacingAfter: 0D, secondSpacingBefore: 0D);
        double beforeOnlyGap = RenderNativeTableCellParagraphSpacingGap("PdfNativeTableCellParagraphSpacingBeforeOnly", firstSpacingAfter: 0D, secondSpacingBefore: 20D);
        double afterOnlyGap = RenderNativeTableCellParagraphSpacingGap("PdfNativeTableCellParagraphSpacingAfterOnly", firstSpacingAfter: 30D, secondSpacingBefore: 0D);
        double collapsedGap = RenderNativeTableCellParagraphSpacingGap("PdfNativeTableCellParagraphSpacingCollapsed", firstSpacingAfter: 30D, secondSpacingBefore: 20D);

        Assert.True(beforeOnlyGap > compactGap + 8D, $"Expected direct Word spacing-before to increase stacked table cell paragraph distance. Compact gap: {compactGap:0.##}; before-only gap: {beforeOnlyGap:0.##}.");
        Assert.InRange(Math.Abs(collapsedGap - afterOnlyGap), 0D, 2D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Table_Cell_Paragraph_Contextual_Spacing() {
        double spacedGap = RenderNativeTableCellContextualParagraphSpacingGap("PdfNativeTableCellContextualOff", contextualSpacing: false);
        double contextualGap = RenderNativeTableCellContextualParagraphSpacingGap("PdfNativeTableCellContextualOn", contextualSpacing: true);

        Assert.True(spacedGap > contextualGap + 16D, $"Expected Word contextual spacing to suppress spacing between same-style table cell paragraphs. Spaced gap: {spacedGap:0.##}; contextual gap: {contextualGap:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Does_Not_Invent_Table_Cell_Paragraph_Spacing_When_Undeclared() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellNoImplicitParagraphSpacing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellNoImplicitParagraphSpacing.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 1);
            table._tableProperties!.TableStyle?.Remove();
            table.Rows[0].Cells[0].Paragraphs[0].Text = "CompactCellOne";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "CompactCellTwo";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        var first = Assert.Single(words, word => word.Text == "CompactCellOne");
        var second = Assert.Single(words, word => word.Text == "CompactCellTwo");

        double rowBaselineGap = first.BoundingBox.Bottom - second.BoundingBox.Bottom;
        Assert.InRange(rowBaselineGap, 12D, 30D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_TableGrid_Style_Spacing_For_Row_Pitch() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableGridRowPitch.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableGridRowPitch.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(3, 1, WordTableStyle.TableGrid);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Alpha";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Beta";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "Gamma";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(72)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        var alpha = Assert.Single(words, word => word.Text == "Alpha");
        var beta = Assert.Single(words, word => word.Text == "Beta");
        var gamma = Assert.Single(words, word => word.Text == "Gamma");

        double firstGap = alpha.BoundingBox.Bottom - beta.BoundingBox.Bottom;
        double secondGap = beta.BoundingBox.Bottom - gamma.BoundingBox.Bottom;
        Assert.InRange(firstGap, 13D, 18D);
        Assert.InRange(secondGap, 13D, 18D);
    }

    private double RenderNativeTableCellParagraphSpacingGap(string fileNamePrefix, double firstSpacingAfter, double secondSpacingBefore) {
        const string firstMarker = "CellFirst";
        const string secondMarker = "CellSecond";
        string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
        string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle?.Remove();
            table.Rows[0].Cells[0].Width = 2880;
            table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = firstMarker;
            cell.Paragraphs[0].LineSpacingAfterPoints = firstSpacingAfter;
            WordParagraph second = cell.AddParagraph(secondMarker);
            second.LineSpacingBeforePoints = secondSpacingBefore;
            second.LineSpacingAfterPoints = 0;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(40),
                FontFamily = "Helvetica"
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        double firstY = Assert.Single(words, word => word.Text == firstMarker).BoundingBox.Bottom;
        double secondY = Assert.Single(words, word => word.Text == secondMarker).BoundingBox.Bottom;
        return firstY - secondY;
    }

    private double RenderNativeTableCellContextualParagraphSpacingGap(string fileNamePrefix, bool contextualSpacing) {
        const string firstMarker = "CellContextFirst";
        const string secondMarker = "CellContextSecond";
        string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
        string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "CellContextualSpacingStyle";
            var styleParagraphProperties = new StyleParagraphProperties(
                new SpacingBetweenLines { After = "600" });
            if (contextualSpacing) {
                styleParagraphProperties.Append(new ContextualSpacing());
            }

            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Cell Contextual Spacing Style" },
                styleParagraphProperties)
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle?.Remove();
            table.Rows[0].Cells[0].Width = 2880;
            table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = firstMarker;
            cell.Paragraphs[0].SetStyleId(styleId);
            cell.AddParagraph(secondMarker).SetStyleId(styleId);

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(40),
                FontFamily = "Helvetica"
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        double firstY = Assert.Single(words, word => word.Text == firstMarker).BoundingBox.Bottom;
        double secondY = Assert.Single(words, word => word.Text == secondMarker).BoundingBox.Bottom;
        return firstY - secondY;
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Hyperlink() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLinkedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLinkedTable.pdf");
        const string linkUri = "https://evotec.xyz/native-table-link";

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            WordParagraph linkParagraph = cell.Paragraphs[0].AddHyperLink("Native table link", new Uri(linkUri), addStyle: true, tooltip: "Native table link metadata");
            linkParagraph.Hyperlink!.Anchor = "IgnoredWhenExternalUriExists";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native table link", text);
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal(linkUri, link.Uri);
        Assert.Equal("Native table link metadata", link.Contents);
        Assert.Equal(new[] { linkUri }, info.LinkUris);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 3);
            WordTableCell leftCell = table.Rows[0].Cells[0];
            WordTableCell centerCell = table.Rows[0].Cells[1];
            WordTableCell rightCell = table.Rows[0].Cells[2];

            leftCell.Width = 1440;
            leftCell.WidthType = TableWidthUnitValues.Dxa;
            centerCell.Width = 1440;
            centerCell.WidthType = TableWidthUnitValues.Dxa;
            rightCell.Width = 1440;
            rightCell.WidthType = TableWidthUnitValues.Dxa;

            leftCell.Paragraphs[0].Text = "TOP";
            leftCell.AddParagraph("PAD");
            leftCell.AddParagraph("PAD");
            leftCell.Paragraphs[0].ParagraphAlignment = JustificationValues.Left;
            leftCell.Paragraphs[1].ParagraphAlignment = JustificationValues.Left;
            leftCell.Paragraphs[2].ParagraphAlignment = JustificationValues.Left;
            leftCell.VerticalAlignment = TableVerticalAlignmentValues.Top;

            centerCell.Paragraphs[0].Text = "MID";
            centerCell.Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
            centerCell.VerticalAlignment = TableVerticalAlignmentValues.Center;

            rightCell.Paragraphs[0].Text = "END";
            rightCell.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
            rightCell.VerticalAlignment = TableVerticalAlignmentValues.Bottom;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var top = Assert.Single(words, word => word.Text == "TOP");
        var mid = Assert.Single(words, word => word.Text == "MID");
        var end = Assert.Single(words, word => word.Text == "END");

        const double columnWidth = 72D;
        double firstColumnLeft = top.BoundingBox.Left - 4D;
        double secondColumnLeft = firstColumnLeft + columnWidth;
        double thirdColumnLeft = secondColumnLeft + columnWidth;

        Assert.InRange(top.BoundingBox.Left, firstColumnLeft + 3D, firstColumnLeft + 8D);
        Assert.InRange(mid.BoundingBox.Left, secondColumnLeft + 20D, secondColumnLeft + 36D);
        Assert.InRange(end.BoundingBox.Right, thirdColumnLeft + columnWidth - 8D, thirdColumnLeft + columnWidth - 2D);
        Assert.True(top.BoundingBox.Bottom > mid.BoundingBox.Bottom + 8D);
        Assert.True(mid.BoundingBox.Bottom > end.BoundingBox.Bottom + 8D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Mixed_Paragraph_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMixedParagraphAlignedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMixedParagraphAlignedTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Width = 2880;
            cell.WidthType = TableWidthUnitValues.Dxa;
            cell.Paragraphs[0].Text = "LeftMix";
            cell.Paragraphs[0].ParagraphAlignment = JustificationValues.Left;
            WordParagraph center = cell.AddParagraph("CenterMix");
            center.ParagraphAlignment = JustificationValues.Center;
            WordParagraph right = cell.AddParagraph("RightMix");
            right.ParagraphAlignment = JustificationValues.Right;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var left = Assert.Single(words, word => word.Text == "LeftMix");
        var centerWord = Assert.Single(words, word => word.Text == "CenterMix");
        var rightWord = Assert.Single(words, word => word.Text == "RightMix");

        double cellLeft = left.BoundingBox.Left - 4D;
        double cellRight = cellLeft + 144D;
        double cellCenter = (cellLeft + cellRight) / 2D;

        Assert.InRange(left.BoundingBox.Left, cellLeft + 3D, cellLeft + 8D);
        Assert.InRange((centerWord.BoundingBox.Left + centerWord.BoundingBox.Right) / 2D, cellCenter - 8D, cellCenter + 8D);
        Assert.InRange(rightWord.BoundingBox.Right, cellRight - 8D, cellRight - 2D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_Style_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStyleAlignedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStyleAlignedTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string centerStyleId = "NativeTableCenterStyle";
            const string rightStyleId = "NativeTableRightStyle";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Native Table Center Style" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new Justification {
                        Val = JustificationValues.Center
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = centerStyleId,
                    CustomStyle = true
                },
                new Style(
                    new StyleName { Val = "Native Table Right Style" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new Justification {
                        Val = JustificationValues.Right
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = rightStyleId,
                    CustomStyle = true
                });

            WordTable table = document.AddTable(1, 3);
            WordTableCell leftCell = table.Rows[0].Cells[0];
            WordTableCell centerCell = table.Rows[0].Cells[1];
            WordTableCell rightCell = table.Rows[0].Cells[2];

            leftCell.Width = 1440;
            leftCell.WidthType = TableWidthUnitValues.Dxa;
            centerCell.Width = 1440;
            centerCell.WidthType = TableWidthUnitValues.Dxa;
            rightCell.Width = 1440;
            rightCell.WidthType = TableWidthUnitValues.Dxa;

            leftCell.Paragraphs[0].Text = "SLEFT";
            centerCell.Paragraphs[0].Text = "SMID";
            centerCell.Paragraphs[0].SetStyleId(centerStyleId);
            rightCell.Paragraphs[0].Text = "SEND";
            rightCell.Paragraphs[0].SetStyleId(rightStyleId);

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var left = Assert.Single(words, word => word.Text == "SLEFT");
        var center = Assert.Single(words, word => word.Text == "SMID");
        var right = Assert.Single(words, word => word.Text == "SEND");

        const double columnWidth = 72D;
        double firstColumnLeft = left.BoundingBox.Left - 4D;
        double secondColumnLeft = firstColumnLeft + columnWidth;
        double thirdColumnLeft = secondColumnLeft + columnWidth;

        Assert.InRange(left.BoundingBox.Left, firstColumnLeft + 3D, firstColumnLeft + 8D);
        Assert.InRange(center.BoundingBox.Left, secondColumnLeft + 20D, secondColumnLeft + 36D);
        Assert.InRange(right.BoundingBox.Right, thirdColumnLeft + columnWidth - 8D, thirdColumnLeft + columnWidth - 2D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Cell_NonUniform_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeNonUniformAlignedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeNonUniformAlignedTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 2);
            foreach (WordTableRow row in table.Rows) {
                row.Height = 1100;
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].Text = "TopPeer";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Left2";
            table.Rows[0].Cells[1].Paragraphs[0].ParagraphAlignment = JustificationValues.Left;
            table.Rows[0].Cells[1].VerticalAlignment = TableVerticalAlignmentValues.Top;

            table.Rows[1].Cells[0].Paragraphs[0].Text = "TopCell";
            table.Rows[1].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Top;
            table.Rows[1].Cells[1].Paragraphs[0].Text = "R2";
            table.Rows[1].Cells[1].Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
            table.Rows[1].Cells[1].VerticalAlignment = TableVerticalAlignmentValues.Bottom;

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 260),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var leftPeer = Assert.Single(words, word => word.Text == "Left2");
        var topCell = Assert.Single(words, word => word.Text == "TopCell");
        var rightBottom = Assert.Single(words, word => word.Text == "R2");

        Assert.True(rightBottom.BoundingBox.Left > leftPeer.BoundingBox.Left + 35D, $"Expected non-uniform right-aligned cell to move right. Left2 x: {leftPeer.BoundingBox.Left}; R2 x: {rightBottom.BoundingBox.Left}.");
        Assert.True(topCell.BoundingBox.Bottom > rightBottom.BoundingBox.Bottom + 20D, $"Expected non-uniform bottom-aligned cell to move down inside the same row. TopCell bottom: {topCell.BoundingBox.Bottom}; R2 bottom: {rightBottom.BoundingBox.Bottom}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Merged_Cells() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMergedTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMergedTable.pdf");
        const string horizontalUri = "https://evotec.xyz/native-table-column-span";
        const string verticalUri = "https://evotec.xyz/native-table-row-span";

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(3, 3);
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].AddHyperLink("Across", new Uri(horizontalUri), addStyle: true, tooltip: "Column span metadata");
            table.Rows[0].Cells[0].Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
            table.Rows[0].Cells[2].Paragraphs[0].Text = "TopTail";

            table.Rows[1].Cells[0].Paragraphs[0].AddHyperLink("Tall", new Uri(verticalUri), addStyle: true, tooltip: "Row span metadata");
            table.Rows[1].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Center;
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Upper";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "UpperTail";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "Lower";
            table.Rows[2].Cells[2].Paragraphs[0].Text = "LowerTail";

            table.Rows[0].Cells[0].MergeHorizontally(1);
            table.Rows[1].Cells[0].MergeVertically(1);

            Assert.Equal(2, table.Rows[0].Cells[0].ColumnSpan);
            Assert.Equal(2, table.Rows[1].Cells[0].RowSpan);

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Across", text);
            Assert.Contains("Tall", text);
            Assert.Contains("TopTail", text);
            Assert.Contains("Upper", text);
            Assert.Contains("Lower", text);
        }

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfLinkAnnotation horizontal = Assert.Single(info.LinkAnnotations, link => link.Uri == horizontalUri);
        PdfCore.PdfLinkAnnotation vertical = Assert.Single(info.LinkAnnotations, link => link.Uri == verticalUri);
        Assert.Equal("Column span metadata", horizontal.Contents);
        Assert.Equal("Row span metadata", vertical.Contents);
        Assert.True(horizontal.Width > 110D);
        Assert.True(vertical.Height > 30D);
    }
}
