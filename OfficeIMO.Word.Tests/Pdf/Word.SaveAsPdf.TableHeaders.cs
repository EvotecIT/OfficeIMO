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
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Header_Row() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRepeatingHeaderTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRepeatingHeaderTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(46, 2);
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].Text = "RepeatHdr";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "ValueHdr";
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;

            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) {
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Text = "Row " + rowIndex.ToString("D2");
                table.Rows[rowIndex].Cells[1].Paragraphs[0].Text = "Value " + rowIndex.ToString("D2");
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(260, 220),
                Margins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);

        int repeatedHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "RepeatHdr");
        Assert.True(repeatedHeaderCount >= 2);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Multiple_Header_Rows() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMultipleRepeatingHeaderRows.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMultipleRepeatingHeaderRows.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(44, 2);
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[1].RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "HdrA";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "HdrB";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "HdrC";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "HdrD";

            for (int rowIndex = 2; rowIndex < table.Rows.Count; rowIndex++) {
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Text = "Metric " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
                table.Rows[rowIndex].Cells[1].Paragraphs[0].Text = "Owner " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(260, 220),
                Margins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);

        int firstHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "HdrA");
        int secondHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "HdrC");

        Assert.True(firstHeaderCount >= 2);
        Assert.True(secondHeaderCount >= 2);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_First_Row_Style_Without_Repeating() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeFirstRowStyleNoRepeat.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeFirstRowStyleNoRepeat.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(44, 2, WordTableStyle.GridTable1Light);
            table.ConditionalFormattingFirstRow = true;
            table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = false;
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    cell.Width = 1440;
                    cell.WidthType = TableWidthUnitValues.Dxa;
                }
            }

            table.Rows[0].Cells[0].Paragraphs[0].Text = "SoloHdr";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "SoloValue";
            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) {
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Text = "Body " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
                table.Rows[rowIndex].Cells[1].Paragraphs[0].Text = "Value " + rowIndex.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(260, 220),
                Margins = PdfCore.PageMargins.Uniform(12)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);

        int firstRowHeaderCount = pdf.GetPages()
            .SelectMany(page => page.GetWords())
            .Count(word => word.Text == "SoloHdr");
        Assert.Equal(1, firstRowHeaderCount);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_First_Row_Conditional_Fill() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleFirstRowConditional.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleFirstRowConditional.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeFirstRowConditionalTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native First Row Conditional Table" },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(new Bold(), new Color { Val = "FFFFFF" }),
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "112233" }))
                { Type = TableStyleOverrideValues.FirstRow })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(2, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "ConditionalHdr";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "HeaderValue";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "BodyLabel";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "BodyValue";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("ConditionalHdr", text);
        Assert.Contains("BodyValue", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 rg", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Last_Row_Conditional_Fill() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleLastRowConditional.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleLastRowConditional.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeLastRowConditionalTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Last Row Conditional Table" },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(new Bold(), new Color { Val = "FFFFFF" }),
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "336699" }))
                { Type = TableStyleOverrideValues.LastRow })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(3, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingLastRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "HeaderLabel";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "HeaderValue";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "BodyLabel";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "BodyValue";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "TotalFooter";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "TotalValue";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 220),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BodyValue", text);
        Assert.Contains("TotalFooter", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.2 0.4 0.6 rg", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Row_Conditional_Rich_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRowConditionalRichText.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRowConditionalRichText.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeRowConditionalRichTextTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Row Conditional Rich Text Table" },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(
                        new Italic(),
                        new Color { Val = "2255AA" },
                        new FontSize { Val = "32" }))
                { Type = TableStyleOverrideValues.FirstRow },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(
                        new Bold(),
                        new Color { Val = "663399" },
                        new FontSize { Val = "28" }))
                { Type = TableStyleOverrideValues.LastRow })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(3, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstRow = true;
            table.ConditionalFormattingLastRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "RichHeader";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "RichHeaderValue";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "RichBody";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "RichBodyValue";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "RichFooter";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "RichFooterValue";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 240),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("RichHeader", text);
        Assert.Contains("RichFooterValue", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.True(
            raw.Contains("Helvetica-Oblique", StringComparison.Ordinal) ||
            raw.Contains("-Italic", StringComparison.Ordinal) ||
            raw.Contains("-Oblique", StringComparison.Ordinal),
            "Expected the first-row table style to preserve italic font selection.");
        Assert.Contains("16 Tf", raw);
        Assert.Contains("14 Tf", raw);
        Assert.Contains("0.133 0.333 0.667 rg", raw);
        Assert.Contains("0.4 0.2 0.6 rg", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Row_Conditional_Borders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRowConditionalBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRowConditionalBorders.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeRowConditionalBorderTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Row Conditional Border Table" },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new TableCellBorders(
                            new BottomBorder { Val = BorderValues.Single, Color = "112233", Size = 16U })))
                { Type = TableStyleOverrideValues.FirstRow },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new TableCellBorders(
                            new TopBorder { Val = BorderValues.Double, Color = "445566", Size = 12U })))
                { Type = TableStyleOverrideValues.LastRow })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(3, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstRow = true;
            table.ConditionalFormattingLastRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "BorderHeader";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "BorderHeaderValue";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "BorderBody";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "BorderBodyValue";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "BorderFooter";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "BorderFooterValue";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 240),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BorderHeader", text);
        Assert.Contains("BorderFooterValue", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 RG", raw);
        Assert.Contains("0.267 0.333 0.4 RG", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Column_And_Banding_Conditional_Borders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnBandBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnBandBorders.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeColumnBandBorderTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Column Band Border Table" },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new TableCellBorders(
                            new RightBorder { Val = BorderValues.Single, Color = "112233", Size = 8U })))
                { Type = TableStyleOverrideValues.FirstColumn },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new TableCellBorders(
                            new LeftBorder { Val = BorderValues.Double, Color = "445566", Size = 12U })))
                { Type = TableStyleOverrideValues.LastColumn },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new TableCellBorders(
                            new TopBorder { Val = BorderValues.Single, Color = "AA0000", Size = 8U },
                            new BottomBorder { Val = BorderValues.Single, Color = "AA0000", Size = 8U })))
                { Type = TableStyleOverrideValues.Band1Horizontal },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new TableCellBorders(
                            new LeftBorder { Val = BorderValues.Single, Color = "004488", Size = 12U },
                            new RightBorder { Val = BorderValues.Single, Color = "004488", Size = 12U })))
                { Type = TableStyleOverrideValues.Band1Vertical })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(4, 4);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstColumn = true;
            table.ConditionalFormattingLastColumn = true;
            table.ConditionalFormattingNoHorizontalBand = false;
            table.ConditionalFormattingNoVerticalBand = false;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                for (int columnIndex = 0; columnIndex < table.Rows[rowIndex].Cells.Count; columnIndex++) {
                    table.Rows[rowIndex].Cells[columnIndex].Paragraphs[0].Text =
                        "BorderCell" +
                        rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                        columnIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(440, 260),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BorderCell00", text);
        Assert.Contains("BorderCell33", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 RG", raw);
        Assert.Contains("0.267 0.333 0.4 RG", raw);
        Assert.Contains("0.667 0 0 RG", raw);
        Assert.Contains("0 0.267 0.533 RG", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_First_And_Last_Column_Conditional_Fills() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnConditional.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnConditional.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeColumnConditionalTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Column Conditional Table" },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "CCEEFF" }))
                { Type = TableStyleOverrideValues.FirstColumn },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "FFCC99" }))
                { Type = TableStyleOverrideValues.LastColumn })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(2, 3);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstColumn = true;
            table.ConditionalFormattingLastColumn = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "FirstColumnTop";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "MiddleTop";
            table.Rows[0].Cells[2].Paragraphs[0].Text = "LastColumnTop";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "FirstColumnBottom";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "MiddleBottom";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "LastColumnBottom";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 220),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("FirstColumnTop", text);
        Assert.Contains("LastColumnBottom", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.8 0.933 1 rg", raw);
        Assert.Contains("1 0.8 0.6 rg", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_First_And_Last_Column_Conditional_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnConditionalText.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnConditionalText.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeColumnConditionalTextTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Column Conditional Text Table" },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(new Bold(), new Color { Val = "112233" }))
                { Type = TableStyleOverrideValues.FirstColumn },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(new Bold(), new Color { Val = "445566" }))
                { Type = TableStyleOverrideValues.LastColumn })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(2, 3);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstColumn = true;
            table.ConditionalFormattingLastColumn = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "FirstTextTop";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "MiddleTextTop";
            table.Rows[0].Cells[2].Paragraphs[0].Text = "LastTextTop";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "FirstTextBottom";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "MiddleTextBottom";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "LastTextBottom";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 220),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("FirstTextTop", text);
        Assert.Contains("MiddleTextBottom", text);
        Assert.Contains("LastTextBottom", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 rg", raw);
        Assert.Contains("0.267 0.333 0.4 rg", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Banding_Conditional_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleBandingConditionalText.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleBandingConditionalText.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeBandingConditionalTextTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Banding Conditional Text Table" },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(new Bold(), new Color { Val = "AA0000" }))
                { Type = TableStyleOverrideValues.Band1Horizontal },
                new TableStyleProperties(
                    new RunPropertiesBaseStyle(new Bold(), new Color { Val = "004488" }))
                { Type = TableStyleOverrideValues.Band1Vertical })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(4, 3);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingNoHorizontalBand = false;
            table.ConditionalFormattingNoVerticalBand = false;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                for (int columnIndex = 0; columnIndex < table.Rows[rowIndex].Cells.Count; columnIndex++) {
                    table.Rows[rowIndex].Cells[columnIndex].Paragraphs[0].Text =
                        "BandText" +
                        rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                        columnIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 260),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BandText00", text);
        Assert.Contains("BandText31", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.667 0 0 rg", raw);
        Assert.Contains("0 0.267 0.533 rg", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Horizontal_Banding_Fill() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleHorizontalBanding.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleHorizontalBanding.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeHorizontalBandTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Horizontal Band Table" },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "99CCFF" }))
                { Type = TableStyleOverrideValues.Band1Horizontal })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(4, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingNoHorizontalBand = false;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Text = "BandLabel" + rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                table.Rows[rowIndex].Cells[1].Paragraphs[0].Text = "BandValue" + rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(360, 240),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BandLabel0", text);
        Assert.Contains("BandValue3", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.6 0.8 1 rg", raw);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Table_Style_Vertical_Banding_Fill() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleVerticalBanding.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleVerticalBanding.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeVerticalBandTable";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Vertical Band Table" },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "CC99FF" }))
                { Type = TableStyleOverrideValues.Band1Vertical })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(2, 4);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingNoVerticalBand = false;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                for (int columnIndex = 0; columnIndex < table.Rows[rowIndex].Cells.Count; columnIndex++) {
                    table.Rows[rowIndex].Cells[columnIndex].Paragraphs[0].Text =
                        "BandCell" +
                        rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                        columnIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 220),
                Margins = PdfCore.PageMargins.Uniform(30)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("BandCell00", text);
        Assert.Contains("BandCell13", text);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.8 0.6 1 rg", raw);
    }

}
