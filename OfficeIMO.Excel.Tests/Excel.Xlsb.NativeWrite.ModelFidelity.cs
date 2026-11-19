using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Model content that a workbook loaded from XLSX keeps when it is saved with the native XLSB writer.</summary>
public sealed class ExcelXlsbModelFidelityTests {
    [Fact]
    public void Vertical_alignment_stays_bottom_unless_the_format_sets_another() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Wrapped");
            sheet.Cell(1, 1).SetWrapText();
            Stylesheet styles = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
            styles.CellFormats!.Append(new CellFormat(new Alignment { Vertical = VerticalAlignmentValues.Top }) {
                FontId = 0U, FillId = 0U, BorderId = 0U, NumberFormatId = 0U, FormatId = 0U, ApplyAlignment = true
            });
            sheet.Cell(2, 1).SetValue("Top");
            sheet.SetCellStyleIndex(2, 1, (uint)styles.CellFormats.ChildElements.Count - 1, save: false);
        });

        byte[] xlsb = SaveAsXlsb(xlsx, out uint wrappedStyle, out uint topStyle);

        // MS-XLSB BrtXF alcV: 0 is top and 2 is bottom, the default when the format has no vertical value.
        byte[][] cellFormats = ReadCellFormatRecords(xlsb);
        Assert.Equal(2, (cellFormats[wrappedStyle][12] >> 3) & 0x07);
        Assert.Equal(0, (cellFormats[topStyle][12] >> 3) & 0x07);
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(xlsb));
        Assert.Null(reloaded["Data"].Cell(1, 1).GetStyle().VerticalAlignment);
        Assert.True(reloaded["Data"].Cell(1, 1).GetStyle().WrapText);
        Assert.Equal("top", reloaded["Data"].Cell(2, 1).GetStyle().VerticalAlignment);
    }

    [Fact]
    public void Built_in_and_custom_cell_styles_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            document.AddSheet("Data").Cell(1, 1).SetValue("Styled");
            Stylesheet styles = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
            styles.CellStyleFormats!.Append(
                new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U, NumberFormatId = 0U },
                new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U, NumberFormatId = 10U, ApplyNumberFormat = true });
            styles.CellStyles!.RemoveAllChildren();
            styles.CellStyles.Append(
                new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U },
                new CellStyle { Name = "Good", FormatId = 1U, BuiltinId = 26U },
                new CellStyle { Name = "Report total", FormatId = 2U, Hidden = true });
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        CellStyle[] cellStyles = reloaded.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.CellStyles!.Elements<CellStyle>().ToArray();
        Assert.Equal(new[] { "Normal", "Good", "Report total" }, cellStyles.Select(style => style.Name?.Value));
        Assert.Equal(new uint?[] { 0U, 1U, 2U }, cellStyles.Select(style => style.FormatId?.Value));
        Assert.Equal(new uint?[] { 0U, 26U, null }, cellStyles.Select(style => style.BuiltinId?.Value));
        Assert.True(cellStyles[2].Hidden?.Value);
    }

    [Fact]
    public void Formulas_referencing_other_sheets_ranges_of_sheets_and_names_save_natively() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet data = document.AddSheet("Data");
            ExcelSheet other = document.AddSheet("Other");
            ExcelSheet third = document.AddSheet("Third");
            data.Cell(1, 1).SetValue(1);
            other.Cell(1, 1).SetValue(2);
            third.Cell(1, 1).SetValue(3);
            document.SetNamedRange("Inputs", "Other!$A$1:$A$1");
            data.Cell(2, 1).SetFormula("Other!A1*2");
            data.Cell(3, 1).SetFormula("SUM(Data:Third!A1)");
            data.Cell(4, 1).SetFormula("SUM(Inputs)+Third!A1");
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        ExcelSheet sheet = reloaded["Data"];
        Assert.Equal("'Other'!A1*2", sheet.GetFormulaText(2, 1));
        Assert.Equal("SUM('Data:Third'!A1)", sheet.GetFormulaText(3, 1));
        Assert.Equal("SUM(Inputs)+'Third'!A1", sheet.GetFormulaText(4, 1));
        reloaded.RecalculateSupportedFormulas();
        Assert.True(sheet.TryGetCellText(2, 1, out string doubled));
        Assert.Equal("4", doubled);
    }

    [Fact]
    public void Sheet_view_settings_and_selection_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            document.AddSheet("Data").Cell(1, 1).SetValue(0);
            SheetView view = document["Data"].WorksheetPart.Worksheet!.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!;
            view.RemoveAllChildren();
            view.ShowZeros = false;
            view.RightToLeft = true;
            view.ShowRowColHeaders = false;
            view.View = SheetViewValues.PageLayout;
            view.TopLeftCell = "B3";
            view.ZoomScale = 85U;
            view.ZoomScaleNormal = 85U;
            view.ZoomScalePageLayoutView = 120U;
            view.Append(new Selection { ActiveCell = "C4", SequenceOfReferences = new ListValue<StringValue> { InnerText = "C4:D6" } });
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        SheetView actual = reloaded["Data"].WorksheetPart.Worksheet!.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!;
        Assert.False(actual.ShowZeros?.Value);
        Assert.True(actual.RightToLeft?.Value);
        Assert.False(actual.ShowRowColHeaders?.Value);
        Assert.Equal(SheetViewValues.PageLayout, actual.View?.Value);
        Assert.Equal("B3", actual.TopLeftCell?.Value);
        Assert.Equal(85U, actual.ZoomScale?.Value);
        Assert.Equal(85U, actual.ZoomScaleNormal?.Value);
        Assert.Equal(120U, actual.ZoomScalePageLayoutView?.Value);
        Selection selection = Assert.Single(actual.Elements<Selection>());
        Assert.Equal("C4", selection.ActiveCell?.Value);
        Assert.Equal("C4:D6", selection.SequenceOfReferences?.InnerText);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Rich_cell_text_runs_survive_native_xlsb_save(bool sharedString) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            var runs = new[] {
                new ExcelRichTextRun("Bold red") { Bold = true, Color = OfficeColor.Parse("#C00000") },
                new ExcelRichTextRun(" then "),
                new ExcelRichTextRun("italic") { Italic = true, FontFamily = "Arial", FontSize = OfficeLength.Points(14) }
            };
            sheet.Cell(1, 1).SetRichText(runs);
            if (sharedString) MoveInlineTextToSharedStrings(document, sheet);
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        IReadOnlyList<ExcelRichTextRun> actual = reloaded["Data"].Cell(1, 1).GetRichText();
        Assert.Equal(new[] { "Bold red", " then ", "italic" }, actual.Select(run => run.Text));
        Assert.True(actual[0].Bold);
        Assert.Equal(OfficeColor.Parse("#C00000"), actual[0].Color);
        Assert.False(actual[1].Bold || actual[1].Italic);
        Assert.True(actual[2].Italic);
        Assert.Equal("Arial", actual[2].FontFamily);
        Assert.Equal(OfficeLength.Points(14), actual[2].FontSize);
    }

    [Fact]
    public void Rich_text_runs_without_a_typeface_or_size_keep_the_cell_font() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetRichText(new ExcelRichTextRun("Bold") { Bold = true }, new ExcelRichTextRun(" plain"));
            sheet.Cell(1, 1).SetFontFamily("Arial");
            sheet.Cell(1, 1).SetFontSize(OfficeLength.Points(20));
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        ExcelRichTextRun bold = reloaded["Data"].Cell(1, 1).GetRichText()[0];
        Assert.True(bold.Bold);
        Assert.Equal("Arial", bold.FontFamily);
        Assert.Equal(OfficeLength.Points(20), bold.FontSize);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Rich_text_run_font_metadata_wins_over_the_inherited_cell_font(bool runSetsTypeface) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            var run = new ExcelRichTextRun("Greek") {
                Bold = true,
                FontFamily = runSetsTypeface ? "Times New Roman" : null,
                FontFamilyClassification = 3,
                FontCharacterSet = 161
            };
            sheet.Cell(1, 1).SetRichText(run, new ExcelRichTextRun(" plain"));
            sheet.Cell(1, 1).SetFontFamily("Arial");
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        ExcelRichTextRun actual = reloaded["Data"].Cell(1, 1).GetRichText()[0];
        Assert.Equal(runSetsTypeface ? "Times New Roman" : "Arial", actual.FontFamily);
        Assert.Equal((byte)3, actual.FontFamilyClassification);
        Assert.Equal((byte)161, actual.FontCharacterSet);
    }

    [Fact]
    public void A_shared_string_index_with_surrounding_spaces_keeps_its_runs() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetRichText(new ExcelRichTextRun("Bold") { Bold = true }, new ExcelRichTextRun(" plain"));
            MoveInlineTextToSharedStrings(document, sheet);
            Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
            cell.CellValue = new CellValue(" " + cell.CellValue!.Text + " ");
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        IReadOnlyList<ExcelRichTextRun> runs = reloaded["Data"].Cell(1, 1).GetRichText();
        Assert.Equal(new[] { "Bold", " plain" }, runs.Select(run => run.Text));
        Assert.True(runs[0].Bold);
    }

    [Fact]
    public void A_text_literal_that_looks_like_an_external_reference_does_not_block_native_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            document.AddSheet("Other").Cell(1, 1).SetValue(2);
            sheet.Cell(1, 1).SetFormula("\"'[Book.xlsx]Sheet1'!A1\"&Other!A1");
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx, out _, out _)));

        Assert.Equal("\"'[Book.xlsx]Sheet1'!A1\"&'Other'!A1", reloaded["Data"].GetFormulaText(1, 1));
    }

    [Fact]
    public void A_cell_style_with_extensions_is_refused() {
        byte[] xlsx = CreateXlsx(document => {
            document.AddSheet("Data").Cell(1, 1).SetValue(1);
            CellStyle normal = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.CellStyles!.Elements<CellStyle>().First();
            normal.Append(new ExtensionList(new Extension { Uri = "{00000000-0000-0000-0000-000000000001}" }));
        });

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        NotSupportedException refused = Assert.Throws<NotSupportedException>(() =>
            loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("extensions on cell style", refused.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Run_edits_with_unchanged_text_in_a_loaded_xlsb_are_refused(bool sourceHasRuns) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            if (sourceHasRuns) sheet.Cell(1, 1).SetRichText(new ExcelRichTextRun("Bold") { Bold = true }, new ExcelRichTextRun(" plain"));
            else sheet.Cell(1, 1).SetValue("Bold plain");
            sheet.Cell(2, 1).SetRichText(new ExcelRichTextRun("Kept") { Italic = true }, new ExcelRichTextRun(" runs"));
        });
        byte[] xlsb = SaveAsXlsb(xlsx, out _, out _);

        using (ExcelDocument unchanged = ExcelDocument.Load(new MemoryStream(xlsb))) {
            unchanged["Data"].Cell(3, 1).SetValue("edit elsewhere");
            using ExcelDocument resaved = ExcelDocument.Load(new MemoryStream(unchanged.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb })));
            Assert.True(resaved["Data"].Cell(2, 1).GetRichText()[0].Italic);
        }

        using ExcelDocument edited = ExcelDocument.Load(new MemoryStream(xlsb));
        edited["Data"].Cell(1, 1).SetRichText(new ExcelRichTextRun("Bold") { Italic = true }, new ExcelRichTextRun(" plain"));
        NotSupportedException refused = Assert.Throws<NotSupportedException>(() =>
            edited.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("rich text runs", refused.Message);
    }

    [Fact]
    public void Projected_rich_text_runs_count_against_the_cell_limit() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            for (int row = 1; row <= 4; row++) {
                sheet.Cell(row, 1).SetRichText(new ExcelRichTextRun("A") { Bold = true }, new ExcelRichTextRun("B"), new ExcelRichTextRun("C") { Italic = true });
            }
        });
        byte[] xlsb = SaveAsXlsb(xlsx, out _, out _);

        using (ExcelDocument.Load(new MemoryStream(xlsb), new ExcelLoadOptions { MaxCells = 16 })) { }
        InvalidDataException rejected = Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.Load(new MemoryStream(xlsb), new ExcelLoadOptions { MaxCells = 15 }));
        Assert.Contains("text runs", rejected.Message);
    }

    private static byte[] CreateXlsx(Action<ExcelDocument> build) {
        using ExcelDocument document = ExcelDocument.Create();
        build(document);
        return document.ToBytes();
    }

    private static byte[] SaveAsXlsb(byte[] xlsx, out uint firstStyle, out uint secondStyle) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        Cell[] cells = loaded.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().ToArray();
        firstStyle = cells.ElementAtOrDefault(0)?.StyleIndex?.Value ?? 0U;
        secondStyle = cells.ElementAtOrDefault(1)?.StyleIndex?.Value ?? 0U;
        return loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static byte[][] ReadCellFormatRecords(byte[] xlsb) {
        using var archive = new ZipArchive(new MemoryStream(xlsb), ZipArchiveMode.Read);
        using var styles = new MemoryStream();
        using (Stream entry = archive.GetEntry("xl/styles.bin")!.Open()) entry.CopyTo(styles);
        styles.Position = 0;
        var records = XlsbRecordReader.ReadAll(styles);
        int begin = records.ToList().FindIndex(record => record.Type == 617); // BrtBeginCellXFs
        return records.Skip(begin + 1).TakeWhile(record => record.Type == 47).Select(record => record.Data).ToArray();
    }

    private static void MoveInlineTextToSharedStrings(ExcelDocument document, ExcelSheet sheet) {
        Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single(item => item.InlineString != null);
        SharedStringTablePart part = document.WorkbookPartRoot.SharedStringTablePart ?? document.WorkbookPartRoot.AddNewPart<SharedStringTablePart>();
        part.SharedStringTable ??= new SharedStringTable();
        part.SharedStringTable.Append(new SharedStringItem(cell.InlineString!.Elements<Run>().Select(run => run.CloneNode(true))));
        int index = part.SharedStringTable.Elements<SharedStringItem>().Count() - 1;
        cell.InlineString = null;
        cell.DataType = CellValues.SharedString;
        cell.CellValue = new CellValue(index.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }
}
