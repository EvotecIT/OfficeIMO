using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Write;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Xlsb_RecordReader_FramesCanonicalWorkbookBoundaryRecords() {
            byte[] bytes = {
                0x83, 0x01, 0x00, // BrtBeginBook (131), empty payload
                0x84, 0x01, 0x00  // BrtEndBook (132), empty payload
            };

            using var stream = new MemoryStream(bytes, writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(stream);

            Assert.Collection(records,
                begin => {
                    Assert.Equal(0, begin.Offset);
                    Assert.Equal(3, begin.HeaderSize);
                    Assert.Equal(131, begin.Type);
                    Assert.Empty(begin.Data);
                },
                end => {
                    Assert.Equal(3, end.Offset);
                    Assert.Equal(3, end.HeaderSize);
                    Assert.Equal(132, end.Type);
                    Assert.Empty(end.Data);
                });
        }

        [Fact]
        public void Xlsb_RecordReader_DecodesMultiByteTypeAndSize() {
            byte[] bytes = new byte[204];
            bytes[0] = 0xFD;
            bytes[1] = 0x04; // BrtCommentText (637)
            bytes[2] = 0xC8;
            bytes[3] = 0x01; // 200 payload bytes
            for (int index = 4; index < bytes.Length; index++) {
                bytes[index] = (byte)(index - 4);
            }

            using var stream = new MemoryStream(bytes, writable: false);
            XlsbRecord record = Assert.Single(XlsbRecordReader.ReadAll(stream));

            Assert.Equal(637, record.Type);
            Assert.Equal(200, record.Size);
            Assert.Equal(4, record.HeaderSize);
            Assert.Equal(0, record.Data[0]);
            Assert.Equal(199, record.Data[199]);
        }

        [Fact]
        public void Xlsb_RecordWriter_ProducesCanonicalReaderCompatibleFraming() {
            byte[] payload = Enumerable.Range(0, 200).Select(index => (byte)index).ToArray();
            using var stream = new MemoryStream();

            XlsbRecordWriter.Write(stream, 637, payload);

            Assert.Equal(204, stream.Length);
            stream.Position = 0;
            XlsbRecord record = Assert.Single(XlsbRecordReader.ReadAll(stream));
            Assert.Equal(637, record.Type);
            Assert.Equal(payload, record.Data);
        }

        [Fact]
        public void Xlsb_FormulaReader_DecodesUnicodeStringsErrorsAndAttributeControlledIf() {
            byte[] ifTokens = {
                0x1D, 0x01,
                0x19, 0x02, 0x00, 0x00,
                0x17, 0x03, 0x00, (byte)'Y', 0x00, (byte)'e', 0x00, (byte)'s', 0x00,
                0x19, 0x08, 0x00, 0x00,
                0x17, 0x02, 0x00, (byte)'N', 0x00, (byte)'o', 0x00,
                0x19, 0x08, 0x00, 0x00,
                0x22, 0x03, 0x01, 0x00
            };
            byte[] concatenationTokens = {
                0x17, 0x03, 0x00, (byte)'A', 0x00, (byte)'"', 0x00, (byte)'B', 0x00,
                0x17, 0x01, 0x00, (byte)'!', 0x00,
                0x08
            };

            Assert.True(XlsbFormulaTextReader.TryRead(ifTokens, out string? conditional));
            Assert.True(XlsbFormulaTextReader.TryRead(concatenationTokens, out string? concatenation));
            Assert.True(XlsbFormulaTextReader.TryRead(new byte[] { 0x1C, 0x2A }, out string? error));
            Assert.True(XlsbFormulaTextReader.TryRead(new byte[] { 0x1E, 0x01, 0x00, 0x19, 0x10, 0x00, 0x00 }, out string? sum));
            Assert.Equal("IF(TRUE,\"Yes\",\"No\")", conditional);
            Assert.Equal("\"A\"\"B\"&\"!\"", concatenation);
            Assert.Equal("#N/A", error);
            Assert.Equal("SUM(1)", sum);
        }

        [Theory]
        [InlineData(new byte[] { 0x83 })]
        [InlineData(new byte[] { 0x83, 0x01, 0x80 })]
        [InlineData(new byte[] { 0x01, 0x02, 0xAA })]
        public void Xlsb_RecordReader_RejectsTruncatedRecords(byte[] bytes) {
            using var stream = new MemoryStream(bytes, writable: false);

            Assert.Throws<EndOfStreamException>(() => XlsbRecordReader.ReadAll(stream));
        }

        [Fact]
        public void Xlsb_RecordReader_EnforcesAllocationLimitBeforeReadingPayload() {
            using var stream = new MemoryStream(new byte[] { 0x01, 0x02, 0xAA, 0xBB }, writable: false);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                XlsbRecordReader.ReadAll(stream, maxRecordBytes: 1));

            Assert.Contains("configured limit of 1 byte", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_RecordReader_EnforcesWorkbookWideRecordBudget() {
            byte[] bytes = {
                0x83, 0x01, 0x00,
                0x84, 0x01, 0x00
            };
            using var stream = new MemoryStream(bytes, writable: false);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                XlsbRecordReader.ReadAll(stream, budget: new XlsbRecordReadBudget(1)));

            Assert.Contains("BIFF12 records", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_FormulaEncoder_RejectsDeepControlFunctionNesting() {
            string formula = "1";
            for (int index = 0; index < 66; index++) {
                formula = $"IF(TRUE,{formula},0)";
            }

            Assert.False(XlsbFormulaEncoder.TryEncode(
                formula, out _, out string? reason));
            Assert.Contains("nesting", reason, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_FormulaEncoder_RejectsDeepLegacyParserNesting() {
            string formula = new string('(', 65) + "1" + new string(')', 65);

            Assert.False(XlsbFormulaEncoder.TryEncode(
                formula, out _, out string? reason));
            Assert.Contains("nesting", reason, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_FormulaEncoder_RejectsExcessiveLegacyOperatorChains() {
            string postfix = "1" + new string('%', 300);
            string binary = string.Join("+", Enumerable.Repeat("1", 300));

            Assert.False(XlsbFormulaEncoder.TryEncode(postfix, out _, out string? postfixReason));
            Assert.False(XlsbFormulaEncoder.TryEncode(binary, out _, out string? binaryReason));
            Assert.Contains("nesting", postfixReason, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("nesting", binaryReason, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_WorkbookReader_EnforcesRowMetadataBudget() {
            byte[] package;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Rows");
                sheet.CellValue(1, 1, "one");
                sheet.CellValue(2, 1, "two");
                package = document.ToBytes(ExcelFileFormat.Xlsb);
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                XlsbWorkbookReader.Load(package,
                    new XlsbImportOptions { MaxRowDefinitions = 1 }));

            Assert.Contains("row definitions", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_PackageDetector_UsesPackageContentInsteadOfExtension() {
            byte[] package = CreateMinimalXlsbPackage();

            Assert.True(XlsbPackageDetector.TryFindWorkbookPart(package, out string? workbookPart));
            Assert.Equal("xl/workbook.bin", workbookPart);
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(package, "misleading.xlsx"));
        }

        [Fact]
        public void Xlsb_PackageDetector_DoesNotMisclassifyXmlWorkbookPackages() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data").CellValue(1, 1, "XML package");
            byte[] package = document.ToBytes();

            Assert.False(XlsbPackageDetector.TryFindWorkbookPart(package, out _));
            Assert.Equal(ExcelFileFormat.Xlsx, ExcelDocumentLoadRouting.DetectFormat(package, "misleading.xlsb"));
        }

        [Fact]
        public void Xlsb_PackageReader_RejectsOversizedUnreferencedPartsBeforeProjection() {
            byte[] package = AddZipEntry(CreateMinimalXlsbPackage(), "xl/media/unreferenced.bin", new byte[2_048]);
            var options = new XlsbImportOptions {
                MaxPartBytes = 1_024,
                MaxPackageBytes = 64 * 1_024
            };

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                XlsbWorkbookReader.Load(package, options));

            Assert.Contains("unreferenced.bin", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("configured limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_ExcelGeneratedFixture_HasCanonicalPackageAndWorkbookRecords() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "basic-values-formula.xlsb");
            byte[] package = File.ReadAllBytes(path);

            Assert.True(XlsbPackageDetector.TryFindWorkbookPart(package, out string? workbookPart));
            Assert.Equal("xl/workbook.bin", workbookPart);
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(package, path));

            using var packageStream = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
            ZipArchiveEntry workbookEntry = Assert.Single(
                archive.Entries,
                entry => string.Equals(entry.FullName, workbookPart, StringComparison.OrdinalIgnoreCase));
            using Stream workbookStream = workbookEntry.Open();
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(workbookStream);

            Assert.NotEmpty(records);
            Assert.Equal(131, records[0].Type); // BrtBeginBook
            Assert.Equal(132, records[records.Count - 1].Type); // BrtEndBook
        }

        [Fact]
        public void Xlsb_ExcelGeneratedFixture_LoadsThroughNormalExcelDocumentSurface() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "basic-values-formula.xlsb");

            using ExcelDocument document = ExcelDocument.Load(path);

            Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);
            Assert.Equal(path, document.SourcePath);
            ExcelSheet sheet = Assert.Single(document.Sheets);
            Assert.Equal("Arkusz1", sheet.Name);
            Assert.True(sheet.TryGetCellText(1, 1, out string? a1));
            Assert.True(sheet.TryGetCellText(1, 2, out string? b1));
            Assert.True(sheet.TryGetCellText(2, 1, out string? a2));
            Assert.True(sheet.TryGetCellText(2, 2, out string? b2));
            Assert.True(sheet.TryGetCellText(3, 2, out string? b3));
            Assert.Equal("Name", a1);
            Assert.Equal("Amount", b1);
            Assert.Equal("Alpha", a2);
            Assert.Equal("42", b2);
            Assert.Equal("50", b3);

            Cell formulaCell = Assert.Single(
                sheet.DeferredMetadataWorksheetPart.Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "B3");
            Assert.Equal("SUM(B2,8)", formulaCell.CellFormula?.Text);
            CalculationProperties calculation = Assert.IsType<CalculationProperties>(
                document.WorkbookRoot.GetFirstChild<CalculationProperties>());
            Assert.Equal(181029U, calculation.CalculationId?.Value);
            Assert.Equal(CalculateModeValues.Auto, calculation.CalculationMode?.Value);
            Assert.Equal(100U, calculation.IterateCount?.Value);
            Assert.Equal(0.001D, calculation.IterateDelta?.Value);
            Assert.Equal(ReferenceModeValues.A1, calculation.ReferenceMode?.Value);
            Assert.True(calculation.FullPrecision?.Value);
            Assert.True(calculation.CalculationCompleted?.Value);
            Assert.True(calculation.CalculationOnSave?.Value);
            Assert.True(calculation.ConcurrentCalculation?.Value);
            Assert.Null(calculation.ConcurrentManualCount);
            Assert.NotEmpty(document.XlsbPreservedRecords);
            Assert.Contains(document.XlsbImportDiagnostics, diagnostic => diagnostic.Code == "XLSB-RECORDS-PRESERVED");
        }

        [Fact]
        public void Xlsb_NormalLoad_ReadOnlyModeUsesReadOnlyPackageAndRejectsSave() {
            using ExcelDocument document = ExcelDocument.Load(
                GetExcelGeneratedXlsbFixturePath(),
                new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });

            Assert.Equal(OfficeIMO.Drawing.DocumentAccessMode.ReadOnly, document.AccessMode);
            Assert.Equal(FileAccess.Read, document.FileOpenAccess);
            Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);
            Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? value));
            Assert.Equal("42", value);

            using var destination = new MemoryStream();
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                document.Save(destination, ExcelFileFormat.Xlsb));
            Assert.Contains("read-only", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, destination.Length);
        }

        [Fact]
        public void Xlsb_StyledExcelFixture_ProjectsDateSystemFormatsAndCellStyles() {
            using ExcelDocument document = ExcelDocument.Load(GetStyledExcelGeneratedXlsbFixturePath());

            Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);
            Assert.Equal(ExcelDateSystem.NineteenFour, document.DateSystem);
            ExcelSheet sheet = Assert.Single(document.Sheets);
            Assert.Equal("StylesDates", sheet.Name);

            Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
            Assert.Equal(3U, stylesheet.Fonts!.Count!.Value);
            Assert.Equal(4U, stylesheet.Fills!.Count!.Value);
            Assert.Equal(3U, stylesheet.Borders!.Count!.Value);
            Assert.Equal(1U, stylesheet.CellStyleFormats!.Count!.Value);
            Assert.Equal(6U, stylesheet.CellFormats!.Count!.Value);
            Assert.Contains(stylesheet.NumberingFormats!.Elements<NumberingFormat>(), format =>
                format.NumberFormatId?.Value == 164U && format.FormatCode?.Value == "yyyy\\-mm\\-dd");
            Assert.Contains(stylesheet.NumberingFormats!.Elements<NumberingFormat>(), format =>
                format.NumberFormatId?.Value == 165U && format.FormatCode?.Value == "0.0000");

            ExcelCellValueSnapshot date = AssertCellValue(sheet, 2, 1);
            Assert.Equal(ExcelCellValueKind.DateTime, date.Kind);
            Assert.Equal(new DateTime(2024, 2, 29), date.DateTimeValue);
            Assert.Equal(2U, sheet.GetCellStyle(2, 1).StyleIndex);
            Assert.Equal("yyyy\\-mm\\-dd", sheet.GetCellStyle(2, 1).NumberFormatCode);

            ExcelCellStyleSnapshot heading = sheet.GetCellStyle(1, 1);
            Assert.Equal(1U, heading.StyleIndex);
            Assert.True(heading.Bold);
            Assert.Equal("solid", heading.FillPatternType);
            Assert.NotNull(heading.Border);

            ExcelCellStyleSnapshot percent = sheet.GetCellStyle(3, 1);
            Assert.Equal(3U, percent.StyleIndex);
            Assert.Equal(10U, percent.NumberFormatId);

            ExcelCellStyleSnapshot precise = sheet.GetCellStyle(4, 1);
            Assert.Equal(4U, precise.StyleIndex);
            Assert.Equal("0.0000", precise.NumberFormatCode);

            ExcelCellStyleSnapshot decorated = sheet.GetCellStyle(5, 1);
            Assert.Equal(5U, decorated.StyleIndex);
            Assert.Equal("solid", decorated.FillPatternType);
            Assert.Equal("center", decorated.HorizontalAlignment);
            Assert.Equal(15, decorated.TextRotation);
            Assert.True(decorated.WrapText);
            Assert.NotNull(decorated.Border);

            Cell formulaCell = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Cell>(),
                cell => string.Equals(cell.CellReference?.Value, "B2", StringComparison.Ordinal));
            Assert.Equal("SUM(A3,A4)", formulaCell.CellFormula?.Text);
            Assert.True(sheet.TryGetCellText(2, 2, out string? cachedFormulaValue));
            Assert.Equal("1234.6928", cachedFormulaValue);
            Assert.Equal(4U, formulaCell.StyleIndex?.Value);

            ExcelCellValueSnapshot boolean = AssertCellValue(sheet, 3, 2);
            Assert.Equal(ExcelCellValueKind.Boolean, boolean.Kind);
            Assert.Equal("1", boolean.RawValue);
        }

        [Fact]
        public void Xlsb_StyledExcelFixture_NativeRewritePreservesStylesDatesAndFormulaPayload() {
            byte[] original = File.ReadAllBytes(GetStyledExcelGeneratedXlsbFixturePath());
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(original, writable: false));
            document.Sheets[0].CellValue(3, 1, 0.2D);

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            ExcelSheet sheet = Assert.Single(reloaded.Sheets);
            Assert.Equal(ExcelDateSystem.NineteenFour, reloaded.DateSystem);
            Assert.Equal(3U, sheet.GetCellStyle(3, 1).StyleIndex);
            Assert.Equal(10U, sheet.GetCellStyle(3, 1).NumberFormatId);
            Assert.Equal(new DateTime(2024, 2, 29), AssertCellValue(sheet, 2, 1).DateTimeValue);
            Assert.True(sheet.TryGetCellText(3, 1, out string? percentValue));
            Assert.Equal("0.2", percentValue);
            Cell formulaCell = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Cell>(),
                cell => string.Equals(cell.CellReference?.Value, "B2", StringComparison.Ordinal));
            Assert.Equal("SUM(A3,A4)", formulaCell.CellFormula?.Text);
            AssertFormulaPayloadEqual(original, rewritten, "xl/worksheets/sheet1.bin", (2, 2));
            AssertPackageEntriesEqualExcept(original, rewritten, "xl/worksheets/sheet1.bin");
        }

        [Fact]
        public void Xlsb_StyledExcelFixture_ConvertsToEditableXlsxWithFormattingIntact() {
            using ExcelDocument source = ExcelDocument.Load(GetStyledExcelGeneratedXlsbFixturePath());
            using var destination = new MemoryStream();

            source.Save(destination, ExcelFileFormat.Xlsx);

            using ExcelDocument converted = ExcelDocument.Load(new MemoryStream(destination.ToArray(), writable: false));
            ExcelSheet sheet = Assert.Single(converted.Sheets);
            Assert.Equal(ExcelDateSystem.NineteenFour, converted.DateSystem);
            Assert.Equal(new DateTime(2024, 2, 29), AssertCellValue(sheet, 2, 1).DateTimeValue);
            Assert.True(sheet.GetCellStyle(1, 1).Bold);
            Assert.Equal("solid", sheet.GetCellStyle(5, 1).FillPatternType);
            Assert.Equal(15, sheet.GetCellStyle(5, 1).TextRotation);
            Assert.Equal("SUM(A3,A4)", sheet.CellAt(2, 2).GetValue().Formula);
        }

        [Fact]
        public void Xlsb_GeometryFixture_ProjectsDimensionsRowsColumnsPanesAndMerges() {
            using ExcelDocument document = ExcelDocument.Load(GetGeometryExcelGeneratedXlsbFixturePath());

            ExcelSheet sheet = Assert.Single(document.Sheets);
            Assert.Equal("Geometry", sheet.Name);
            Worksheet worksheet = sheet.WorksheetPart.Worksheet;
            Assert.Equal("A1:D5", worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value);

            SheetFormatProperties format = Assert.IsType<SheetFormatProperties>(worksheet.GetFirstChild<SheetFormatProperties>());
            Assert.Equal(9D, format.DefaultColumnWidth?.Value);
            Assert.Equal(18D, format.DefaultRowHeight?.Value);
            Assert.True(format.CustomHeight?.Value);
            Assert.Equal((byte)2, format.OutlineLevelRow?.Value);
            Assert.Equal((byte)1, format.OutlineLevelColumn?.Value);

            Column[] columns = Assert.IsType<Columns>(worksheet.GetFirstChild<Columns>()).Elements<Column>().ToArray();
            Assert.Equal(4, columns.Length);
            AssertColumn(columns[0], 1, 1, 18D, hidden: false, outlineLevel: 0, collapsed: false);
            AssertColumn(columns[1], 2, 2, 12D, hidden: false, outlineLevel: 1, collapsed: false);
            AssertColumn(columns[2], 3, 3, 12D, hidden: false, outlineLevel: 1, collapsed: true);
            AssertColumn(columns[3], 4, 4, 8D, hidden: true, outlineLevel: 0, collapsed: false);

            IReadOnlyDictionary<int, ExcelRowSnapshot> rows = sheet.GetRowDefinitions().ToDictionary(row => row.Index);
            Assert.Equal(3, rows.Count);
            Assert.Equal(30D, rows[1].Height);
            Assert.True(rows[1].CustomHeight);
            Assert.True(rows[3].Hidden);
            Assert.Equal((byte)2, rows[4].OutlineLevel);
            Assert.True(rows[4].Collapsed);

            ExcelMergedRangeSnapshot merge = Assert.Single(sheet.GetMergedRanges());
            Assert.Equal("A1:C1", merge.A1Range);

            Pane pane = Assert.IsType<Pane>(worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>());
            Assert.Equal(1D, pane.HorizontalSplit?.Value);
            Assert.Equal(1D, pane.VerticalSplit?.Value);
            Assert.Equal("B2", pane.TopLeftCell?.Value);
            Assert.Equal(PaneValues.BottomRight, pane.ActivePane?.Value);
            Assert.Equal(PaneStateValues.FrozenSplit, pane.State?.Value);
        }

        [Fact]
        public void Xlsb_GeometryFixture_NativeRewritePreservesWorksheetMetadata() {
            byte[] original = File.ReadAllBytes(GetGeometryExcelGeneratedXlsbFixturePath());
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(original, writable: false));
            document.Sheets[0].CellValue(2, 1, "Edited");

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            ExcelSheet sheet = Assert.Single(reloaded.Sheets);
            Assert.True(sheet.TryGetCellText(2, 1, out string? value));
            Assert.Equal("Edited", value);
            Assert.Equal("A1:D5", sheet.WorksheetPart.Worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value);
            Assert.Equal("A1:C1", Assert.Single(sheet.GetMergedRanges()).A1Range);
            Assert.Equal(30D, sheet.GetRowDefinitions().Single(row => row.Index == 1).Height);
            Assert.True(sheet.GetColumnDefinitions().Single(column => column.StartIndex == 4).Hidden);
            Assert.Equal(PaneStateValues.FrozenSplit, sheet.WorksheetPart.Worksheet
                .GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>()?.State?.Value);
            AssertPackageEntriesEqualExcept(original, rewritten, "xl/worksheets/sheet1.bin");
            AssertWorksheetRecordsEqualExceptCells(original, rewritten, "xl/worksheets/sheet1.bin", (2, 1));
        }

        [Fact]
        public void Xlsb_GeometryFixture_ConvertsToXlsxWithWorksheetMetadataIntact() {
            using ExcelDocument source = ExcelDocument.Load(GetGeometryExcelGeneratedXlsbFixturePath());
            using var destination = new MemoryStream();

            source.Save(destination, ExcelFileFormat.Xlsx);

            using ExcelDocument converted = ExcelDocument.Load(new MemoryStream(destination.ToArray(), writable: false));
            ExcelSheet sheet = Assert.Single(converted.Sheets);
            Assert.Equal("A1:D5", sheet.WorksheetPart.Worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value);
            Assert.Equal(9D, sheet.DefaultColumnWidth);
            Assert.Equal(18D, sheet.DefaultRowHeight);
            Assert.Equal("A1:C1", Assert.Single(sheet.GetMergedRanges()).A1Range);
            Assert.Equal(30D, sheet.GetRowDefinitions().Single(row => row.Index == 1).Height);
            Assert.True(sheet.GetRowDefinitions().Single(row => row.Index == 3).Hidden);
            Assert.True(sheet.GetColumnDefinitions().Single(column => column.StartIndex == 4).Hidden);
            Pane pane = Assert.IsType<Pane>(sheet.WorksheetPart.Worksheet
                .GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>());
            Assert.Equal("B2", pane.TopLeftCell?.Value);
        }

        [Fact]
        public void Xlsb_HyperlinkFixture_ProjectsExternalAndInternalLinks() {
            using ExcelDocument document = ExcelDocument.Load(GetHyperlinkExcelGeneratedXlsbFixturePath());

            Assert.Equal(2, document.Sheets.Count);
            ExcelSheet links = document.Sheets[0];
            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> hyperlinks = links.GetHyperlinks();
            Assert.Equal(2, hyperlinks.Count);
            Assert.True(hyperlinks["A1"].IsExternal);
            Assert.Equal("https://example.org/officeimo?source=xlsb", hyperlinks["A1"].Target);
            Assert.Equal("External screen tip", hyperlinks["A1"].Tooltip);
            Assert.False(hyperlinks["A2"].IsExternal);
            Assert.Equal("'Target Sheet'!B2", hyperlinks["A2"].Target);
            Assert.Equal("Internal screen tip", hyperlinks["A2"].Tooltip);
            Assert.True(links.TryGetCellText(1, 1, out string? externalDisplay));
            Assert.True(links.TryGetCellText(2, 1, out string? internalDisplay));
            Assert.Equal("External link", externalDisplay);
            Assert.Equal("Internal link", internalDisplay);
        }

        [Fact]
        public void Xlsb_HyperlinkFixture_NativeRewritePreservesLinkRecordsAndRelationships() {
            byte[] original = File.ReadAllBytes(GetHyperlinkExcelGeneratedXlsbFixturePath());
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(original, writable: false));
            document.Sheets[0].CellValue(3, 1, "Edited plain value");

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> hyperlinks = reloaded.Sheets[0].GetHyperlinks();
            Assert.Equal("https://example.org/officeimo?source=xlsb", hyperlinks["A1"].Target);
            Assert.Equal("'Target Sheet'!B2", hyperlinks["A2"].Target);
            Assert.True(reloaded.Sheets[0].TryGetCellText(3, 1, out string? value));
            Assert.Equal("Edited plain value", value);
            AssertPackageEntriesEqualExcept(original, rewritten, "xl/worksheets/sheet1.bin");
            AssertWorksheetRecordsEqualExceptCells(original, rewritten, "xl/worksheets/sheet1.bin", (3, 1));
        }

        [Fact]
        public void Xlsb_HyperlinkFixture_ConvertsToXlsxWithLinksIntact() {
            using ExcelDocument source = ExcelDocument.Load(GetHyperlinkExcelGeneratedXlsbFixturePath());
            using var destination = new MemoryStream();

            source.Save(destination, ExcelFileFormat.Xlsx);

            using ExcelDocument converted = ExcelDocument.Load(new MemoryStream(destination.ToArray(), writable: false));
            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> hyperlinks = converted.Sheets[0].GetHyperlinks();
            Assert.Equal(2, hyperlinks.Count);
            Assert.Equal("https://example.org/officeimo?source=xlsb", hyperlinks["A1"].Target);
            Assert.Equal("External screen tip", hyperlinks["A1"].Tooltip);
            Assert.Equal("'Target Sheet'!B2", hyperlinks["A2"].Target);
            Assert.Equal("Internal screen tip", hyperlinks["A2"].Tooltip);
        }

        [Fact]
        public void Xlsb_HyperlinkMutation_RejectsBeforeNativeWrite() {
            using ExcelDocument document = ExcelDocument.Load(GetHyperlinkExcelGeneratedXlsbFixturePath());
            document.Sheets[0].SetInternalLink(
                2,
                1,
                "'Target Sheet'!C3",
                display: "Internal link",
                style: false,
                tooltip: "Internal screen tip");
            using var destination = new MemoryStream();

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(destination, ExcelFileFormat.Xlsb));

            Assert.Contains("cannot modify hyperlinks", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, destination.Length);
        }

        [Fact]
        public void Xlsb_HyperlinkReference_RejectsMissingRelationship() {
            byte[] package = File.ReadAllBytes(GetHyperlinkExcelGeneratedXlsbFixturePath());
            byte[] worksheet = ReadZipEntry(package, "xl/worksheets/sheet1.bin");
            using var input = new MemoryStream(worksheet, writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
            XlsbRecord hyperlink = records.First(record => record.Type == 494 && ReadXlsbTestUInt32(record.Data, 16) > 0U);
            byte[] tampered = (byte[])hyperlink.Data.Clone();
            tampered[20] = (byte)'x';
            tampered[21] = 0;
            byte[] malformed = ReplaceWorksheetRecords(package, records, hyperlink, tampered);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.Load(new MemoryStream(malformed, writable: false)));

            Assert.Contains("missing or invalid hyperlink relationship", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_HyperlinkImport_EnforcesConfiguredLimit() {
            var options = new ExcelLoadOptions {
                XlsbImportOptions = new XlsbImportOptions { MaxHyperlinks = 1 }
            };

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.Load(GetHyperlinkExcelGeneratedXlsbFixturePath(), options));

            Assert.Contains("limit of 1 worksheet hyperlinks", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_RichFormulaFixture_ProjectsCommonBusinessFormulaTokens() {
            using ExcelDocument document = ExcelDocument.Load(GetRichFormulaExcelGeneratedXlsbFixturePath());

            ExcelSheet sheet = Assert.Single(document.Sheets);
            string[] expected = {
                "IF(A1>5,\"High\",\"Low\")",
                "IFERROR(A1/A2,\"Divide error\")",
                "\"Hello\"&\" \"&\"XLSB\"",
                "ROUND(PI(),2)",
                "SUM(A1:A2)",
                "CHOOSE(2,\"First\",\"Second\",\"Third\")",
                "NA()",
                "(A1+2)*10%"
            };
            for (int row = 1; row <= expected.Length; row++) {
                Assert.Equal(expected[row - 1], sheet.CellAt(row, 2).GetValue().Formula);
            }
            Assert.True(sheet.TryGetCellText(1, 2, out string? ifValue));
            Assert.True(sheet.TryGetCellText(2, 2, out string? ifErrorValue));
            Assert.True(sheet.TryGetCellText(3, 2, out string? concatenatedValue));
            Assert.True(sheet.TryGetCellText(6, 2, out string? chooseValue));
            Assert.Equal("High", ifValue);
            Assert.Equal("Divide error", ifErrorValue);
            Assert.Equal("Hello XLSB", concatenatedValue);
            Assert.Equal("Second", chooseValue);
            Assert.DoesNotContain(document.XlsbImportDiagnostics, diagnostic => diagnostic.Code == "XLSB-FORMULA-PRESERVED");
        }

        [Fact]
        public void Xlsb_RichFormulaFixture_NativeRewritePreservesEveryFormulaPayload() {
            byte[] original = File.ReadAllBytes(GetRichFormulaExcelGeneratedXlsbFixturePath());
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(original, writable: false));
            document.Sheets[0].CellValue(1, 1, 12D);

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.Equal("IF(A1>5,\"High\",\"Low\")", reloaded.Sheets[0].CellAt(1, 2).GetValue().Formula);
            Assert.Equal("IFERROR(A1/A2,\"Divide error\")", reloaded.Sheets[0].CellAt(2, 2).GetValue().Formula);
            Assert.Equal("CHOOSE(2,\"First\",\"Second\",\"Third\")", reloaded.Sheets[0].CellAt(6, 2).GetValue().Formula);
            for (int row = 1; row <= 8; row++) {
                AssertFormulaPayloadEqual(original, rewritten, "xl/worksheets/sheet1.bin", (row, 2));
            }
            AssertPackageEntriesEqualExcept(original, rewritten, "xl/worksheets/sheet1.bin");
        }

        [Fact]
        public void Xlsb_RichFormulaFixture_ConvertsToXlsxWithFormulaTextAndCachedValues() {
            using ExcelDocument source = ExcelDocument.Load(GetRichFormulaExcelGeneratedXlsbFixturePath());
            using var destination = new MemoryStream();

            source.Save(destination, ExcelFileFormat.Xlsx);

            using ExcelDocument converted = ExcelDocument.Load(new MemoryStream(destination.ToArray(), writable: false));
            ExcelSheet sheet = Assert.Single(converted.Sheets);
            Assert.Equal("IF(A1>5,\"High\",\"Low\")", sheet.CellAt(1, 2).GetValue().Formula);
            Assert.Equal("IFERROR(A1/A2,\"Divide error\")", sheet.CellAt(2, 2).GetValue().Formula);
            Assert.Equal("\"Hello\"&\" \"&\"XLSB\"", sheet.CellAt(3, 2).GetValue().Formula);
            Assert.Equal("CHOOSE(2,\"First\",\"Second\",\"Third\")", sheet.CellAt(6, 2).GetValue().Formula);
            Assert.True(sheet.TryGetCellText(2, 2, out string? cachedValue));
            Assert.Equal("Divide error", cachedValue);
        }

        [Fact]
        public void Xlsb_RowSpan_RejectsCellOutsideDeclaredSegmentBounds() {
            byte[] package = File.ReadAllBytes(GetGeometryExcelGeneratedXlsbFixturePath());
            byte[] worksheet = ReadZipEntry(package, "xl/worksheets/sheet1.bin");
            using var input = new MemoryStream(worksheet, writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
            XlsbRecord row = records.First(record => record.Type == 0 && ReadXlsbTestUInt32(record.Data, 0) == 1U);
            byte[] tamperedRow = (byte[])row.Data.Clone();
            WriteXlsbTestUInt32(tamperedRow, 13, 1U);
            WriteXlsbTestUInt32(tamperedRow, 17, 1U);
            WriteXlsbTestUInt32(tamperedRow, 21, 3U);
            byte[] malformed = ReplaceWorksheetRecords(package, records, row, tamperedRow);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.Load(new MemoryStream(malformed, writable: false)));

            Assert.Contains("not covered", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_MergeCollection_EnforcesConfiguredLimitBeforeExpansion() {
            byte[] package = File.ReadAllBytes(GetGeometryExcelGeneratedXlsbFixturePath());
            byte[] worksheet = ReadZipEntry(package, "xl/worksheets/sheet1.bin");
            using var input = new MemoryStream(worksheet, writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
            XlsbRecord beginMerges = Assert.Single(records, record => record.Type == 177);
            byte[] tamperedBegin = (byte[])beginMerges.Data.Clone();
            WriteXlsbTestUInt32(tamperedBegin, 0, 2U);
            byte[] malformed = ReplaceWorksheetRecords(package, records, beginMerges, tamperedBegin);
            var options = new ExcelLoadOptions {
                XlsbImportOptions = new XlsbImportOptions { MaxMergedRanges = 1 }
            };

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.Load(new MemoryStream(malformed, writable: false), options));

            Assert.Contains("configured limit of 1", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_CellStyleReference_RejectsMissingCellFormatBeforeProjection() {
            byte[] package = File.ReadAllBytes(GetStyledExcelGeneratedXlsbFixturePath());
            byte[] worksheet = ReadZipEntry(package, "xl/worksheets/sheet1.bin");
            using var input = new MemoryStream(worksheet, writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
            XlsbRecord firstCell = records.First(record => record.Type >= 1 && record.Type <= 11);
            byte[] tampered = (byte[])firstCell.Data.Clone();
            tampered[4] = 0xFF;
            tampered[5] = 0x00;
            tampered[6] = 0x00;
            tampered[7] = 0x00;

            using var output = new MemoryStream();
            foreach (XlsbRecord record in records) {
                XlsbRecordWriter.Write(output, record.Type, ReferenceEquals(record, firstCell) ? tampered : record.Data);
            }
            byte[] malformed = ReplaceZipEntry(package, "xl/worksheets/sheet1.bin", output.ToArray());

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.Load(new MemoryStream(malformed, writable: false)));

            Assert.Contains("missing cell format 255", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_UnmodifiedSource_CopiesByteForByteToNativeTarget() {
            byte[] source = File.ReadAllBytes(GetExcelGeneratedXlsbFixturePath());
            using var input = new MemoryStream(source, writable: false);
            using ExcelDocument document = ExcelDocument.Load(input);

            byte[] saved = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal(source, saved);
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(saved, "copy.xlsb"));
        }

        [Fact]
        public void Xlsb_ModifiedSource_RewritesCellsAndPreservesOtherPackageParts() {
            byte[] original = File.ReadAllBytes(GetExcelGeneratedXlsbFixturePath());
            using ExcelDocument document = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            document.Sheets[0].CellValue(2, 2, 43);

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(rewritten, "rewritten.xlsb"));
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.True(reloaded.Sheets[0].TryGetCellText(2, 2, out string? value));
            Assert.Equal("43", value);
            AssertPackageEntriesEqualExcept(original, rewritten, "xl/worksheets/sheet1.bin");
            AssertWorksheetCellRecordsEqualExcept(original, rewritten, "xl/worksheets/sheet1.bin", (2, 2));
        }

        [Fact]
        public void Xlsb_NativeRewrite_PreservesCompleteFormulaPayloadWhenCachedResultChanges() {
            byte[] original = File.ReadAllBytes(GetExcelGeneratedXlsbFixturePath());
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(original, writable: false));
            document.Sheets[0].CellValue(3, 2, 51D);
            Cell formulaCell = Assert.Single(document.Sheets[0].WorksheetPart.Worksheet.Descendants<Cell>(),
                cell => string.Equals(cell.CellReference?.Value, "B3", StringComparison.Ordinal));
            Assert.NotNull(formulaCell.CellFormula);

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.True(reloaded.Sheets[0].TryGetCellText(3, 2, out string? cachedValue));
            Assert.Equal("51", cachedValue);
            Cell reloadedFormulaCell = Assert.Single(reloaded.Sheets[0].WorksheetPart.Worksheet.Descendants<Cell>(),
                cell => string.Equals(cell.CellReference?.Value, "B3", StringComparison.Ordinal));
            Assert.Equal("SUM(B2,8)", reloadedFormulaCell.CellFormula?.Text);
            AssertFormulaPayloadEqual(original, rewritten, "xl/worksheets/sheet1.bin", (3, 2));
        }

        [Fact]
        public void Xlsb_NativeRewrite_EncodesChangedCommonFormula() {
            using ExcelDocument document = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            document.Sheets[0].CellFormula(3, 2, "SUM(B2,9)");

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.Equal("SUM(B2,9)", reloaded.Sheets[0].CellAt(3, 2).GetValue().Formula);
        }

        [Fact]
        public void Xlsb_NativeRewrite_HandlesTextBooleanAndNewRowsAcrossSequentialSaves() {
            using ExcelDocument document = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            ExcelSheet sheet = document.Sheets[0];
            sheet.CellValue(2, 1, "Beta");
            sheet.CellValue(1, 3, true);
            sheet.CellValue(4, 1, "New row");

            byte[] first = document.ToBytes(ExcelFileFormat.Xlsb);
            sheet.CellValue(2, 2, 44);
            byte[] second = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(second, writable: false));
            ExcelSheet reloadedSheet = Assert.Single(reloaded.Sheets);
            Assert.True(reloadedSheet.TryGetCellText(2, 1, out string? a2));
            Assert.True(reloadedSheet.TryGetCellText(1, 3, out string? c1));
            Assert.True(reloadedSheet.TryGetCellText(4, 1, out string? a4));
            Assert.True(reloadedSheet.TryGetCellText(2, 2, out string? b2));
            Assert.Equal("Beta", a2);
            Assert.Equal("1", c1);
            Assert.Equal("New row", a4);
            Assert.Equal("44", b2);
            Assert.NotEqual(first, second);
        }

        [Fact]
        public void Xlsb_NativeRewrite_ExpandsDimensionsAndRebuildsSegmentedRowSpans() {
            using ExcelDocument document = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            ExcelSheet sheet = document.Sheets[0];
            sheet.CellValue(2, 1025, "second segment");
            sheet.CellValue(4, 1025, "new row");

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal((0, 3, 0, 1024), ReadWorksheetDimension(rewritten, "xl/worksheets/sheet1.bin"));
            Assert.Equal(new[] { (0, 1), (1024, 1024) }, ReadRowSpans(rewritten, "xl/worksheets/sheet1.bin", 2));
            Assert.Equal(new[] { (1024, 1024) }, ReadRowSpans(rewritten, "xl/worksheets/sheet1.bin", 4));
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.True(reloaded.Sheets[0].TryGetCellText(2, 1025, out string? existingRowValue));
            Assert.True(reloaded.Sheets[0].TryGetCellText(4, 1025, out string? newRowValue));
            Assert.Equal("second segment", existingRowValue);
            Assert.Equal("new row", newRowValue);
        }

        [Fact]
        public void Xlsb_NativeRewrite_UsesExactFirstCellDimensionForEmptySourceSheet() {
            byte[] emptySource = RemoveWorksheetRowsAndCells(File.ReadAllBytes(GetExcelGeneratedXlsbFixturePath()));
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(emptySource, writable: false));
            document.Sheets[0].CellValue(3, 3, "Only cell");

            byte[] rewritten = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal((2, 2, 2, 2), ReadWorksheetDimension(rewritten, "xl/worksheets/sheet1.bin"));
            Assert.Equal(new[] { (2, 2) }, ReadRowSpans(rewritten, "xl/worksheets/sheet1.bin", 3));
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.True(reloaded.Sheets[0].TryGetCellText(3, 3, out string? value));
            Assert.Equal("Only cell", value);
        }

        [Fact]
        public void Xlsb_UnsupportedStructuralMutation_RejectsBeforeWriting() {
            using ExcelDocument document = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            document.Sheets[0].MergeRange("A1:B1");
            using var destination = new MemoryStream();

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(destination, ExcelFileFormat.Xlsb));

            Assert.Contains("merged ranges", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, destination.Length);
        }

        [Fact]
        public void Xlsb_ProjectedWorkbook_SavesAsValidEditableXlsx() {
            using ExcelDocument source = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            source.Sheets[0].CellValue(2, 2, 43);
            using var destination = new MemoryStream();

            source.Save(destination, ExcelFileFormat.Xlsx);
            byte[] xlsx = destination.ToArray();

            Assert.Equal(ExcelFileFormat.Xlsx, ExcelDocumentLoadRouting.DetectFormat(xlsx, "converted.xlsx"));
            using ExcelDocument converted = ExcelDocument.Load(new MemoryStream(xlsx, writable: false));
            Assert.True(converted.Sheets[0].TryGetCellText(2, 2, out string? value));
            Assert.Equal("43", value);
            Cell formulaCell = Assert.Single(
                converted.Sheets[0].DeferredMetadataWorksheetPart.Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "B3");
            Assert.Equal("SUM(B2,8)", formulaCell.CellFormula?.Text);
        }

        [Fact]
        public void Xlsb_ImportLimits_BlockCellExpansion() {
            var options = new ExcelLoadOptions {
                XlsbImportOptions = new XlsbImportOptions { MaxCells = 4 }
            };

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath(), options));

            Assert.Contains("limit of 4 populated cells", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Xlsb_NewWorkbook_SyncAndAsyncSaveProduceReadableNativePackages() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data").CellValue(1, 1, "XLSB");
            using var synchronousDestination = new MemoryStream();
            using var asynchronousDestination = new MemoryStream();

            document.Save(synchronousDestination, ExcelFileFormat.Xlsb);
            await document.SaveAsync(asynchronousDestination, ExcelFileFormat.Xlsb);

            byte[] synchronous = synchronousDestination.ToArray();
            byte[] asynchronous = asynchronousDestination.ToArray();
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(synchronous, "sync.xlsb"));
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(asynchronous, "async.xlsb"));
            using ExcelDocument syncReloaded = ExcelDocument.Load(new MemoryStream(synchronous, writable: false));
            using ExcelDocument asyncReloaded = ExcelDocument.Load(new MemoryStream(asynchronous, writable: false));
            Assert.True(syncReloaded.Sheets[0].TryGetCellText(1, 1, out string? syncValue));
            Assert.True(asyncReloaded.Sheets[0].TryGetCellText(1, 1, out string? asyncValue));
            Assert.Equal("XLSB", syncValue);
            Assert.Equal("XLSB", asyncValue);
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesMultipleSheetsValuesVisibilityAndDimensions() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet values = document.AddWorksheet("Values");
            values.CellValue(2, 2, "Text");
            values.CellValue(3, 3, 42.5D);
            values.CellValue(4, 4, true);
            ExcelSheet hidden = document.AddWorksheet("Hidden Data");
            hidden.CellValue(1, 1, "Hidden");
            hidden.SetHidden(true);

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal((1, 3, 1, 3), ReadWorksheetDimension(package, "xl/worksheets/sheet1.bin"));
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Equal(2, reloaded.Sheets.Count);
            Assert.Equal("Values", reloaded.Sheets[0].Name);
            Assert.Equal("Hidden Data", reloaded.Sheets[1].Name);
            Assert.True(reloaded.Sheets[1].Hidden);
            Assert.True(reloaded.Sheets[0].TryGetCellText(2, 2, out string? text));
            Assert.True(reloaded.Sheets[0].TryGetCellText(3, 3, out string? number));
            Assert.True(reloaded.Sheets[0].TryGetCellText(4, 4, out string? boolean));
            Assert.Equal("Text", text);
            Assert.Equal("42.5", number);
            Assert.Equal("1", boolean);
        }

        [Fact]
        public async Task Xlsb_NewWorkbook_StreamsToNonSeekableDestinations() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Streamed").CellValue(1, 1, "BIFF12");
            using var synchronousDestination = new NonSeekableReadWriteBuffer(Array.Empty<byte>());
            using var asynchronousDestination = new NonSeekableReadWriteBuffer(Array.Empty<byte>());

            document.Save(synchronousDestination, ExcelFileFormat.Xlsb);
            await document.SaveAsync(asynchronousDestination, ExcelFileFormat.Xlsb);

            using ExcelDocument syncReloaded = ExcelDocument.Load(new MemoryStream(synchronousDestination.ToArray(), writable: false));
            using ExcelDocument asyncReloaded = ExcelDocument.Load(new MemoryStream(asynchronousDestination.ToArray(), writable: false));
            Assert.True(syncReloaded.Sheets[0].TryGetCellText(1, 1, out string? syncValue));
            Assert.True(asyncReloaded.Sheets[0].TryGetCellText(1, 1, out string? asyncValue));
            Assert.Equal("BIFF12", syncValue);
            Assert.Equal("BIFF12", asyncValue);
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesCommonFormulaTokensAndCachedResults() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Formula");
            sheet.CellValue(1, 1, 2D);
            sheet.CellValue(2, 1, 3D);
            sheet.CellFormula(3, 1, "SUM(A1:A2)");
            sheet.CellFormula(4, 1, "\"Hello\"&\" \"&\"XLSB\"");
            sheet.CellFormula(5, 1, "IF(A1>A2,\"High\",\"Low\")");
            sheet.CellFormula(6, 1, "IFERROR(A1/0,\"Divide error\")");
            sheet.CellFormula(7, 1, "CHOOSE(2,\"First\",\"Second\",\"Third\")");

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Equal("SUM(A1:A2)", reloaded.Sheets[0].CellAt(3, 1).GetValue().Formula);
            Assert.Equal("\"Hello\"&\" \"&\"XLSB\"", reloaded.Sheets[0].CellAt(4, 1).GetValue().Formula);
            Assert.Equal("IF(A1>A2,\"High\",\"Low\")", reloaded.Sheets[0].CellAt(5, 1).GetValue().Formula);
            Assert.Equal("IFERROR(A1/0,\"Divide error\")", reloaded.Sheets[0].CellAt(6, 1).GetValue().Formula);
            Assert.Equal("CHOOSE(2,\"First\",\"Second\",\"Third\")", reloaded.Sheets[0].CellAt(7, 1).GetValue().Formula);
        }

        [Fact]
        public void Xlsb_NewWorkbook_RejectsUnsupportedFormulaBeforeTouchingDestinationContent() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Formula").CellFormula(1, 1, "Sheet1!A1");
            byte[] sentinel = Enumerable.Range(0, 64).Select(index => (byte)index).ToArray();
            using var destination = new MemoryStream();
            destination.Write(sentinel, 0, sentinel.Length);

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(destination, ExcelFileFormat.Xlsb));

            Assert.Contains("cannot encode", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(sentinel, destination.ToArray());
        }

        [Fact]
        public void Xlsb_NewWorkbook_FileSaveAdoptsNativeStateForSubsequentRewrites() {
            string path = Path.Combine(Path.GetTempPath(), "officeimo-new-native-" + Guid.NewGuid().ToString("N") + ".xlsb");
            try {
                using (ExcelDocument document = ExcelDocument.Create()) {
                    ExcelSheet sheet = document.AddWorksheet("Sequential");
                    sheet.CellValue(1, 1, "First");
                    document.Save(path);
                    Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);

                    sheet.CellValue(1, 1, "Second");
                    document.Save();
                }

                using ExcelDocument reloaded = ExcelDocument.Load(path);
                Assert.True(reloaded.Sheets[0].TryGetCellText(1, 1, out string? value));
                Assert.Equal("Second", value);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesStylesCustomFormatsAndDates() {
            var date = new DateTime(2024, 2, 29);
            using ExcelDocument document = ExcelDocument.Create();
            document.DateSystem = ExcelDateSystem.NineteenFour;
            ExcelSheet sheet = document.AddWorksheet("Styled");
            sheet.CellValue(1, 1, "Heading");
            sheet.CellBold(1, 1);
            sheet.CellBackground(1, 1, "#4472C4");
            sheet.CellWrapText(1, 1);
            sheet.CellValue(2, 1, date);
            sheet.CellValue(3, 1, 12.34567D);
            sheet.FormatCell(3, 1, "0.0000");
            Assert.True(sheet.GetCellStyle(1, 1).Bold);

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Equal(ExcelDateSystem.NineteenFour, reloaded.DateSystem);
            ExcelSheet reloadedSheet = Assert.Single(reloaded.Sheets);
            ExcelCellStyleSnapshot heading = reloadedSheet.GetCellStyle(1, 1);
            Assert.True(heading.Bold);
            Assert.True(heading.WrapText);
            Assert.Equal("solid", heading.FillPatternType);
            Assert.Equal(date, AssertCellValue(reloadedSheet, 2, 1).DateTimeValue);
            Assert.Equal("0.0000", reloadedSheet.GetCellStyle(3, 1).NumberFormatCode);
            Assert.Contains(reloaded.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
                .Elements<NumberingFormat>(), format => format.FormatCode?.Value == "0.0000");
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesWorksheetGeometry() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Geometry");
            sheet.CellValue(1, 1, "Merged heading");
            sheet.CellValue(5, 4, "Extent");
            sheet.SetDefaultColumnWidth(9D);
            sheet.SetDefaultRowHeight(18D);
            sheet.SetColumnWidth(1, 20D);
            sheet.SetColumnHidden(4, true);
            sheet.SetColumnOutline(3, 2, collapsed: true);
            sheet.SetRowHeight(1, 30D);
            sheet.SetRowHidden(3, true);
            sheet.SetRowOutline(4, 2, collapsed: true);
            sheet.Freeze(topRows: 1, leftCols: 1);
            sheet.MergeRange("A1:C1");

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            ExcelSheet result = Assert.Single(reloaded.Sheets);
            Assert.Equal("A1:D5", result.WorksheetPart.Worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value);
            Assert.Equal(9D, result.DefaultColumnWidth);
            Assert.Equal(18D, result.DefaultRowHeight);
            Assert.Equal(30D, result.GetRowDefinitions().Single(row => row.Index == 1).Height);
            Assert.True(result.GetRowDefinitions().Single(row => row.Index == 3).Hidden);
            Assert.Equal((byte)2, result.GetRowDefinitions().Single(row => row.Index == 4).OutlineLevel);
            Assert.True(result.GetRowDefinitions().Single(row => row.Index == 4).Collapsed);
            Assert.Equal(20D, result.GetColumnDefinitions().Single(column => column.StartIndex == 1).Width);
            Assert.True(result.GetColumnDefinitions().Single(column => column.StartIndex == 4).Hidden);
            Assert.Equal((byte)2, result.GetColumnDefinitions().Single(column => column.StartIndex == 3).OutlineLevel);
            Assert.True(result.GetColumnDefinitions().Single(column => column.StartIndex == 3).Collapsed);
            Assert.Equal("A1:C1", Assert.Single(result.GetMergedRanges()).A1Range);
            Pane pane = Assert.IsType<Pane>(result.WorksheetPart.Worksheet
                .GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>());
            Assert.Equal(1D, pane.HorizontalSplit?.Value);
            Assert.Equal(1D, pane.VerticalSplit?.Value);
            Assert.Equal("B2", pane.TopLeftCell?.Value);
            Assert.Equal(PaneStateValues.Frozen, pane.State?.Value);
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesExternalAndInternalHyperlinks() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet links = document.AddWorksheet("Links");
            ExcelSheet target = document.AddWorksheet("Target Sheet");
            links.SetHyperlink(
                1,
                1,
                "https://example.org/officeimo?source=xlsb&mode=native",
                display: "External link",
                style: false,
                tooltip: "External screen tip");
            links.SetInternalLink(
                2,
                1,
                target,
                "B2",
                display: "Internal link",
                style: false,
                tooltip: "Internal screen tip");
            links.SetHyperlink(
                3,
                1,
                "../docs/zażółć-spec.pdf",
                display: "Relative link",
                style: false,
                tooltip: "Relative screen tip");

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using (var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read)) {
                ZipArchiveEntry relationshipsEntry = Assert.IsType<ZipArchiveEntry>(
                    archive.GetEntry("xl/worksheets/_rels/sheet1.bin.rels"));
                using var reader = new StreamReader(relationshipsEntry.Open());
                string relationshipsXml = reader.ReadToEnd();
                Assert.Contains("TargetMode=\"External\"", relationshipsXml, StringComparison.Ordinal);
                Assert.Contains("source=xlsb&amp;mode=native", relationshipsXml, StringComparison.Ordinal);
                Assert.Contains("../docs/zażółć-spec.pdf", relationshipsXml, StringComparison.Ordinal);
                Assert.Null(archive.GetEntry("xl/worksheets/_rels/sheet2.bin.rels"));
            }

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> hyperlinks = reloaded.Sheets[0].GetHyperlinks();
            Assert.Equal(3, hyperlinks.Count);
            Assert.True(hyperlinks["A1"].IsExternal);
            Assert.Equal("https://example.org/officeimo?source=xlsb&mode=native", hyperlinks["A1"].Target);
            Assert.Equal("External screen tip", hyperlinks["A1"].Tooltip);
            Assert.False(hyperlinks["A2"].IsExternal);
            Assert.Equal("'Target Sheet'!B2", hyperlinks["A2"].Target);
            Assert.Equal("Internal screen tip", hyperlinks["A2"].Tooltip);
            Assert.True(hyperlinks["A3"].IsExternal);
            Assert.Equal("../docs/zażółć-spec.pdf", hyperlinks["A3"].Target);
            Assert.Equal("Relative screen tip", hyperlinks["A3"].Tooltip);
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesCalculationProperties() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Calculation").CellValue(1, 1, 1D);
            document.WorkbookRoot.Append(new CalculationProperties {
                CalculationId = 42U,
                CalculationMode = CalculateModeValues.AutoNoTable,
                FullCalculationOnLoad = true,
                ReferenceMode = ReferenceModeValues.R1C1,
                Iterate = true,
                IterateCount = 42U,
                IterateDelta = 0.0005D,
                FullPrecision = false,
                CalculationCompleted = false,
                CalculationOnSave = false,
                ConcurrentCalculation = true,
                ConcurrentManualCount = 4U,
                ForceFullCalculation = true
            });

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using (var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read)) {
                using Stream workbookStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/workbook.bin")).Open();
                XlsbRecord record = Assert.Single(XlsbRecordReader.ReadAll(workbookStream), item => item.Type == 157);
                Assert.Equal(26, record.Data.Length);
                Assert.Equal(42U, BitConverter.ToUInt32(record.Data, 0));
                Assert.Equal(2U, BitConverter.ToUInt32(record.Data, 4));
                Assert.Equal(42U, BitConverter.ToUInt32(record.Data, 8));
                Assert.Equal(0.0005D, BitConverter.ToDouble(record.Data, 12));
                Assert.Equal(4, BitConverter.ToInt32(record.Data, 20));
                Assert.Equal((ushort)0x01D5, BitConverter.ToUInt16(record.Data, 24));
            }

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            CalculationProperties properties = Assert.IsType<CalculationProperties>(
                reloaded.WorkbookRoot.GetFirstChild<CalculationProperties>());
            Assert.Equal(42U, properties.CalculationId?.Value);
            Assert.Equal(CalculateModeValues.AutoNoTable, properties.CalculationMode?.Value);
            Assert.True(properties.FullCalculationOnLoad?.Value);
            Assert.Equal(ReferenceModeValues.R1C1, properties.ReferenceMode?.Value);
            Assert.True(properties.Iterate?.Value);
            Assert.Equal(42U, properties.IterateCount?.Value);
            Assert.Equal(0.0005D, properties.IterateDelta?.Value);
            Assert.False(properties.FullPrecision?.Value);
            Assert.False(properties.CalculationCompleted?.Value);
            Assert.False(properties.CalculationOnSave?.Value);
            Assert.True(properties.ConcurrentCalculation?.Value);
            Assert.Equal(4U, properties.ConcurrentManualCount?.Value);
            Assert.True(properties.ForceFullCalculation?.Value);
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesClassicWorkbookProtection() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Protected").CellValue(1, 1, "Locked structure");
            document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                ProtectStructure = true,
                ProtectWindows = true,
                LegacyPasswordHash = "CAFE"
            });
            WorkbookProtection protection = Assert.IsType<WorkbookProtection>(
                document.WorkbookRoot.GetFirstChild<WorkbookProtection>());
            protection.LockRevision = true;
            protection.RevisionsPassword = "BEEF";

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using (var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read)) {
                using Stream workbookStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/workbook.bin")).Open();
                XlsbRecord record = Assert.Single(XlsbRecordReader.ReadAll(workbookStream), item => item.Type == 534);
                Assert.Equal(new byte[] { 0xFE, 0xCA, 0xEF, 0xBE, 0x07, 0x00 }, record.Data);
            }

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.True(reloaded.IsWorkbookProtected);
            WorkbookProtection result = Assert.IsType<WorkbookProtection>(
                reloaded.WorkbookRoot.GetFirstChild<WorkbookProtection>());
            Assert.Equal("CAFE", result.WorkbookPassword?.Value);
            Assert.Equal("BEEF", result.RevisionsPassword?.Value);
            Assert.True(result.LockStructure?.Value);
            Assert.True(result.LockWindows?.Value);
            Assert.True(result.LockRevision?.Value);

            reloaded.Sheets[0].CellValue(1, 1, "Updated under protection");
            byte[] rewritten = reloaded.ToBytes(ExcelFileFormat.Xlsb);
            using (var archive = new ZipArchive(new MemoryStream(rewritten, writable: false), ZipArchiveMode.Read)) {
                using Stream workbookStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/workbook.bin")).Open();
                XlsbRecord record = Assert.Single(XlsbRecordReader.ReadAll(workbookStream), item => item.Type == 534);
                Assert.Equal(new byte[] { 0xFE, 0xCA, 0xEF, 0xBE, 0x07, 0x00 }, record.Data);
            }
            using ExcelDocument rewrittenDocument = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.True(rewrittenDocument.IsWorkbookProtected);
            Assert.True(rewrittenDocument.Sheets[0].TryGetCellText(1, 1, out string? rewrittenValue));
            Assert.Equal("Updated under protection", rewrittenValue);
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesAndProjectsDefinedNamesAndPrintRanges() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet data = document.AddWorksheet("Data Sheet");
            document.AddWorksheet("Summary");
            document.SetNamedRange("RevenueData", "'Data Sheet'!A1:B3", save: false, hidden: true);
            data.SetNamedRange("LocalCell", "C2", save: false);
            document.SetPrintArea(data, "A1:C10", save: false);
            document.SetPrintTitles(data, firstRow: 1, lastRow: 2, firstCol: 1, lastCol: 1, save: false);
            DefinedName revenueName = Assert.Single(
                document.WorkbookRoot.DefinedNames!.Elements<DefinedName>(),
                name => name.Name?.Value == "RevenueData");
            revenueName.Comment = "Projected from BIFF12";

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using (var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read)) {
                using Stream workbookStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/workbook.bin")).Open();
                IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(workbookStream);
                Assert.Equal(4, records.Count(record => record.Type == 39));
                Assert.Single(records, record => record.Type == 353);
                Assert.Single(records, record => record.Type == 357);
                XlsbRecord externSheet = Assert.Single(records, record => record.Type == 362);
                Assert.Single(records, record => record.Type == 354);
                Assert.True(
                    records.ToList().FindIndex(record => record.Type == 353)
                    < records.ToList().FindIndex(record => record.Type == 39));
                Assert.Equal(4U, BitConverter.ToUInt32(externSheet.Data, 0));
                Assert.Equal(-2, BitConverter.ToInt32(externSheet.Data, 8));
                Assert.Equal(-2, BitConverter.ToInt32(externSheet.Data, 12));
                Assert.Equal(-1, BitConverter.ToInt32(externSheet.Data, 20));
                Assert.Equal(-1, BitConverter.ToInt32(externSheet.Data, 24));
                Assert.Equal(0, BitConverter.ToInt32(externSheet.Data, 32));
                Assert.Equal(0, BitConverter.ToInt32(externSheet.Data, 36));
                Assert.Equal(1, BitConverter.ToInt32(externSheet.Data, 44));
                Assert.Equal(1, BitConverter.ToInt32(externSheet.Data, 48));
                XlsbRecord firstName = Assert.Single(records, record =>
                    record.Type == 39 && BitConverter.ToUInt32(record.Data, 0) == 0x00000001U);
                int formulaOffset = 17 + "RevenueData".Length * 2;
                Assert.Equal((byte)0x3B, firstName.Data[formulaOffset]);
                Assert.Equal((ushort)2, BitConverter.ToUInt16(firstName.Data, formulaOffset + 1));
            }

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            ExcelSheet reloadedData = Assert.Single(reloaded.Sheets, sheet => sheet.Name == "Data Sheet");
            Assert.Equal("'Data Sheet'!$A$1:$B$3", reloaded.GetNamedRange("RevenueData"));
            Assert.Equal("$C$2", reloadedData.GetNamedRange("LocalCell"));
            Assert.Equal("$A$1:$C$10", reloadedData.GetPrintArea());
            ExcelPrintTitles titles = reloadedData.GetPrintTitles();
            Assert.Equal(1, titles.FirstRow);
            Assert.Equal(2, titles.LastRow);
            Assert.Equal(1, titles.FirstColumn);
            Assert.Equal(1, titles.LastColumn);
            DefinedName projectedRevenueName = Assert.Single(
                reloaded.WorkbookRoot.DefinedNames!.Elements<DefinedName>(),
                name => name.Name?.Value == "RevenueData");
            Assert.True(projectedRevenueName.Hidden?.Value);
            Assert.Equal("Projected from BIFF12", projectedRevenueName.Comment?.Value);

            reloadedData.CellValue(1, 1, "Edited with names preserved");
            byte[] rewritten = reloaded.ToBytes(ExcelFileFormat.Xlsb);
            using ExcelDocument rewrittenDocument = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.Equal("'Data Sheet'!$A$1:$B$3", rewrittenDocument.GetNamedRange("RevenueData"));
            Assert.Equal("$A$1:$C$10", rewrittenDocument.Sheets[0].GetPrintArea());
        }

        [Fact]
        public void Xlsb_NewWorkbook_WritesAndProjectsCommonWorksheetMetadata() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "Status");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "Open");
            sheet.CellValue(2, 2, 12.5D);
            sheet.CellValue(3, 1, "Closed");
            sheet.CellValue(3, 2, 8D);
            sheet.AddAutoFilter("A1:B3", new Dictionary<uint, IEnumerable<string>> {
                [0U] = new[] { "Open", "Closed" }
            });
            sheet.Protect(new ExcelSheetProtectionOptions {
                LegacyPasswordHash = "CAFE",
                AllowSelectLockedCells = true,
                AllowSelectUnlockedCells = true,
                AllowSort = true,
                AllowAutoFilter = true,
                ProtectObjects = true,
                ProtectScenarios = false
            });
            sheet.SetPrintOptions(
                printGridLines: true,
                printHeadings: true,
                horizontalCentered: true,
                verticalCentered: false,
                save: false);
            sheet.SetMargins(0.25D, 0.35D, 0.45D, 0.55D, 0.2D, 0.3D);
            sheet.SetPageSetup(
                fitToWidth: 1U,
                fitToHeight: 0U,
                scale: 80U,
                pageOrder: ExcelPageOrder.OverThenDown,
                paperSize: ExcelPaperSize.Letter);
            sheet.SetOrientation(ExcelPageOrientation.Landscape);
            sheet.SetTabColor("336699");
            sheet.SetOutlineSummary(summaryBelow: false, summaryRight: true);
            SheetProperties properties = Assert.IsType<SheetProperties>(
                sheet.WorksheetPart.Worksheet.GetFirstChild<SheetProperties>());
            properties.CodeName = "ReportSheet";
            properties.Published = true;
            PageSetup setup = Assert.IsType<PageSetup>(sheet.WorksheetPart.Worksheet.GetFirstChild<PageSetup>());
            setup.HorizontalDpi = 300U;
            setup.VerticalDpi = 600U;
            setup.Copies = 2U;
            setup.FirstPageNumber = 3U;
            setup.UseFirstPageNumber = true;
            setup.BlackAndWhite = true;
            setup.Draft = true;
            setup.CellComments = CellCommentsValues.AtEnd;
            setup.Errors = PrintErrorValues.Dash;
            sheet.SetHeaderFooter(
                headerLeft: "Quarterly",
                headerCenter: "Report",
                footerRight: "Page &P of &N",
                differentFirstPage: true,
                differentOddEven: true,
                alignWithMargins: true,
                scaleWithDoc: true);
            sheet.SetFirstPageHeaderFooter(headerCenter: "First page", footerCenter: "Confidential");
            sheet.SetEvenPageHeaderFooter(headerCenter: "Even page", footerRight: "Page &P");

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);
            using (var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read)) {
                using Stream workbookStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/workbook.bin")).Open();
                IReadOnlyList<XlsbRecord> workbookRecords = XlsbRecordReader.ReadAll(workbookStream);
                XlsbRecord filterDatabaseName = Assert.Single(workbookRecords, record =>
                    record.Type == 39 && BitConverter.ToUInt32(record.Data, 0) == 0x00000021U);
                Assert.NotEmpty(filterDatabaseName.Data);

                using Stream worksheetStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/worksheets/sheet1.bin")).Open();
                IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(worksheetStream);
                XlsbRecord protection = Assert.Single(records, record => record.Type == 535);
                Assert.Equal(66, protection.Data.Length);
                Assert.Equal((ushort)0xCAFE, BitConverter.ToUInt16(protection.Data, 0));
                Assert.Equal(1U, BitConverter.ToUInt32(protection.Data, 2));
                Assert.Equal(0U, BitConverter.ToUInt32(protection.Data, 6));
                Assert.Equal(1U, BitConverter.ToUInt32(protection.Data, 10));
                Assert.Single(records, record => record.Type == 161);
                Assert.Single(records, record => record.Type == 162);
                Assert.Single(records, record => record.Type == 163);
                Assert.Single(records, record => record.Type == 165);
                Assert.Equal(2, records.Count(record => record.Type == 167));
                Assert.Single(records, record => record.Type == 166);
                Assert.Single(records, record => record.Type == 164);
                XlsbRecord printOptions = Assert.Single(records, record => record.Type == 477);
                Assert.Equal((ushort)0x000D, BitConverter.ToUInt16(printOptions.Data, 0));
                XlsbRecord margins = Assert.Single(records, record => record.Type == 476);
                Assert.Equal(0.25D, BitConverter.ToDouble(margins.Data, 0));
                Assert.Equal(0.3D, BitConverter.ToDouble(margins.Data, 40));
                XlsbRecord worksheetProperties = Assert.Single(records, record => record.Type == 147);
                uint worksheetFlags = (uint)(worksheetProperties.Data[0]
                    | (worksheetProperties.Data[1] << 8)
                    | (worksheetProperties.Data[2] << 16));
                Assert.Equal(0x000008U, worksheetFlags & 0x000008U);
                Assert.Equal(0U, worksheetFlags & 0x000040U);
                Assert.Equal(0x000080U, worksheetFlags & 0x000080U);
                Assert.Equal(0x000100U, worksheetFlags & 0x000100U);
                Assert.Equal(0x000400U, worksheetFlags & 0x000400U);
                Assert.Equal(0x020000U, worksheetFlags & 0x020000U);
                Assert.Equal((byte)0x05, worksheetProperties.Data[3]);
                Assert.Equal(new byte[] { 0x33, 0x66, 0x99, 0xFF }, worksheetProperties.Data.Skip(7).Take(4));
                XlsbRecord pageSetup = Assert.Single(records, record => record.Type == 478);
                Assert.Equal(38, pageSetup.Data.Length);
                Assert.Equal(1U, BitConverter.ToUInt32(pageSetup.Data, 0));
                Assert.Equal(80U, BitConverter.ToUInt32(pageSetup.Data, 4));
                Assert.Equal(300U, BitConverter.ToUInt32(pageSetup.Data, 8));
                Assert.Equal(600U, BitConverter.ToUInt32(pageSetup.Data, 12));
                Assert.Equal(2U, BitConverter.ToUInt32(pageSetup.Data, 16));
                Assert.Equal(3, BitConverter.ToInt32(pageSetup.Data, 20));
                Assert.Equal((ushort)0x05BB, BitConverter.ToUInt16(pageSetup.Data, 32));
                Assert.Equal(uint.MaxValue, BitConverter.ToUInt32(pageSetup.Data, 34));
                XlsbRecord headerFooter = Assert.Single(records, record => record.Type == 479);
                Assert.Equal((ushort)0x000F, BitConverter.ToUInt16(headerFooter.Data, 0));
                Assert.Single(records, record => record.Type == 480);
            }

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            ExcelSheet reloadedSheet = Assert.Single(reloaded.Sheets);
            Assert.True(reloadedSheet.IsProtected);
            SheetProtection reloadedProtection = Assert.IsType<SheetProtection>(
                reloadedSheet.WorksheetPart.Worksheet.GetFirstChild<SheetProtection>());
            Assert.Equal("CAFE", reloadedProtection.Password?.Value);
            Assert.True(reloadedProtection.Objects?.Value);
            Assert.False(reloadedProtection.Scenarios?.Value);
            Assert.False(reloadedProtection.Sort?.Value);
            Assert.False(reloadedProtection.AutoFilter?.Value);

            AutoFilter reloadedFilter = Assert.IsType<AutoFilter>(
                reloadedSheet.WorksheetPart.Worksheet.GetFirstChild<AutoFilter>());
            Assert.Equal("A1:B3", reloadedFilter.Reference?.Value);
            Filters reloadedFilters = Assert.IsType<Filters>(Assert.Single(reloadedFilter.Elements<FilterColumn>()).GetFirstChild<Filters>());
            Assert.Equal(new[] { "Open", "Closed" }, reloadedFilters.Elements<Filter>().Select(filter => filter.Val?.Value));
            Assert.Contains(
                reloaded.WorkbookRoot.DefinedNames!.Elements<DefinedName>(),
                name => name.Name?.Value == "_FilterDatabase"
                    && name.LocalSheetId?.Value == 0U
                    && name.Text == "'Report'!$A$1:$B$3");

            ExcelSheetPrintOptions projectedPrintOptions = reloadedSheet.GetPrintOptions();
            Assert.True(projectedPrintOptions.PrintGridLines);
            Assert.True(projectedPrintOptions.PrintHeadings);
            Assert.True(projectedPrintOptions.HorizontalCentered);
            Assert.False(projectedPrintOptions.VerticalCentered);
            PageMargins projectedMargins = Assert.IsType<PageMargins>(
                reloadedSheet.WorksheetPart.Worksheet.GetFirstChild<PageMargins>());
            Assert.Equal(0.25D, projectedMargins.Left?.Value);
            Assert.Equal(0.3D, projectedMargins.Footer?.Value);
            SheetProperties projectedProperties = Assert.IsType<SheetProperties>(
                reloadedSheet.WorksheetPart.Worksheet.GetFirstChild<SheetProperties>());
            Assert.Equal("ReportSheet", projectedProperties.CodeName?.Value);
            Assert.True(projectedProperties.Published?.Value);
            Assert.Equal("FF336699", projectedProperties.TabColor?.Rgb?.Value);
            Assert.False(projectedProperties.GetFirstChild<OutlineProperties>()?.SummaryBelow?.Value);
            Assert.True(projectedProperties.GetFirstChild<OutlineProperties>()?.SummaryRight?.Value);
            Assert.True(projectedProperties.GetFirstChild<PageSetupProperties>()?.FitToPage?.Value);
            PageSetup projectedSetup = Assert.IsType<PageSetup>(
                reloadedSheet.WorksheetPart.Worksheet.GetFirstChild<PageSetup>());
            Assert.Equal(ExcelPaperSize.Letter, (ExcelPaperSize)projectedSetup.PaperSize!.Value);
            Assert.Equal(80U, projectedSetup.Scale?.Value);
            Assert.Equal(OrientationValues.Landscape, projectedSetup.Orientation?.Value);
            Assert.Equal(PageOrderValues.OverThenDown, projectedSetup.PageOrder?.Value);
            Assert.Equal(CellCommentsValues.AtEnd, projectedSetup.CellComments?.Value);
            Assert.Equal(PrintErrorValues.Dash, projectedSetup.Errors?.Value);
            Assert.Equal(3U, projectedSetup.FirstPageNumber?.Value);
            HeaderFooter projectedHeaderFooter = Assert.IsType<HeaderFooter>(
                reloadedSheet.WorksheetPart.Worksheet.GetFirstChild<HeaderFooter>());
            Assert.Equal("&LQuarterly&CReport", projectedHeaderFooter.OddHeader?.Text);
            Assert.Equal("&CFirst page", projectedHeaderFooter.FirstHeader?.Text);
            Assert.Equal("&CEven page", projectedHeaderFooter.EvenHeader?.Text);
            Assert.Equal("&RPage &P of &N", projectedHeaderFooter.OddFooter?.Text);

            reloadedSheet.CellValue(2, 2, 13.5D);
            byte[] rewritten = reloaded.ToBytes(ExcelFileFormat.Xlsb);
            using ExcelDocument rewrittenDocument = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            Assert.True(rewrittenDocument.Sheets[0].IsProtected);
            Assert.Equal("A1:B3", rewrittenDocument.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<AutoFilter>()?.Reference?.Value);
            Assert.Equal(0.25D, rewrittenDocument.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<PageMargins>()?.Left?.Value);
            Assert.Equal("ReportSheet", rewrittenDocument.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<SheetProperties>()?.CodeName?.Value);
            Assert.Equal(80U, rewrittenDocument.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<PageSetup>()?.Scale?.Value);
            Assert.Equal("&CFirst page", rewrittenDocument.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<HeaderFooter>()?.FirstHeader?.Text);
        }

        private static byte[] CreateMinimalXlsbPackage() {
            byte[] workbookRecords = {
                0x83, 0x01, 0x00,
                0x84, 0x01, 0x00
            };

            using var package = new MemoryStream();
            using (var archive = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteZipEntry(
                    archive,
                    "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
                    "</Types>");
                WriteZipEntry(
                    archive,
                    "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
                    "</Relationships>");
                WriteZipEntry(archive, "xl/workbook.bin", workbookRecords);
            }

            return package.ToArray();
        }

        private static string GetExcelGeneratedXlsbFixturePath() {
            return Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "basic-values-formula.xlsb");
        }

        private static string GetStyledExcelGeneratedXlsbFixturePath() {
            return Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "styles-dates-formulas.xlsb");
        }

        private static string GetGeometryExcelGeneratedXlsbFixturePath() {
            return Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "worksheet-geometry.xlsb");
        }

        private static string GetHyperlinkExcelGeneratedXlsbFixturePath() {
            return Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "hyperlinks.xlsb");
        }

        private static string GetRichFormulaExcelGeneratedXlsbFixturePath() {
            return Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "formulas-rich.xlsb");
        }

        private static void AssertColumn(
            Column column,
            uint first,
            uint last,
            double width,
            bool hidden,
            byte outlineLevel,
            bool collapsed) {
            Assert.Equal(first, column.Min?.Value);
            Assert.Equal(last, column.Max?.Value);
            Assert.Equal(width, column.Width?.Value);
            Assert.Equal(hidden, column.Hidden?.Value == true);
            Assert.Equal(outlineLevel, column.OutlineLevel?.Value ?? 0);
            Assert.Equal(collapsed, column.Collapsed?.Value == true);
        }

        private static ExcelCellValueSnapshot AssertCellValue(ExcelSheet sheet, int row, int column) {
            Assert.True(sheet.TryGetCellValueSnapshot(row, column, out ExcelCellValueSnapshot? value));
            return Assert.IsType<ExcelCellValueSnapshot>(value);
        }

        private static void AssertPackageEntriesEqualExcept(
            byte[] expectedPackage,
            byte[] actualPackage,
            params string[] excludedEntries) {
            var excluded = new HashSet<string>(excludedEntries, StringComparer.OrdinalIgnoreCase);
            using var expectedStream = new MemoryStream(expectedPackage, writable: false);
            using var actualStream = new MemoryStream(actualPackage, writable: false);
            using var expectedArchive = new ZipArchive(expectedStream, ZipArchiveMode.Read, leaveOpen: false);
            using var actualArchive = new ZipArchive(actualStream, ZipArchiveMode.Read, leaveOpen: false);
            string[] expectedNames = expectedArchive.Entries.Select(entry => entry.FullName).OrderBy(name => name, StringComparer.Ordinal).ToArray();
            string[] actualNames = actualArchive.Entries.Select(entry => entry.FullName).OrderBy(name => name, StringComparer.Ordinal).ToArray();
            Assert.Equal(expectedNames, actualNames);

            foreach (ZipArchiveEntry expectedEntry in expectedArchive.Entries.Where(entry => !excluded.Contains(entry.FullName))) {
                ZipArchiveEntry actualEntry = Assert.Single(
                    actualArchive.Entries,
                    entry => string.Equals(entry.FullName, expectedEntry.FullName, StringComparison.OrdinalIgnoreCase));
                using Stream expected = expectedEntry.Open();
                using Stream actual = actualEntry.Open();
                using var expectedBytes = new MemoryStream();
                using var actualBytes = new MemoryStream();
                expected.CopyTo(expectedBytes);
                actual.CopyTo(actualBytes);
                Assert.Equal(expectedBytes.ToArray(), actualBytes.ToArray());
            }
        }

        private static void AssertWorksheetCellRecordsEqualExcept(
            byte[] expectedPackage,
            byte[] actualPackage,
            string partName,
            params (int Row, int Column)[] excludedCells) {
            var excluded = new HashSet<(int Row, int Column)>(excludedCells);
            IReadOnlyDictionary<(int Row, int Column), XlsbRecord> expected = ReadWorksheetCellRecords(expectedPackage, partName);
            IReadOnlyDictionary<(int Row, int Column), XlsbRecord> actual = ReadWorksheetCellRecords(actualPackage, partName);
            Assert.Equal(expected.Keys.OrderBy(key => key), actual.Keys.OrderBy(key => key));

            foreach (KeyValuePair<(int Row, int Column), XlsbRecord> pair in expected) {
                if (excluded.Contains(pair.Key)) continue;
                Assert.Equal(pair.Value.Type, actual[pair.Key].Type);
                Assert.Equal(pair.Value.Data, actual[pair.Key].Data);
            }
        }

        private static void AssertWorksheetRecordsEqualExceptCells(
            byte[] expectedPackage,
            byte[] actualPackage,
            string partName,
            params (int Row, int Column)[] excludedCells) {
            var excluded = new HashSet<(int Row, int Column)>(excludedCells);
            IReadOnlyList<(XlsbRecord Record, (int Row, int Column)? Cell)> expected = ReadWorksheetRecords(expectedPackage, partName);
            IReadOnlyList<(XlsbRecord Record, (int Row, int Column)? Cell)> actual = ReadWorksheetRecords(actualPackage, partName);
            Assert.Equal(expected.Count, actual.Count);
            for (int index = 0; index < expected.Count; index++) {
                Assert.Equal(expected[index].Cell, actual[index].Cell);
                if (expected[index].Cell.HasValue && excluded.Contains(expected[index].Cell!.Value)) continue;
                Assert.Equal(expected[index].Record.Type, actual[index].Record.Type);
                Assert.True(
                    expected[index].Record.Data.SequenceEqual(actual[index].Record.Data),
                    $"Worksheet record {index} (type {expected[index].Record.Type}, cell {expected[index].Cell}) changed unexpectedly. " +
                    $"Expected {BitConverter.ToString(expected[index].Record.Data).Replace("-", string.Empty)}, " +
                    $"actual {BitConverter.ToString(actual[index].Record.Data).Replace("-", string.Empty)}.");
            }
        }

        private static IReadOnlyList<(XlsbRecord Record, (int Row, int Column)? Cell)> ReadWorksheetRecords(
            byte[] package,
            string partName) {
            using var part = new MemoryStream(ReadZipEntry(package, partName), writable: false);
            var result = new List<(XlsbRecord Record, (int Row, int Column)? Cell)>();
            int row = -1;
            foreach (XlsbRecord record in XlsbRecordReader.ReadAll(part)) {
                if (record.Type == 0) {
                    row = checked((int)ReadXlsbTestUInt32(record.Data, 0) + 1);
                    result.Add((record, null));
                } else if ((record.Type >= 1 && record.Type <= 11) || record.Type == 62) {
                    int column = checked((int)ReadXlsbTestUInt32(record.Data, 0) + 1);
                    result.Add((record, (row, column)));
                } else {
                    result.Add((record, null));
                }
            }
            return result.AsReadOnly();
        }

        private static byte[] ReplaceWorksheetRecords(
            byte[] package,
            IReadOnlyList<XlsbRecord> records,
            XlsbRecord target,
            byte[] replacementData) {
            using var output = new MemoryStream();
            foreach (XlsbRecord record in records) {
                XlsbRecordWriter.Write(output, record.Type, ReferenceEquals(record, target) ? replacementData : record.Data);
            }
            return ReplaceZipEntry(package, "xl/worksheets/sheet1.bin", output.ToArray());
        }

        private static byte[] RemoveWorksheetRowsAndCells(byte[] package) {
            byte[] worksheet = ReadZipEntry(package, "xl/worksheets/sheet1.bin");
            using var input = new MemoryStream(worksheet, writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
            using var output = new MemoryStream();
            foreach (XlsbRecord record in records) {
                if (record.Type == 0 || (record.Type >= 1 && record.Type <= 11) || record.Type == 62) continue;
                XlsbRecordWriter.Write(output, record.Type, record.Type == 148 ? new byte[16] : record.Data);
            }
            return ReplaceZipEntry(package, "xl/worksheets/sheet1.bin", output.ToArray());
        }

        private static uint ReadXlsbTestUInt32(byte[] data, int offset) {
            return (uint)(data[offset]
                | (data[offset + 1] << 8)
                | (data[offset + 2] << 16)
                | (data[offset + 3] << 24));
        }

        private static void WriteXlsbTestUInt32(byte[] data, int offset, uint value) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
            data[offset + 2] = (byte)(value >> 16);
            data[offset + 3] = (byte)(value >> 24);
        }

        private static void AssertFormulaPayloadEqual(
            byte[] expectedPackage,
            byte[] actualPackage,
            string partName,
            (int Row, int Column) cell) {
            XlsbRecord expected = ReadWorksheetCellRecords(expectedPackage, partName)[cell];
            XlsbRecord actual = ReadWorksheetCellRecords(actualPackage, partName)[cell];
            Assert.Equal(expected.Type, actual.Type);
            int cachedValueSize = expected.Type == 9 ? 8 : expected.Type == 8 ? ReadWideStringSize(expected.Data, 8) : 1;
            int formulaOffset = 8 + cachedValueSize;
            Assert.Equal(expected.Data.Skip(formulaOffset), actual.Data.Skip(formulaOffset));
        }

        private static int ReadWideStringSize(byte[] data, int offset) {
            uint characters = (uint)(data[offset]
                | (data[offset + 1] << 8)
                | (data[offset + 2] << 16)
                | (data[offset + 3] << 24));
            return checked(4 + (int)characters * 2);
        }

        private static IReadOnlyDictionary<(int Row, int Column), XlsbRecord> ReadWorksheetCellRecords(
            byte[] package,
            string partName) {
            using var packageStream = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
            ZipArchiveEntry entry = Assert.Single(archive.Entries,
                candidate => string.Equals(candidate.FullName, partName, StringComparison.OrdinalIgnoreCase));
            using Stream part = entry.Open();
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(part);
            var cells = new Dictionary<(int Row, int Column), XlsbRecord>();
            int row = -1;
            foreach (XlsbRecord record in records) {
                if (record.Type == 0) {
                    var cursor = new XlsbBinaryCursor(record.Data);
                    row = checked(cursor.ReadInt32() + 1);
                } else if ((record.Type >= 1 && record.Type <= 11) || record.Type == 62) {
                    var cursor = new XlsbBinaryCursor(record.Data);
                    int column = checked(cursor.ReadInt32() + 1);
                    cells.Add((row, column), record);
                }
            }

            return cells;
        }

        private static (int FirstRow, int LastRow, int FirstColumn, int LastColumn) ReadWorksheetDimension(
            byte[] package,
            string partName) {
            using var part = new MemoryStream(ReadZipEntry(package, partName), writable: false);
            XlsbRecord dimension = Assert.Single(XlsbRecordReader.ReadAll(part), record => record.Type == 148);
            var cursor = new XlsbBinaryCursor(dimension.Data);
            return (cursor.ReadInt32(), cursor.ReadInt32(), cursor.ReadInt32(), cursor.ReadInt32());
        }

        private static IReadOnlyList<(int FirstColumn, int LastColumn)> ReadRowSpans(
            byte[] package,
            string partName,
            int row) {
            using var part = new MemoryStream(ReadZipEntry(package, partName), writable: false);
            foreach (XlsbRecord record in XlsbRecordReader.ReadAll(part).Where(record => record.Type == 0)) {
                var cursor = new XlsbBinaryCursor(record.Data);
                if (cursor.ReadInt32() != row - 1) continue;
                cursor.Skip(9);
                uint count = cursor.ReadUInt32();
                var spans = new List<(int FirstColumn, int LastColumn)>();
                for (uint index = 0; index < count; index++) {
                    spans.Add((cursor.ReadInt32(), cursor.ReadInt32()));
                }
                Assert.Equal(0, cursor.Remaining);
                return spans.AsReadOnly();
            }

            throw new Xunit.Sdk.XunitException($"Row {row} was not found in '{partName}'.");
        }

        private static void WriteZipEntry(ZipArchive archive, string name, string content) {
            WriteZipEntry(archive, name, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content));
        }

        private static void WriteZipEntry(ZipArchive archive, string name, byte[] content) {
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Fastest);
            using Stream stream = entry.Open();
            stream.Write(content, 0, content.Length);
        }

        private static byte[] AddZipEntry(byte[] package, string name, byte[] content) {
            using var stream = new MemoryStream();
            stream.Write(package, 0, package.Length);
            stream.Position = 0;
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
                WriteZipEntry(archive, name, content);
            }

            return stream.ToArray();
        }

        private static byte[] ReadZipEntry(byte[] package, string name) {
            using var stream = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            ZipArchiveEntry entry = Assert.Single(archive.Entries,
                candidate => string.Equals(candidate.FullName, name, StringComparison.OrdinalIgnoreCase));
            using Stream input = entry.Open();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] ReplaceZipEntry(byte[] package, string name, byte[] content) {
            using var stream = new MemoryStream();
            stream.Write(package, 0, package.Length);
            stream.Position = 0;
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry entry = Assert.Single(archive.Entries,
                    candidate => string.Equals(candidate.FullName, name, StringComparison.OrdinalIgnoreCase));
                entry.Delete();
                WriteZipEntry(archive, name, content);
            }

            return stream.ToArray();
        }

    }
}
