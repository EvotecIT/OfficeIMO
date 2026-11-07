using OfficeIMO.Excel;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

namespace OfficeIMO.LegacyImport.Tests;

public sealed partial class EarlyBiffImportTests {
    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void ComparisonBeforeConcatenationRetainsGroupingAndTheOriginalNumericCache(int generation) {
        using var source = new MemoryStream();
        using (var writer = new BinaryWriter(source, Encoding.UTF8, leaveOpen: true)) {
            Record(writer, BofCode(generation), record => WriteBof(record, generation, 0x10));
            Record(writer, generation switch { 2 => (ushort)6, 3 => (ushort)0x206, _ => (ushort)0x406 }, record => {
                WriteCell(record, generation, 0);
                record.Write(77d);
                byte[] tokens = { 0x1E, 1, 0, 0x1E, 2, 0, 0x0B, 0x17, 1, (byte)'x', 0x08 };
                if (generation == 2) { record.Write((byte)0); record.Write((byte)tokens.Length); }
                else { record.Write((ushort)0); record.Write((ushort)tokens.Length); }
                record.Write(tokens);
            });
            Record(writer, 0x000A, _ => { });
        }
        using ExcelDocument document = ExcelDocument.Load(source);
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { AllowLoss = true });
        output.Position = 0;
        using var zip = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Read);
        using Stream part = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cell = Assert.Single(System.Xml.Linq.XDocument.Load(part).Descendants(ns + "c"));
        Assert.Equal("(1=2)&\"x\"", (string?)cell.Element(ns + "f"));
        Assert.Equal("77", (string?)cell.Element(ns + "v"));
        Assert.Equal("n", (string?)cell.Attribute("t"));
    }

    [Fact]
    public void Biff20InlineAttributesRetainFontAlignmentBordersFillAndProtectionWithoutAnXfTable() {
        using var source = new MemoryStream();
        using (var writer = new BinaryWriter(source, Encoding.UTF8, leaveOpen: true)) {
            Record(writer, BofCode(2), record => WriteBof(record, 2, 0x10));
            foreach (string fontName in new[] { "Arial", "Courier New" }) {
                Record(writer, 0x0031, record => {
                    record.Write((ushort)(fontName == "Arial" ? 200 : 280));
                    record.Write((ushort)(fontName == "Arial" ? 0 : 3));
                    WriteString(record, Encoding.ASCII.GetBytes(fontName), shortLength: true);
                });
            }
            Record(writer, 0x0024, record => {
                record.Write((byte)0);
                record.Write((byte)0);
                record.Write((ushort)3456);
            });
            for (ushort column = 0; column < 2; column++) {
                Record(writer, 0x0003, record => {
                    record.Write((ushort)0);
                    record.Write(column);
                    record.Write(column == 0 ? new byte[] { 0xC0, 0x40, 0xA3 } : new byte[] { 0x40, 0, 0 });
                    record.Write(42.25d + column);
                });
            }
            Record(writer, 0x0008, record => {
                record.Write((ushort)1);
                record.Write((ushort)0);
                record.Write((ushort)2);
                record.Write((ushort)300);
                record.Write((ushort)0);
                record.Write((byte)1);
                record.Write((ushort)0);
                record.Write(new byte[] { 0x40, 0, 2 });
            });
            Record(writer, 0x000A, _ => { });
        }
        using ExcelDocument document = ExcelDocument.Load(source);
        document.Sheets[0].CellValue(3, 1, "new");
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { AllowLoss = true });
        output.Position = 0;
        using var zip = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Read);
        using Stream part = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using Stream stylePart = zip.GetEntry("xl/styles.xml")!.Open();
        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var worksheet = System.Xml.Linq.XDocument.Load(part);
        var cells = worksheet.Descendants(ns + "c").ToArray();
        var styles = System.Xml.Linq.XDocument.Load(stylePart).Root!;
        var formats = styles.Element(ns + "cellXfs")!.Elements(ns + "xf").ToArray();
        var applied = formats[(int?)cells[0].Attribute("s") ?? 0];
        var control = formats[(int?)cells[1].Attribute("s") ?? 0];
        var font = styles.Element(ns + "fonts")!.Elements(ns + "font").ElementAt((int)applied.Attribute("fontId")!);
        var controlFont = styles.Element(ns + "fonts")!.Elements(ns + "font").ElementAt((int)control.Attribute("fontId")!);
        var row = Assert.Single(worksheet.Descendants(ns + "row"), item => (string?)item.Attribute("r") == "2");
        var rowFormat = formats[(int)row.Attribute("s")!];
        var rowFont = styles.Element(ns + "fonts")!.Elements(ns + "font").ElementAt((int)rowFormat.Attribute("fontId")!);
        Assert.Equal("42.25", (string?)cells[0].Element(ns + "v"));
        Assert.Equal("43.25", (string?)cells[1].Element(ns + "v"));
        Assert.Equal("Courier New", (string?)font.Element(ns + "name")!.Attribute("val"));
        Assert.Equal("14", (string?)font.Element(ns + "sz")!.Attribute("val"));
        Assert.NotNull(font.Element(ns + "b"));
        Assert.NotNull(font.Element(ns + "i"));
        Assert.Equal("Arial", (string?)controlFont.Element(ns + "name")!.Attribute("val"));
        Assert.Equal("right", (string?)applied.Element(ns + "alignment")!.Attribute("horizontal"));
        Assert.Equal("1", (string?)applied.Element(ns + "protection")!.Attribute("locked"));
        Assert.Equal("1", (string?)applied.Element(ns + "protection")!.Attribute("hidden"));
        Assert.Equal("0", (string?)control.Element(ns + "protection")!.Attribute("hidden"));
        var fill = styles.Element(ns + "fills")!.Elements(ns + "fill").ElementAt((int)applied.Attribute("fillId")!);
        Assert.Equal("gray125", (string?)fill.Element(ns + "patternFill")!.Attribute("patternType"));
        var border = styles.Element(ns + "borders")!.Elements(ns + "border").ElementAt((int)applied.Attribute("borderId")!);
        Assert.Equal("thin", (string?)border.Element(ns + "top")!.Attribute("style"));
        Assert.Equal("FF000000", (string?)border.Element(ns + "top")!.Element(ns + "color")!.Attribute("rgb"));
        var columnLayout = Assert.Single(worksheet.Descendants(ns + "col"));
        Assert.Equal("13.5", (string?)columnLayout.Attribute("width"));
        Assert.Null(columnLayout.Attribute("style"));
        var added = Assert.Single(cells, cell => (string?)cell.Attribute("r") == "A3");
        var addedFormat = formats[(int?)added.Attribute("s") ?? 0];
        var addedFont = styles.Element(ns + "fonts")!.Elements(ns + "font").ElementAt((int)addedFormat.Attribute("fontId")!);
        Assert.Equal("Arial", (string?)addedFont.Element(ns + "name")!.Attribute("val"));
        Assert.Equal("0", (string?)addedFormat.Attribute("fillId"));
        Assert.Equal("0", (string?)addedFormat.Attribute("borderId"));
        Assert.Equal("Arial", (string?)rowFont.Element(ns + "name")!.Attribute("val"));
        Assert.Equal("center", (string?)rowFormat.Element(ns + "alignment")!.Attribute("horizontal"));
    }

    [Fact]
    public void FormulaExpansionBeyondExcelLimitRetainsOnlyTheOriginalCache() {
        using var expression = new MemoryStream();
        using (var tokens = new BinaryWriter(expression, Encoding.UTF8, leaveOpen: true)) {
            for (int index = 0; index < 16; index++) {
                tokens.Write((byte)0x17);
                tokens.Write((byte)255);
                tokens.Write(Enumerable.Repeat((byte)'"', 255).ToArray());
                if (index != 0) tokens.Write((byte)0x08);
            }
        }
        using var source = new MemoryStream();
        using (var writer = new BinaryWriter(source, Encoding.UTF8, leaveOpen: true)) {
            Record(writer, 0x0409, record => WriteBof(record, 4, 0x10));
            Record(writer, 0x0406, record => {
                WriteCell(record, 4, 0);
                record.Write(77d);
                record.Write((ushort)0);
                record.Write((ushort)expression.Length);
                record.Write(expression.ToArray());
            });
            Record(writer, 0x000A, _ => { });
        }
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "early-biff-formula-unsupported");
        var cell = Assert.Single(Assert.Single(document.Inspect().Worksheets).Cells);
        Assert.Equal(77d, cell.Value);
        Assert.Null(cell.Formula);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void FormulaProjectionPreservesExpressionsAndEachCacheTypeWithoutEvaluation(int generation) {
        using var source = new MemoryStream();
        using (var writer = new BinaryWriter(source, Encoding.UTF8, leaveOpen: true)) {
            Record(writer, BofCode(generation), record => WriteBof(record, generation, 0x10));
            Record(writer, 0x0042, record => record.Write((ushort)1252));
            byte[] call = generation == 4 ? new byte[] { 0x41, 24, 0 } : new byte[] { 0x41, 24 };
            var expressions = new[] {
                new byte[] { 0x44, 1, 0xC0, 0 }.Concat(call).Concat(new byte[] { 0x1E, 2, 0, 0x03 }).ToArray(),
                new byte[] { 0x1D, 1 }, new byte[] { 0x1C, 7 }, new byte[] { 0x17, 4, 67, 97, 102, 233 }
            };
            for (int column = 0; column < expressions.Length; column++) {
                Record(writer, generation switch { 2 => (ushort)6, 3 => (ushort)0x206, _ => (ushort)0x406 }, record => {
                    WriteCell(record, generation, (ushort)column);
                    if (column == 0) record.Write(77d);
                    else record.Write(new byte[] { (byte)(column == 1 ? 1 : column == 2 ? 2 : 0), 0, (byte)(column == 2 ? 7 : 1), 0, 0, 0, 255, 255 });
                    if (generation == 2) { record.Write((byte)0); record.Write((byte)expressions[column].Length); }
                    else { record.Write((ushort)0); record.Write((ushort)expressions[column].Length); }
                    record.Write(expressions[column]);
                });
                if (column == 3) Record(writer, generation == 2 ? (ushort)7 : (ushort)0x207,
                    record => WriteString(record, new byte[] { 67, 97, 102, 233 }, shortLength: generation == 2));
            }
            Record(writer, 0x000A, _ => { });
        }
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.DoesNotContain(report.Diagnostics, finding => finding.Code == "early-biff-formula-cache");
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { AllowLoss = true });
        output.Position = 0;
        using var zip = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Read);
        using Stream part = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cells = System.Xml.Linq.XDocument.Load(part).Descendants(ns + "c").ToArray();
        Assert.Equal(new[] { "ABS(A2)+2", "TRUE", "#DIV/0!", "\"Café\"" }, cells.Select(cell => (string?)cell.Element(ns + "f")));
        Assert.Equal(new[] { "77", "1", "#DIV/0!", "Café" }, cells.Select(cell => (string?)cell.Element(ns + "v")));
        Assert.Equal(new[] { "n", "b", "e", "str" }, cells.Select(cell => (string?)cell.Attribute("t")));
        using var readerSource = new MemoryStream(source.ToArray());
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(readerSource,
            new ExcelReaderOptions { HasHeaderRow = false, UseCachedFormulaResult = false });
        Assert.True(reader.Read());
        Assert.Equal("ABS(A2)+2", reader.GetString(0));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void WorksheetRetainsTypedCellsAndFormulaExpressionsWithTheirOriginalCaches(int generation) {
        using var source = new MemoryStream(BuildWorksheet(generation));
        using ExcelDocument imported = ExcelDocument.Load(source, null, out OfficeConversionReport report);

        Assert.Equal(ExcelFileFormat.Xls, imported.SourceFormat);
        Assert.DoesNotContain(report.Diagnostics, item => item.Code == "early-biff-formula-cache");
        using var xlsx = new MemoryStream();
        imported.Save(xlsx, new ExcelSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        xlsx.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(xlsx);
        var cells = Assert.Single(reopened.Inspect().Worksheets).Cells;

        Assert.Equal("Café", Assert.Single(cells, item => item.Column == 1).Value);
        Assert.Equal(42.25d, Assert.Single(cells, item => item.Column == 2).Value);
        var error = Assert.Single(cells, item => item.Column == 3);
        Assert.Equal("#DIV/0!", error.Value);
        Assert.Equal(ExcelCellDataKind.Error, Assert.Single(reopened.Sheets).Cell(1, 3).GetValue().Kind);
        var formula = Assert.Single(cells, item => item.Column == 4);
        Assert.Equal("cached", formula.Value);
        Assert.Equal("1", formula.Formula);
    }

    [Fact]
    public void WorkbookRetainsEmbeddedSheetOrderAndSkipsMacroSheetsWithNestedCharts() {
        using var source = new MemoryStream(BuildWorkbook());
        using ExcelDocument imported = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Equal(new[] { "First", "Second" }, imported.Sheets.Select(sheet => sheet.Name));
        Assert.Contains(report.Diagnostics, item => item.Code == "early-biff-macro-sheet");
        Assert.Contains(report.Diagnostics, item => item.Code == "early-biff-embedded-chart");

        using var output = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => imported.Save(output, new ExcelSaveOptions { Format = ExcelFileFormat.Xls }));
        Assert.Equal(0, output.Length);

        source.Position = 0;
        OfficePackageSecurityException rejected = Assert.Throws<OfficePackageSecurityException>(
            () => ExcelDocument.Load(source, ExcelLoadOptions.UntrustedDefaults));
        Assert.Equal(OfficePackageSecurityRule.Macros, rejected.Rule);
    }

    [Fact]
    public void ForwardReaderAndDocumentReaderAdmitRawBiffThroughExistingExcelRoutes() {
        byte[] bytes = BuildWorksheet(3);
        using var source = new MemoryStream();
        source.WriteByte(0xCC);
        source.Write(bytes, 0, bytes.Length);
        source.Position = 1;
        using (ExcelWorkbookDataReader rows = ExcelDocument.OpenReader(source, new ExcelReaderOptions { HasHeaderRow = false })) {
            Assert.True(rows.Read());
            Assert.Equal("Café", rows.GetString(0));
            Assert.Equal(42.25d, rows.GetDouble(1));
        }
        Assert.Equal(1, source.Position);

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler().Build();
        using var readerSource = new MemoryStream(bytes);
        OfficeDocumentReadResult result = reader.ReadDocument(readerSource, "book.xls");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Café", StringComparison.Ordinal));
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("excel-biff3-worksheet", StringComparison.Ordinal));
    }

    [Fact]
    public void RecognizedWorksheetRejectsTruncatedRecordsAndHonorsInputLimitAndCancellation() {
        byte[] bytes = BuildWorksheet(4);
        using var truncated = new MemoryStream(bytes, 0, bytes.Length - 1);
        var malformed = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(truncated));
        Assert.Contains("Truncated early BIFF record header", malformed.Message, StringComparison.Ordinal);

        using var bounded = new MemoryStream(bytes);
        var limit = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(bounded,
            new ExcelLoadOptions { MaxLegacyInputBytes = bytes.Length - 1 }));
        Assert.Contains("configured maximum size", limit.Message, StringComparison.Ordinal);

        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        using var source = new MemoryStream(bytes);
        Assert.Throws<OperationCanceledException>(() => ExcelDocument.Load(source, null, canceled.Token));
    }


    [Fact]
    public void FileConversionRequiresExplicitAcceptanceOfCachedFormulaLoss() {
        string stem = Path.Combine(Path.GetTempPath(), "OfficeIMO-EarlyBiff-" + Guid.NewGuid().ToString("N"));
        string source = stem + ".xls";
        string output = stem + ".xlsx";
        byte[] bytes = BuildWorksheet(4, unsupportedFormula: true);
        try {
            File.WriteAllBytes(source, bytes);
            var blocked = Assert.Throws<ExcelDocumentConversionException>(() => ExcelDocument.Convert(source, output));
            Assert.Equal(OfficeConversionFailureReason.DataLossBlocked, blocked.Reason);
            Assert.False(File.Exists(output));

            var allowed = ExcelDocument.Convert(source, output, new ExcelDocumentConversionOptions {
                LossPolicy = OfficeConversionLossPolicy.Allow
            });
            Assert.Contains(allowed.Report.Diagnostics, item => item.Code == "early-biff-formula-cache" && item.LossKind != OfficeConversionLossKind.None);
            Assert.Equal(bytes, File.ReadAllBytes(source));
            using ExcelDocument converted = ExcelDocument.Load(output);
            Assert.Equal("cached", Assert.Single(converted.Sheets).Cell(1, 4).GetValue<string>());
        } finally {
            File.Delete(source);
            File.Delete(output);
        }
    }


    [Fact]
    public async Task NonseekableAdmissionStopsAtRawBiffLimitBeforeReadingTail() {
        byte[] raw = BuildWorksheet(4);
        const int limit = 24;
        using var sync = new ForwardInput(raw);
        Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(sync, new ExcelLoadOptions { MaxLegacyInputBytes = limit }));
        Assert.Equal(limit + 1, sync.BytesRead);
        using var asyncSource = new ForwardInput(raw);
        await Assert.ThrowsAsync<InvalidDataException>(() => ExcelDocument.LoadAsync(
            asyncSource, new ExcelLoadOptions { MaxLegacyInputBytes = limit }));
        Assert.Equal(limit + 1, asyncSource.BytesRead);
    }

    [Fact]
    public async Task DataReaderReportsTheEffectiveRawBiffLimitForSeekableInput() {
        const long rawLimit = 64L * 1024 * 1024;
        using var input = new SizedInput(BuildWorksheet(4), rawLimit + 1);
        var sync = Assert.Throws<OfficeInspectionResourceLimitException>(() => ExcelDocument.OpenReader(input));
        Assert.Equal(nameof(ExcelLoadOptions.MaxLegacyInputBytes), sync.LimitName);
        Assert.Equal(rawLimit, sync.Maximum);

        var asyncFailure = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() => ExcelDocument.OpenReaderAsync(input));
        Assert.Equal(sync.LimitName, asyncFailure.LimitName);
        Assert.Equal(rawLimit, asyncFailure.Maximum);
    }

    [Fact]
    public async Task HttpLoadChecksRawBiffLimitBeforeReadingTheResponseTail() {
        byte[] bytes = BuildWorksheet(4);
        const int limit = 24;
        var uri = new Uri("https://example.test/early.xls");
        using var partial = new ForwardInput(bytes);
        using var boundedHandler = new RawWorkbookHandler(partial);
        var error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() => ExcelDocument.LoadAsync(uri,
            new ExcelLoadOptions { HttpMessageHandler = boundedHandler, MaxLegacyInputBytes = limit }));
        Assert.Equal(nameof(ExcelLoadOptions.MaxLegacyInputBytes), error.LimitName);
        Assert.Equal(limit + 1, partial.BytesRead);

        using var completeHandler = new RawWorkbookHandler(new ForwardInput(bytes));
        using ExcelDocument workbook = await ExcelDocument.LoadAsync(uri,
            new ExcelLoadOptions { HttpMessageHandler = completeHandler });
        Assert.Equal("Café", Assert.Single(workbook.Sheets).Cell(1, 1).GetValue<string>());
    }

    [Fact]
    public async Task AsyncReaderRecognizesOnlyEmbeddedBiff4WorkbooksUnderXlw() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-EarlyBiff-" + Guid.NewGuid().ToString("N") + ".xlw");
        try {
            File.WriteAllBytes(path, BuildWorkbook());
            using (ExcelWorkbookDataReader reader = await ExcelDocument.OpenReaderAsync(path, new ExcelReaderOptions { HasHeaderRow = false })) {
                Assert.Equal(new[] { "First", "Second" }, reader.SheetNames);
                Assert.True(reader.Read());
                Assert.Equal("Café", reader.GetString(0));
            }

            using var noCaches = new MemoryStream(BuildWorksheet(4, unsupportedFormula: true));
            Assert.Throws<NotSupportedException>(() => ExcelDocument.OpenReader(noCaches,
                new ExcelReaderOptions { UseCachedFormulaResult = false }));
            File.WriteAllBytes(path, BuildWorksheet(4));
            await Assert.ThrowsAsync<NotSupportedException>(() => ExcelDocument.OpenReaderAsync(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task SavingProtectsSourceIdentityAndRequiresExplicitLossAcceptanceBeforePublication() {
        byte[] raw = BuildWorksheet(4, unsupportedFormula: true);
        string stem = Path.Combine(Path.GetTempPath(), "OfficeIMO-EarlyBiff-" + Guid.NewGuid().ToString("N"));
        string original = stem + ".xlsx";
        string wrongExtension = stem + ".xlw";
        try {
            File.WriteAllBytes(original, raw);
            using ExcelDocument fromFile = ExcelDocument.Load(original);
            var allow = new ExcelSaveOptions { AllowLoss = true };
            var encrypted = new ExcelSaveOptions { Password = "validation", AllowLoss = true };
            string alias = Path.Combine(Path.GetDirectoryName(original)!, ".", Path.GetFileName(original));
            Assert.Throws<NotSupportedException>(() => fromFile.SaveCopy(alias, allow));
            Assert.Throws<NotSupportedException>(() => fromFile.Save(alias, encrypted));
            await Assert.ThrowsAsync<NotSupportedException>(() => fromFile.SaveAsync(wrongExtension, encrypted));
            Assert.Equal(raw, File.ReadAllBytes(original));
            Assert.False(File.Exists(wrongExtension));

            using var originalStream = new MemoryStream((byte[])raw.Clone(), writable: true);
            using ExcelDocument fromStream = ExcelDocument.Load(originalStream);
            await Assert.ThrowsAsync<NotSupportedException>(() => fromStream.SaveAsync(originalStream, encrypted));
            Assert.Equal(raw, originalStream.ToArray());

            using var destination = new MemoryStream();
            NotSupportedException refusal = Assert.Throws<NotSupportedException>(() => fromFile.Save(destination));
            Assert.Contains(nameof(ExcelSaveOptions.AllowLoss), refusal.Message);
            Assert.Throws<NotSupportedException>(() => fromFile.Save(destination,
                new ExcelSaveOptions { Password = "validation", Format = ExcelFileFormat.Xls, LossPolicy = OfficeConversionLossPolicy.Allow }));
            Assert.Equal(0, destination.Length);
        } finally {
            File.Delete(original);
            File.Delete(wrongExtension);
        }
    }

    [Theory]
    [InlineData(3)]
    [InlineData(4)]
    public void WorksheetPreservesInheritedDatesBlanksAndLayoutAcrossAnEmbeddedChart(int generation) {
        using var source = new MemoryStream();
        using (var output = new BinaryWriter(source, Encoding.UTF8, leaveOpen: true)) {
            Record(output, BofCode(generation), record => WriteBof(record, generation, 0x10));
            // An own-sheet reference is an internal link, even under the strict external-relationship policy.
            Record(output, 0x0017, record => record.Write(new byte[] { 1, 3, (byte)'A' }));
            ushort formatRecord = generation == 3 ? (ushort)0x001E : (ushort)0x041E;
            foreach (string format in new[] { "General", "yyyy-mm-dd" }) {
                Record(output, formatRecord, record => {
                    if (generation == 4) record.Write((ushort)0);
                    WriteString(record, Encoding.ASCII.GetBytes(format), shortLength: true);
                });
            }
            Record(output, 0x0231, record => {
                record.Write((ushort)240);
                record.Write((ushort)3);
                record.Write((ushort)4);
                WriteString(record, Encoding.ASCII.GetBytes("Arial"), shortLength: true);
            });
            ushort xf = generation == 3 ? (ushort)0x0243 : (ushort)0x0443;
            byte[] parent = new byte[12];
            parent[1] = 1;
            parent[2] = 4;
            parent[4] = 10;
            ushort background = 1 | (3 << 6) | (9 << 11);
            parent[6] = (byte)background;
            parent[7] = (byte)(background >> 8);
            parent[8] = 1 | (2 << 3);
            Record(output, xf, record => record.Write(parent));
            Record(output, xf, record => record.Write(new byte[12]));
            Record(output, 0x0208, record => {
                record.Write((ushort)4);
                record.Write((ushort)0);
                record.Write((ushort)5);
                record.Write((ushort)400);
                record.Write(0);
                record.Write((uint)(32 | 16 | 128 | 2 | (1 << 16)));
            });
            foreach (var definition in new[] { (First: 1, Last: 3, Width: 3456, Flags: 0x1301), (First: 2, Last: 2, Width: 4608, Flags: 0) }) {
                Record(output, 0x007D, record => {
                    record.Write((ushort)definition.First);
                    record.Write((ushort)definition.Last);
                    record.Write((ushort)definition.Width);
                    record.Write((ushort)1);
                    record.Write((ushort)definition.Flags);
                    record.Write((ushort)0);
                });
            }
            Record(output, 0x0201, record => {
                record.Write((ushort)4);
                record.Write((ushort)4);
                record.Write((ushort)0);
            });
            Record(output, BofCode(generation), record => WriteBof(record, generation, 0x20));
            Record(output, 0x000A, _ => { });
            Record(output, 0x0203, record => {
                record.Write((ushort)0);
                record.Write((ushort)0);
                record.Write((ushort)1);
                record.Write(36525d);
            });
            Record(output, 0x000A, _ => { });
        }

        using ExcelDocument workbook = ExcelDocument.Load(source, ExcelLoadOptions.UntrustedDefaults, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, finding => finding.Code == "early-biff-embedded-chart");
        using var result = new MemoryStream();
        workbook.Save(result, new ExcelSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        result.Position = 0;
        using var package = new System.IO.Compression.ZipArchive(result, System.IO.Compression.ZipArchiveMode.Read);
        using Stream sheet = package.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using Stream stylePart = package.GetEntry("xl/styles.xml")!.Open();
        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var worksheet = System.Xml.Linq.XDocument.Load(sheet);
        var cells = worksheet.Descendants(ns + "c").ToArray();
        var rowLayout = Assert.Single(worksheet.Descendants(ns + "row"), row => (string?)row.Attribute("r") == "5");
        Assert.Equal("20", (string?)rowLayout.Attribute("ht"));
        Assert.Equal("1", (string?)rowLayout.Attribute("hidden"));
        Assert.Equal("2", (string?)rowLayout.Attribute("outlineLevel"));
        var columnLayouts = worksheet.Descendants(ns + "col").ToArray();
        Assert.Equal(new[] { "2", "3", "4" }, columnLayouts.Select(column => (string?)column.Attribute("min")));
        Assert.Equal(new[] { "13.5", "18", "13.5" }, columnLayouts.Select(column => (string?)column.Attribute("width")));
        Assert.Equal(new[] { "1", "0", "1" }, columnLayouts.Select(column => (string?)column.Attribute("hidden")));
        var blank = Assert.Single(cells, cell => (string?)cell.Attribute("r") == "E5");
        Assert.Null(blank.Element(ns + "v"));

        var date = Assert.Single(cells, cell => (string?)cell.Attribute("r") == "A1");
        Assert.Equal("36525", (string?)date.Element(ns + "v"));
        var styles = System.Xml.Linq.XDocument.Load(stylePart).Root!;
        var appliedFormat = styles.Element(ns + "cellXfs")!.Elements(ns + "xf").ElementAt((int)date.Attribute("s")!);
        var expectedFormat = Assert.Single(styles.Element(ns + "numFmts")!.Elements(ns + "numFmt"),
            format => (string?)format.Attribute("formatCode") == "yyyy-mm-dd");
        Assert.Equal((string?)expectedFormat.Attribute("numFmtId"), (string?)appliedFormat.Attribute("numFmtId"));
        var font = styles.Element(ns + "fonts")!.Elements(ns + "font").ElementAt((int)appliedFormat.Attribute("fontId")!);
        Assert.Equal("Arial", (string?)font.Element(ns + "name")!.Attribute("val"));
        Assert.Equal("12", (string?)font.Element(ns + "sz")!.Attribute("val"));
        Assert.NotNull(font.Element(ns + "b"));
        Assert.NotNull(font.Element(ns + "i"));
        Assert.Equal("center", (string?)appliedFormat.Element(ns + "alignment")!.Attribute("horizontal"));
        Assert.Equal("1", (string?)appliedFormat.Element(ns + "alignment")!.Attribute("wrapText"));
        var fill = styles.Element(ns + "fills")!.Elements(ns + "fill").ElementAt((int)appliedFormat.Attribute("fillId")!);
        Assert.Equal("solid", (string?)fill.Element(ns + "patternFill")!.Attribute("patternType"));
        var border = styles.Element(ns + "borders")!.Elements(ns + "border").ElementAt((int)appliedFormat.Attribute("borderId")!);
        Assert.Equal("thin", (string?)border.Element(ns + "top")!.Attribute("style"));
    }

    private sealed class RawWorkbookHandler : System.Net.Http.HttpMessageHandler {
        private readonly Stream _source;
        internal RawWorkbookHandler(Stream source) => _source = source;
        protected override Task<System.Net.Http.HttpResponseMessage> SendAsync(System.Net.Http.HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(new System.Net.Http.HttpResponseMessage(System.Net.HttpStatusCode.OK) {
                Content = new System.Net.Http.StreamContent(_source)
            });
    }

    private sealed class ForwardInput : MemoryStream {
        internal ForwardInput(byte[] bytes) : base(bytes, writable: false) { }
        public override bool CanSeek => false;
        internal int BytesRead { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(Read(buffer, offset, count));
        }
    }

    private sealed class SizedInput : MemoryStream {
        internal SizedInput(byte[] header, long sourceLength) : base(header, writable: false) => Length = sourceLength;
        public override long Length { get; }
    }

    private static byte[] BuildWorksheet(int generation, ushort kind = 0x10, bool unsupportedFormula = false) {
        using var bytes = new MemoryStream();
        using var output = new BinaryWriter(bytes);
        Record(output, BofCode(generation), record => WriteBof(record, generation, kind));
        Record(output, 0x0042, record => record.Write((ushort)1252));

        Record(output, generation == 2 ? (ushort)0x0004 : (ushort)0x0204, record => {
            WriteCell(record, generation, 0);
            WriteString(record, new byte[] { 0x43, 0x61, 0x66, 0xE9 }, shortLength: generation == 2);
        });
        Record(output, generation == 2 ? (ushort)0x0003 : (ushort)0x0203, record => {
            WriteCell(record, generation, 1);
            record.Write(42.25d);
        });
        Record(output, generation == 2 ? (ushort)0x0005 : (ushort)0x0205, record => {
            WriteCell(record, generation, 2);
            record.Write((byte)0x07);
            record.Write((byte)1);
        });
        Record(output, generation switch { 2 => (ushort)0x0006, 3 => (ushort)0x0206, _ => (ushort)0x0406 }, record => {
            WriteCell(record, generation, 3);
            record.Write(new byte[] { 0, 0, 0, 0, 0, 0, 0xFF, 0xFF });
            if (generation == 2) {
                record.Write((byte)0);
                record.Write((byte)3);
            } else {
                record.Write((ushort)0);
                record.Write((ushort)3);
            }
            record.Write(unsupportedFormula ? new byte[] { 0xFF, 0, 0 } : new byte[] { 0x1E, 0x01, 0x00 });
        });
        Record(output, generation == 2 ? (ushort)0x0007 : (ushort)0x0207,
            record => WriteString(record, Encoding.ASCII.GetBytes("cached"), shortLength: generation == 2));
        if (kind == 0x0040) {
            Record(output, BofCode(generation), record => WriteBof(record, generation, 0x20));
            Record(output, 0x000A, _ => { });
        }
        Record(output, 0x000A, _ => { });
        return bytes.ToArray();
    }

    private static byte[] BuildWorkbook() {
        using var bytes = new MemoryStream();
        using var output = new BinaryWriter(bytes);
        Record(output, 0x0409, record => WriteBof(record, generation: 4, kind: 0x0100));
        const int firstSheetOffsetField = 14;
        Record(output, 0x008E, record => record.Write(0));
        var names = new[] { "First", "Macro", "Second" };
        foreach (string name in names) {
            Record(output, 0x0085, record => WriteString(record, Encoding.ASCII.GetBytes(name), shortLength: true));
        }
        long firstSheet = bytes.Position;
        bytes.Position = firstSheetOffsetField;
        output.Write((int)firstSheet);
        bytes.Position = firstSheet;
        foreach (string name in names) {
            byte[] sheet = BuildWorksheet(4, name == "Macro" ? (ushort)0x40 : (ushort)0x10);
            Record(output, 0x008F, record => {
                record.Write(sheet.Length);
                WriteString(record, Encoding.ASCII.GetBytes(name), shortLength: true);
            });
            output.Write(sheet);
        }
        Record(output, 0x000A, _ => { });
        return bytes.ToArray();
    }

    private static void Record(BinaryWriter output, ushort type, Action<BinaryWriter> body) {
        using var payload = new MemoryStream();
        using var record = new BinaryWriter(payload);
        body(record);
        output.Write(type);
        output.Write((ushort)payload.Length);
        output.Write(payload.ToArray());
    }

    private static void WriteBof(BinaryWriter output, int generation, ushort kind) {
        output.Write((ushort)0);
        output.Write(kind);
        if (generation != 2) output.Write((ushort)0);
    }

    private static ushort BofCode(int generation) => generation switch { 2 => 0x0009, 3 => 0x0209, _ => 0x0409 };

    private static void WriteCell(BinaryWriter output, int generation, ushort column) {
        output.Write((ushort)0);
        output.Write(column);
        output.Write((ushort)0);
        if (generation == 2) output.Write((byte)0);
    }

    private static void WriteString(BinaryWriter output, byte[] value, bool shortLength) {
        if (shortLength) output.Write((byte)value.Length);
        else output.Write((ushort)value.Length);
        output.Write(value);
    }
}
