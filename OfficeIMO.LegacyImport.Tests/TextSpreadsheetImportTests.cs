using OfficeIMO.Excel;
using OfficeIMO.Excel.Legacy;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class TextSpreadsheetImportTests {
    [Fact]
    public void SylkValuesUseStickyCoordinatesAndKeepFormulaTextLiteralThroughXlsx() {
        const string source = "ID;PTEST\nC;X1;Y1;K\"  =SUM(A1)  \"\nC;X2;K42.25\nC;Y2;KTRUE\nC;X1;K\"Semi;;quoted \"\"text\"\"\"\nE\n";
        using LegacySpreadsheetImportResult imported = Import(source);
        Assert.Equal(LegacySpreadsheetFormat.Sylk, imported.Detection.Format);
        Assert.Equal(OfficeLegacyImportQuality.Structured, imported.Report.Quality);
        Assert.False(imported.HasLoss);
        Assert.Equal("  =SUM(A1)  ", Cell(imported, 1, 1).CachedValue);
        Assert.Equal(42.25, Cell(imported, 1, 2).CachedValue);
        Assert.Equal(true, Cell(imported, 2, 2).CachedValue);
        Assert.Equal("Semi;quoted \"text\"", Cell(imported, 2, 1).CachedValue);
        using var output = new MemoryStream();
        imported.Value.Save(output);
        output.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(output);
        var cells = Assert.Single(reopened.CreateInspectionSnapshot().Worksheets).Cells;
        var literal = Assert.Single(cells, cell => cell.Row == 1 && cell.Column == 1);
        Assert.Equal("  =SUM(A1)  ", literal.Value);
        Assert.Null(literal.Formula);
    }

    [Fact]
    public void SylkOmissionsRetainCachesAndNeverEvaluateUncachedExpressions() {
        using LegacySpreadsheetImportResult imported = Import("ID;PTEST\nP;P0.00\nF;X1;Y1;P0\nC;X1;Y1;K84;E40+44\nC;X2;EHYPERLINK(\"https://example.invalid\")\nC;X3;K99;E1+1;I\nNN;Nnamed;E$A$1\nE\n");
        Assert.Equal(84d, Cell(imported, 1, 1).CachedValue);
        Assert.Null(Cell(imported, 1, 1).Formula);
        Assert.Null(Cell(imported, 1, 2).CachedValue);
        Assert.Null(Cell(imported, 1, 3).CachedValue);
        Assert.Equal("3", imported.Metadata["OmittedFormulaCount"]);
        Assert.Equal("2", imported.Metadata["MissingFormulaValueCount"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "SYLK_FORMULA_STORED_VALUE");
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "SYLK_FORMATTING_OMITTED");
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "SYLK_RECORDS_OMITTED");
        Assert.True(imported.HasLoss);
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireNoLoss());
    }

    [Fact]
    public void SylkFormattingCoordinatesLocateValuesAndExplicitBlanksAreRetained() {
        using LegacySpreadsheetImportResult imported = Import("ID;PTEST\nF;X3;Y4;P0\nC;K42\nC;X4\nE\n");
        Assert.Equal(42d, Cell(imported, 4, 3).CachedValue);
        Assert.Null(Cell(imported, 4, 4).CachedValue);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "SYLK_FORMATTING_OMITTED");
    }

    [Fact]
    public void SylkErrorCachesAndReferencedFormulasHaveExplicitLoss() {
        using LegacySpreadsheetImportResult imported = Import("ID;PTEST\nC;X1;Y1;K#N/A;E1/0\nC;X2;K42;R1;C1\nE\n");
        Assert.Equal("#N/A", Cell(imported, 1, 1).CachedValue);
        Assert.Equal("2", imported.Metadata["OmittedFormulaCount"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "SYLK_ERROR_AS_TEXT");
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "SYLK_FORMULA_STORED_VALUE");
        Assert.Throws<InvalidDataException>(() => Import("ID;PTEST\nC;X1;Y1;K1" + string.Concat(Enumerable.Repeat(";Zignored", 61)) + "\nE\n"));
    }

    [Fact]
    public void DifRecoversTypedValuesEscapedQuotesMultilineTextAndErrors() {
        string source = Dif("0,42.25\nV\n0,1\nTRUE\n0,0\nFALSE\n1,0\n\"semi; \"\"quote\"\"\nand newline\"\n1,0\n\"=SUM(A1)\"\n0,0\nERROR\n", 6, 1);
        using LegacySpreadsheetImportResult imported = Import(source);
        Assert.Equal(LegacySpreadsheetFormat.Dif, imported.Detection.Format);
        Assert.Equal(42.25, Cell(imported, 1, 1).CachedValue);
        Assert.Equal(true, Cell(imported, 1, 2).CachedValue);
        Assert.Equal(false, Cell(imported, 1, 3).CachedValue);
        Assert.Equal("semi; \"quote\"\nand newline", Cell(imported, 1, 4).CachedValue);
        Assert.Equal("=SUM(A1)", Cell(imported, 1, 5).CachedValue);
        Assert.Null(Cell(imported, 1, 5).Formula);
        Assert.Equal("ERROR", Cell(imported, 1, 6).CachedValue);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "DIF_ERROR_AS_TEXT");
    }

    [Theory]
    [InlineData("slk", "Semi;quoted \"value\"")]
    [InlineData("dif", "Semi;quoted \"value\"")]
    public void ImportsIndependentLibreOfficeOutputAndReaderTables(string extension, string expected) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "TextSpreadsheets", "libreoffice-cells." + extension));
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(source, new LegacySpreadsheetImportOptions { RequireStructured = true });
        Assert.Equal(9, imported.Cells.Count);
        Assert.Equal(expected, Cell(imported, 2, 1).CachedValue);
        Assert.Equal("https://example.invalid", Cell(imported, 3, 1).CachedValue);
        Assert.Null(Cell(imported, 3, 1).Formula);
        if (extension == "slk") Assert.Contains(imported.Report.Findings, finding => finding.Code == "SYLK_FORMULA_STORED_VALUE");
        else Assert.False(imported.HasLoss);

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelAndLegacyHandlers().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(source, "book." + extension);
        Assert.Equal(ReaderInputKind.Excel, result.Kind);
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>()), table => table.Rows.Any(row => row.Contains(expected)));
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()), warning => warning.Contains(imported.Detection.ProfileId, StringComparison.Ordinal));
    }

    [Fact]
    public void StreamsUseTheirCurrentPositionPreserveOwnershipAndHonorBom() {
        const string text = "ID;PTEST\r\nC;X1;Y1;K\"Zażółć\"\r\nE\r\n";
        byte[] source = new byte[] { 1, 2, 3 }.Concat(Encoding.Unicode.GetPreamble()).Concat(Encoding.Unicode.GetBytes(text)).ToArray();
        using var stream = new MemoryStream(source);
        stream.Position = 3;
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(stream);
        Assert.Equal("Zażółć", Assert.Single(imported.Cells).CachedValue);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void ReaderSnapshotsEncodingAndImportOptionsForTextInterchange() {
        var options = new LegacySpreadsheetImportOptions { TextEncoding = new UTF8Encoding(false, true) };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacySpreadsheetHandler(options).Build();
        options.TextEncoding = Encoding.ASCII;
        options.Limits.MaxItems = 1;
        byte[] source = Encoding.UTF8.GetBytes("ID;PTEST\nC;X1;Y1;K\"Łódź\"\nC;X2;K42\nE\n");
        OfficeDocumentReadResult result = reader.ReadDocument(source, "text.slk", new ReaderOptions { TextEncoding = Encoding.ASCII });
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Łódź", StringComparison.Ordinal));
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("42", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("ID;PTEST\nC;X0;Y1;K1\nE\n")]
    [InlineData("ID;PTEST\nC;X16385;Y1;K1\nE\n")]
    [InlineData("ID;PTEST\nC;X1;Y1048577;K1\nE\n")]
    [InlineData("ID;PTEST\nC;X1;Y1;K1\nC;X1;Y1;K2\nE\n")]
    [InlineData("ID;PTEST\nC;X1;Y1;KInfinity\nE\n")]
    [InlineData("ID;PTEST\nC;X1;Y1;K\"unterminated\nE\n")]
    [InlineData("ID;PTEST\nC;X1;Y1;K1\n")]
    [InlineData("ID;PTEST\nE\nC;X1;Y1;K1\n")]
    public void RejectsMalformedSylkInsteadOfReturningApparentlyCompleteValues(string source) =>
        Assert.Throws<InvalidDataException>(() => Import(source));

    [Theory]
    [InlineData("0,1e999\nV\n")]
    [InlineData("0,1\nFALSE\n")]
    [InlineData("2,0\n\"unknown\"\n")]
    [InlineData("-1,0\nOTHER\n")]
    [InlineData("1,0\n\"unterminated\n")]
    public void RejectsMalformedDifDatasets(string body) =>
        Assert.Throws<InvalidDataException>(() => Import(Dif(body, 1, 1)));

    [Fact]
    public void EmptyInterchangeWorkbooksAreStructuredAndDimensionsDoNotAllocateCells() {
        using LegacySpreadsheetImportResult sylk = Import("ID;PTEST\nB;X16384;Y1048576\nE\n");
        Assert.Empty(sylk.Cells);
        Assert.Equal(OfficeLegacyImportQuality.Structured, sylk.Report.Quality);
        using LegacySpreadsheetImportResult dif = Import(DifHeader(0, 0) + "-1,0\nEOD\n");
        Assert.Empty(dif.Cells);
        Assert.False(dif.HasLoss);
    }

    [Fact]
    public void TransposedDifDimensionsPreserveRowOrderAndMismatchesAreReported() {
        using LegacySpreadsheetImportResult transposed = Import(Dif("0,2\nV\n0,3\nV\n", 1, 2));
        Assert.Equal(2, transposed.Cells.Count);
        Assert.Equal(1, transposed.Cells[1].Row);
        Assert.Contains("DimensionsProfile", transposed.Metadata.Keys);
        Assert.False(transposed.HasLoss);
        using LegacySpreadsheetImportResult mismatch = Import(Dif("0,2\nV\n", 10, 10));
        Assert.Contains(mismatch.Report.Findings, finding => finding.Code == "DIF_DIMENSIONS_MISMATCH");
    }

    [Fact]
    public void ImportLimitsAndCancellationCoverBothFormatsBeforeWorkbookProjection() {
        foreach (string source in new[] { "ID;PTEST\nC;X1;Y1;K\"abcd\"\nC;X2;K\"efgh\"\nE\n", Dif("1,0\n\"abcd\"\n1,0\n\"efgh\"\n", 2, 1) }) {
            foreach (OfficeLegacyImportLimits limits in new[] {
                new OfficeLegacyImportLimits { MaxInputBytes = 8 },
                new OfficeLegacyImportLimits { MaxItems = 1 },
                new OfficeLegacyImportLimits { MaxTextCharacters = 5 },
                new OfficeLegacyImportLimits { MaxRecords = 1 }
            }) {
                Assert.Throws<InvalidDataException>(() => Import(source, new LegacySpreadsheetImportOptions { Limits = limits }));
            }
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => LegacySpreadsheetImporter.Import(Encoding.UTF8.GetBytes(source), cancellationToken: cancellation.Token));
        }
    }

    [Fact]
    public void InvalidEncodedTextAndOversizedCellsCannotBeSilentlyReplacedOrTruncated() {
        byte[] invalid = Encoding.UTF8.GetPreamble().Concat(Encoding.ASCII.GetBytes("ID;PTEST\nC;X1;Y1;K\"")).Concat(new byte[] { 0xFF }).Concat(Encoding.ASCII.GetBytes("\"\nE\n")).ToArray();
        Assert.Throws<DecoderFallbackException>(() => LegacySpreadsheetImporter.Import(invalid));
        Assert.Throws<InvalidDataException>(() => Import("ID;PTEST\nC;X1;Y1;K\"" + new string('x', 32_768) + "\"\nE\n"));
        Assert.Throws<InvalidDataException>(() => Import(Dif("1,0\n\"" + new string('x', 32_768) + "\"\n", 1, 1)));
    }

    private static LegacySpreadsheetCellContent Cell(LegacySpreadsheetImportResult result, int row, int column) =>
        Assert.Single(result.Cells, cell => cell.Row == row && cell.Column == column);

    private static LegacySpreadsheetImportResult Import(string source, LegacySpreadsheetImportOptions? options = null) =>
        LegacySpreadsheetImporter.Import(Encoding.UTF8.GetBytes(source), options);

    private static string Dif(string body, int columns, int rows) => DifHeader(columns, rows) + "-1,0\nBOT\n" + body + "-1,0\nEOD\n";

    private static string DifHeader(int columns, int rows) =>
        $"TABLE\n0,1\n\"Data\"\nVECTORS\n0,{columns}\n\"\"\nTUPLES\n0,{rows}\n\"\"\nDATA\n0,0\n\"\"\n";
}
