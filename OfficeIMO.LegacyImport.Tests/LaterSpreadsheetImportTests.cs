using OfficeIMO.Excel;
using OfficeIMO.Excel.Legacy;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class LaterSpreadsheetImportTests {
    [Theory]
    [InlineData("Lotus123_98.123", "lotus-1-2-3-later-records", "titre")]
    [InlineData("QuattroPro.wb1", "quattro-pro-wb1-records", "green")]
    [InlineData("Works_Windows.wks", "microsoft-works-wks-windows-records", "a simple sheet")]
    public void ImportsIndependentLaterGenerationWorkbooks(string file, string profile, string text) {
        using var result = LegacySpreadsheetImporter.Import(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Spreadsheets", file)),
            new LegacySpreadsheetImportOptions { RequireStructured = true, SourceName = file });
        Assert.Equal(profile, result.Detection.ProfileId);
        Assert.Equal(OfficeLegacyImportQuality.Structured, result.Report.Quality);
        Assert.Contains(result.Cells, cell => cell.CachedValue as string == text);
        using var xlsx = new MemoryStream(); result.Value.Save(xlsx); xlsx.Position = 0;
        using var loaded = ExcelDocument.Load(xlsx);
        Assert.Equal(result.Value.Sheets.Count, loaded.Sheets.Count);
        Assert.Contains(loaded.CreateInspectionSnapshot().Worksheets.SelectMany(sheet => sheet.Cells), cell => cell.Value as string == text);
        if (file.EndsWith(".wb1")) {
            Assert.Equal(3, result.Value.Sheets.Count);
            Assert.Equal("B1", Assert.Single(result.Cells, cell => cell.SheetName == "page1" && cell.Row == 1 && cell.Column == 1).Formula);
        }
        if (file.EndsWith(".wks")) Assert.Equal(2d, Assert.Single(result.Cells, cell => cell.Row == 6 && cell.Column == 2).CachedValue);
    }

    [Fact]
    public void LotusReferencesKeepNamesCoordinatesAndOperatorMeaning() {
        byte[] header = new byte[26]; header[0] = 5; header[1] = 0x10;
        byte[] cell = Join(new byte[] { 0, 0, 0, 1 }, BitConverter.GetBytes(5d), new byte[] { 1, 3, 0, 0, 1, 0, 5 }, BitConverter.GetBytes(128u), new byte[] { 0x0f, 3 });
        byte[] source = Join(Record(0, header), Record(0x23, Join(new byte[] { 0xb0, 0x36, 1, 0 }, Encoding.ASCII.GetBytes("Other\0"))), Record(0x28, cell), Record(1));
        using var result = Import(source);
        Assert.Equal("('Other'!A1+2)", Assert.Single(result.Cells).Formula);
        Assert.Equal(5d, Assert.Single(result.Cells).CachedValue);
        Assert.Equal(2, result.Value.Sheets.Count);
        Assert.DoesNotContain(result.Report.Findings, finding => finding.Code == "LATER_FORMULA_CACHED_FALLBACK");
    }

    [Fact]
    public void WorksWindowsTranslatesFormulaAndPreservesNumericCellAddresses() {
        byte[] tokens = { 1, 0, 0, 0, 0, 5, 2, 0, 9, 3 };
        byte[] source = Join(Record(0xff, new byte[] { 4, 4 }), Record(0x545b, Join(new byte[] { 0, 0, 0, 0, 0, 0 }, BitConverter.GetBytes(3f))),
            Record(0x10, Join(new byte[] { 1, 0, 0, 0, 0, 0 }, BitConverter.GetBytes(5d), BitConverter.GetBytes((ushort)tokens.Length), tokens)), Record(1));
        using var result = Import(source);
        Assert.Equal("($A$1+2)", Assert.Single(result.Cells, cell => cell.Column == 2).Formula);
        Assert.Equal(3d, Assert.Single(result.Cells, cell => cell.Column == 1).CachedValue);
    }

    [Theory]
    [InlineData(0x1000)]
    [InlineData(0x1002)]
    public void LotusExtendedGenerationsDecodeNumbersAndFormulaConstants(int version) {
        byte[] header = new byte[26]; BitConverter.GetBytes((ushort)version).CopyTo(header, 0);
        byte[] source = Join(Record(0, header),
            Record(0x17, Join(new byte[] { 0, 0, 0, 0 }, Extended80(0xc000, 0xa000000000000000))),
            Record(0x19, Join(new byte[] { 1, 0, 0, 0 }, Extended80(0x4001, 0xa000000000000000),
                new byte[] { 0 }, Extended80(0x4000, 0xc000000000000000), new byte[] { 5, 4, 0, 0x0f, 3 })),
            Record(0x18, new byte[] { 2, 0, 0, 0, 37, 0 }), Record(1));
        using var result = Import(source);
        Assert.Equal(-2.5d, Assert.Single(result.Cells, cell => cell.Row == 1).CachedValue);
        Assert.Equal("(3+2)", Assert.Single(result.Cells, cell => cell.Row == 2).Formula);
        Assert.Equal(5d, Assert.Single(result.Cells, cell => cell.Row == 2).CachedValue);
        Assert.Equal(.1d, Assert.Single(result.Cells, cell => cell.Row == 3).CachedValue);
    }

    [Theory]
    [InlineData(1, false, "'Other'!$A$1:$B$2")]
    [InlineData(2, false, null)]
    [InlineData(1, true, null)]
    public void QuattroRangesUseOneSheetQualifierAndRequireConsumedReferences(int lastSheet, bool extra, string? expected) {
        byte[] references = Join(new byte[] { 0, 0x10, 0, 1, 0, 0, 1, (byte)lastSheet, 1, 0 },
            extra ? new byte[] { 0, 0, 0, 1, 0, 0 } : Array.Empty<byte>());
        byte[] tokens = { 2, 3 };
        byte[] formula = Join(new byte[6], BitConverter.GetBytes(42d), new byte[2],
            BitConverter.GetBytes((ushort)(tokens.Length + references.Length)), BitConverter.GetBytes((ushort)tokens.Length), tokens, references);
        using var result = Import(Join(Record(0, new byte[] { 1, 0x10 }), Record(0xca, 1),
            Record(0xcc, Encoding.ASCII.GetBytes("Other\0")), Record(0x10, formula), Record(1)));
        Assert.Equal(expected, Assert.Single(result.Cells).Formula);
        Assert.Equal(42d, Assert.Single(result.Cells).CachedValue);
        Assert.Equal(expected == null, result.Report.Findings.Any(finding => finding.Code == "LATER_FORMULA_CACHED_FALLBACK"));
    }

    [Fact]
    public void CommentsAndContinuedLabelsKeepTheirCellWithoutOverwritingDuplicateValues() {
        byte[] header = new byte[26]; header[0] = 5; header[1] = 0x10;
        byte[] number = Record(0x27, Join(new byte[4], BitConverter.GetBytes(7d)));
        byte[] comment = Record(0x26, Join(new byte[4], Encoding.ASCII.GetBytes("'Note\0")));
        using (var result = Import(Join(Record(0, header), comment, number, Record(1)))) {
            Assert.Equal(7d, Assert.Single(result.Cells).CachedValue);
            Assert.Equal("Note", Assert.Single(result.Cells).Comment);
        }
        Assert.Throws<InvalidDataException>(() => Import(Join(Record(0, header), number, comment, number, Record(1))));
        byte[] bof = Record(0xff, new byte[] { 4, 4 });
        byte[] label = Record(0xf, Join(new byte[6], Encoding.ASCII.GetBytes("Hello\0")));
        byte[] continuation = Record(0x36, Join(new byte[6], Encoding.ASCII.GetBytes(" world\0")));
        using (var result = Import(Join(bof, label, continuation, Record(1)))) Assert.Equal("Hello world", Assert.Single(result.Cells).CachedValue);
        Assert.Throws<InvalidDataException>(() => Import(Join(bof, continuation, Record(1))));
        byte[] longLabel = Record(0xf, Join(new byte[6], Encoding.ASCII.GetBytes(new string('a', 32767) + "\0")));
        Assert.Throws<InvalidDataException>(() => Import(Join(bof, longLabel, continuation, Record(1))));
    }

    [Fact]
    public void RejectsBrokenRecordsMissingEofDuplicatesAndResourceLimits() {
        byte[] label = Record(0xf, Join(new byte[] { 0, 0, 0, 0, 0, 0 }, Encoding.ASCII.GetBytes("Hello\0")));
        byte[] bof = Record(0xff, new byte[] { 4, 4 });
        Assert.Throws<InvalidDataException>(() => Import(Join(bof, label)));
        Assert.Throws<InvalidDataException>(() => Import(Join(bof, label, label, Record(1))));
        byte[] truncated = Join(bof, label, Record(1)); truncated[8] = 255;
        Assert.Throws<InvalidDataException>(() => Import(truncated));
        Assert.Throws<InvalidDataException>(() => Import(Join(bof, label, Record(1)), new OfficeLegacyImportLimits { MaxTextCharacters = 4 }));
        Assert.Throws<InvalidDataException>(() => Import(Join(bof, label, Record(1)), new OfficeLegacyImportLimits { MaxRecords = 2 }));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => LegacySpreadsheetImporter.Import(Join(bof, label, Record(1)), cancellationToken: cancellation.Token));
    }
    private static LegacySpreadsheetImportResult Import(byte[] source, OfficeLegacyImportLimits? limits = null) =>
        LegacySpreadsheetImporter.Import(source, new LegacySpreadsheetImportOptions { RequireStructured = true, Limits = limits ?? new OfficeLegacyImportLimits() });
    private static byte[] Record(ushort type, params byte[] payload) => Join(BitConverter.GetBytes(type), BitConverter.GetBytes((ushort)payload.Length), payload);
    private static byte[] Extended80(ushort exponent, ulong significand) => Join(BitConverter.GetBytes(significand), BitConverter.GetBytes(exponent));
    private static byte[] Join(params byte[][] chunks) => chunks.SelectMany(chunk => chunk).ToArray();
}
