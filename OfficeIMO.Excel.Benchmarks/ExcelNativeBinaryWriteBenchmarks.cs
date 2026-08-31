using BenchmarkDotNet.Attributes;
using System.Diagnostics;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;
using OfficeIMO.Excel.LegacyXls.Model;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Provides a diagnostic scenario for OfficeIMO's native binary workbook
/// formats and ExcelReader.NET. Structurally invalid competitor output remains
/// visible as a conformance probe, but is not timed or ranked as equivalent.
/// </summary>
internal sealed class ExcelNativeBinaryWriteBenchmarks {
    private BinaryWriteRow[] _rows = null!;
    private int _validatedOfficeOutputBytes;

    public int RowCount { get; set; }

    // The conformance observation is deliberately evaluated at runtime. A new
    // competitor release can move from diagnostic to equivalent without hiding
    // the structural risk from the comparison runner.
    public ExcelFileFormat Format { get; set; }

    internal BinaryWriteConformanceObservation SetupComparison() {
        SetupOfficeIMOOnly();
        byte[] excelReaderWorkbook = WriteExcelReaderWorkbook();
        return InspectExcelReaderWorkbook(excelReaderWorkbook) with {
            OfficeOutputBytes = _validatedOfficeOutputBytes,
            CompetitorOutputBytes = excelReaderWorkbook.Length
        };
    }

    internal void SetupOfficeIMOOnly() {
        _rows = new BinaryWriteRow[RowCount];
        for (int index = 0; index < _rows.Length; index++) {
            _rows[index] = new BinaryWriteRow(
                index + 1,
                "Region " + (index % 8),
                "Owner " + (index % 32),
                Math.Round(100d + ((index * 17.25d) % 9_000d), 2),
                (index & 1) == 0);
        }

        byte[] workbook = WriteWorkbook();
        Validate(workbook);
        if (Format == ExcelFileFormat.Xls) {
            (int actualDbCellBlocks, int expectedDbCellBlocks) = InspectLegacyXlsDbCellBlocks(workbook);
            if (actualDbCellBlocks != expectedDbCellBlocks) {
                throw new InvalidDataException(
                    $"OfficeIMO XLS emitted {actualDbCellBlocks} BIFF8 DBCell blocks instead of {expectedDbCellBlocks}.");
            }
        }
        _validatedOfficeOutputBytes = workbook.Length;
    }

    public int OfficeIMO_PublicTabularWrite() => WriteWorkbook().Length;

    public int ExcelReaderNet_DiagnosticWrite() => WriteExcelReaderWorkbook().Length;

    internal IReadOnlyList<(string Name, double Milliseconds)> ProfileOfficeIMOWriteStages() {
        var stages = new List<(string Name, double Milliseconds)>();
        _ = WriteWorkbook((name, elapsed) => stages.Add((name, elapsed.TotalMilliseconds)));
        return stages;
    }

    private byte[] WriteWorkbook(Action<string, TimeSpan>? reportStage = null) {
        Stopwatch? stageWatch = reportStage is null ? null : Stopwatch.StartNew();
        using ExcelDocument document = ExcelDocument.Create();
        ReportProfileStage(reportStage, stageWatch, "CreateDocument");
        if (reportStage is not null) {
            document.Execution.OnTiming = reportStage;
        }
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.InsertObjects(
            _rows,
            ("Id", static row => row.Id),
            ("Region", static row => row.Region),
            ("Owner", static row => row.Owner),
            ("Amount", static row => row.Amount),
            ("Active", static row => row.Active));
        ReportProfileStage(reportStage, stageWatch, "AddSheetAndInsertObjects");

        byte[] workbook = document.ToBytes(Format);
        ReportProfileStage(reportStage, stageWatch, "ToBytes");
        if (document.LastSaveDiagnostics.Writer != ExcelSavePackageWriter.NativeBinaryDirectPackage) {
            throw new InvalidOperationException(
                $"OfficeIMO {Format} benchmark did not use the native direct tabular writer: "
                + document.LastSaveDiagnostics.FastPackageSkipReason);
        }

        return workbook;
    }

    private static void ReportProfileStage(
        Action<string, TimeSpan>? reportStage,
        Stopwatch? stopwatch,
        string name) {
        if (reportStage is null || stopwatch is null) return;
        reportStage(name, stopwatch.Elapsed);
        stopwatch.Restart();
    }

    private byte[] WriteExcelReaderWorkbook() => Format switch {
        ExcelFileFormat.Xls => WriteExcelReaderXlsWorkbook(),
        ExcelFileFormat.Xlsb => WriteExcelReaderXlsbWorkbook(),
        _ => throw new InvalidOperationException($"ExcelReader.NET binary benchmark does not support {Format}.")
    };

    private byte[] WriteExcelReaderXlsWorkbook() {
        using var stream = new MemoryStream();
        using XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(stream, leaveOpen: true);
        workbook.Start();
        using (XlsSheetWriter sheet = workbook.AddSheet("Data")) {
            sheet.Start();
            using (XlsRowWriter header = sheet.StartRow()) {
                WriteHeaders(header);
            }

            foreach (BinaryWriteRow value in _rows) {
                using XlsRowWriter row = sheet.StartRow();
                WriteRow(row, value);
            }
            sheet.End();
        }
        workbook.End();
        return stream.ToArray();
    }

    private byte[] WriteExcelReaderXlsbWorkbook() {
        using var stream = new MemoryStream();
        using XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(stream, leaveOpen: true);
        workbook.Start();
        using (XlsbSheetWriter sheet = workbook.AddSheet("Data")) {
            sheet.Start();
            using (XlsbRowWriter header = sheet.StartRow()) {
                WriteHeaders(header);
            }

            foreach (BinaryWriteRow value in _rows) {
                using XlsbRowWriter row = sheet.StartRow();
                WriteRow(row, value);
            }
            sheet.End();
        }
        workbook.End();
        return stream.ToArray();
    }

    private static void WriteHeaders<TRow>(TRow row)
        where TRow : IRowWriter {
        row.Write("Id");
        row.Write("Region");
        row.Write("Owner");
        row.Write("Amount");
        row.Write("Active");
    }

    private static void WriteRow<TRow>(TRow row, BinaryWriteRow value)
        where TRow : IRowWriter {
        row.Write(value.Id);
        row.Write(value.Region);
        row.Write(value.Owner);
        row.Write(value.Amount);
        row.Write(value.Active);
    }

    private void Validate(byte[] workbook) {
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
        string[] headers = ["Id", "Region", "Owner", "Amount", "Active"];
        if (reader.FieldCount != headers.Length) {
            throw new InvalidDataException($"{Format} exposed {reader.FieldCount} fields instead of {headers.Length}.");
        }

        for (int column = 0; column < headers.Length; column++) {
            if (!string.Equals(reader.GetName(column), headers[column], StringComparison.Ordinal)) {
                throw new InvalidDataException($"{Format} header {column + 1} did not round-trip.");
            }
        }

        int rowIndex = 0;
        while (reader.Read()) {
            if (rowIndex >= _rows.Length) {
                throw new InvalidDataException($"{Format} emitted extra rows.");
            }

            BinaryWriteRow expected = _rows[rowIndex];
            if (reader.GetInt32(0) != expected.Id
                || !string.Equals(reader.GetString(1), expected.Region, StringComparison.Ordinal)
                || !string.Equals(reader.GetString(2), expected.Owner, StringComparison.Ordinal)
                || reader.GetDouble(3) != expected.Amount
                || reader.GetBoolean(4) != expected.Active) {
                throw new InvalidDataException($"{Format} row {rowIndex + 2} did not round-trip.");
            }

            rowIndex++;
        }

        if (rowIndex != _rows.Length || reader.NextResult()) {
            throw new InvalidDataException($"{Format} did not round-trip the exact single-sheet row set.");
        }
    }

    private BinaryWriteConformanceObservation InspectExcelReaderWorkbook(byte[] workbook) {
        bool semanticRoundTrip = TryValidateExcelReaderWorkbook(workbook, out string? semanticFailure);
        bool structurallyConformant;
        string structureDetail;

        if (Format == ExcelFileFormat.Xls) {
            try {
                (int actualDbCellBlocks, int expectedDbCellBlocks) = InspectLegacyXlsDbCellBlocks(workbook);
                structurallyConformant = actualDbCellBlocks == expectedDbCellBlocks;
                structureDetail = structurallyConformant
                    ? $"BIFF8 Index/DBCell structure present ({actualDbCellBlocks} blocks)."
                    : $"BIFF8 Index/DBCell structure missing or incomplete ({actualDbCellBlocks} of {expectedDbCellBlocks} blocks).";
            }
            catch (Exception exception) {
                structurallyConformant = false;
                structureDetail = "BIFF8 structural inspection failed: " + exception.Message;
            }
        } else {
            try {
                Validate(workbook);
                structurallyConformant = true;
                structureDetail = "Strict OfficeIMO XLSB validation passed, including BrtWsDim.";
            }
            catch (Exception exception) {
                structurallyConformant = false;
                structureDetail = "Strict OfficeIMO XLSB validation failed: " + exception.Message;
            }
        }

        string semanticDetail = semanticRoundTrip
            ? "ExcelReader.NET round-tripped the exact single-sheet row set with its own reader."
            : "ExcelReader.NET semantic round-trip failed: " + semanticFailure;
        return new BinaryWriteConformanceObservation(
            semanticRoundTrip,
            structurallyConformant,
            semanticDetail + " " + structureDetail);
    }

    private (int Actual, int Expected) InspectLegacyXlsDbCellBlocks(byte[] workbook) {
        LegacyXlsWorkbook parsed = LegacyXlsWorkbook.Load(workbook);
        LegacyXlsWorksheet? sheet = parsed.Worksheets.FirstOrDefault();
        return (sheet?.RowBlockIndex?.DbCellBlockCount ?? 0, (RowCount + 1 + 31) / 32);
    }

    private bool TryValidateExcelReaderWorkbook(byte[] workbook, out string? failure) {
        try {
            if (Format == ExcelFileFormat.Xls) {
                using XlsReader reader = ExcelReaderApi.FromXls(workbook);
                ValidateExcelReaderRows(reader);
            } else {
                using XlsbReader reader = ExcelReaderApi.FromXlsb(workbook);
                ValidateExcelReaderRows(reader);
            }

            failure = null;
            return true;
        }
        catch (Exception exception) {
            failure = exception.Message;
            return false;
        }
    }

    private void ValidateExcelReaderRows(XlsReader reader) {
        int rowIndex = -1;
        foreach (Row row in reader) {
            ValidateExcelReaderRow(row, ref rowIndex);
        }
        ValidateExcelReaderRowCount(rowIndex);
    }

    private void ValidateExcelReaderRows(XlsbReader reader) {
        int rowIndex = -1;
        foreach (Row row in reader) {
            ValidateExcelReaderRow(row, ref rowIndex);
        }
        ValidateExcelReaderRowCount(rowIndex);
    }

    private void ValidateExcelReaderRow(Row row, ref int rowIndex) {
        string[] headers = ["Id", "Region", "Owner", "Amount", "Active"];
        if (rowIndex < 0) {
            if (row.ColumnCount < headers.Length) {
                throw new InvalidDataException(
                    $"{Format} exposed {row.ColumnCount} header fields instead of {headers.Length}.");
            }
            for (int column = 0; column < headers.Length; column++) {
                if (!string.Equals(row[column].GetString(), headers[column], StringComparison.Ordinal)) {
                    throw new InvalidDataException($"{Format} header {column + 1} did not round-trip.");
                }
            }
            rowIndex = 0;
            return;
        }

        if (rowIndex >= _rows.Length) {
            throw new InvalidDataException($"{Format} emitted extra rows.");
        }

        BinaryWriteRow expected = _rows[rowIndex];
        if (!row[0].TryParse(System.Globalization.CultureInfo.InvariantCulture, out int id)
            || id != expected.Id
            || !string.Equals(row[1].GetString(), expected.Region, StringComparison.Ordinal)
            || !string.Equals(row[2].GetString(), expected.Owner, StringComparison.Ordinal)
            || !row[3].TryGetDouble(out double amount)
            || amount != expected.Amount
            || !TryReadExcelReaderBoolean(row[4], out bool active)
            || active != expected.Active) {
            throw new InvalidDataException($"{Format} row {rowIndex + 2} did not round-trip.");
        }

        rowIndex++;
    }

    private static bool TryReadExcelReaderBoolean(Cell cell, out bool value) {
        if (bool.TryParse(cell.GetString(), out value)) {
            return true;
        }
        if (cell.TryGetDouble(out double numeric) && (numeric == 0d || numeric == 1d)) {
            value = numeric == 1d;
            return true;
        }

        value = default;
        return false;
    }

    private void ValidateExcelReaderRowCount(int rowIndex) {
        if (rowIndex != _rows.Length) {
            throw new InvalidDataException($"{Format} emitted {Math.Max(rowIndex, 0)} rows instead of {_rows.Length}.");
        }
    }

    private readonly record struct BinaryWriteRow(int Id, string Region, string Owner, double Amount, bool Active);
}

internal readonly record struct BinaryWriteConformanceObservation(
    bool SemanticRoundTrip,
    bool StructurallyConformant,
    string Detail,
    int OfficeOutputBytes = 0,
    int CompetitorOutputBytes = 0) {
    internal bool IsEquivalent => SemanticRoundTrip && StructurallyConformant;
}

/// <summary>
/// Measures OfficeIMO's validated native binary write paths without ranking
/// structurally invalid competitor output.
/// </summary>
[MemoryDiagnoser]
public class OfficeNativeBinaryWriteBenchmarks {
    private ExcelNativeBinaryWriteBenchmarks _scenario = null!;

    [Params(2_500, 25_000)]
    public int RowCount { get; set; }

    [Params(ExcelFileFormat.Xls, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = new ExcelNativeBinaryWriteBenchmarks {
            RowCount = RowCount,
            Format = Format
        };
        _scenario.SetupOfficeIMOOnly();
    }

    [Benchmark]
    public int OfficeIMO_ValidatedNativeBinaryWrite() =>
        _scenario.OfficeIMO_PublicTabularWrite();
}
