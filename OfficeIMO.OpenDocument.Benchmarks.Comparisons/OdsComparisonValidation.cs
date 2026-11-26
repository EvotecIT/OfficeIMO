namespace OfficeIMO.OpenDocument.Benchmarks.Comparisons;

internal static class OdsComparisonValidation {
    internal static IReadOnlyList<OdsComparisonReport> ValidateAll() =>
        OdsComparisonCorpus.Scales.Select(scale => Validate(scale.Name)).ToArray();

    internal static OdsComparisonReport Validate(string scaleName) {
        OdsComparisonScale scale = OdsComparisonCorpus.Get(scaleName);
        byte[] officeIMOBytes = OdsComparisonWorkflows.CreateOfficeIMO(scale);
        byte[] openStandardLibraryBytes = OdsComparisonWorkflows.CreateOpenStandardLibrary(scale).GetAwaiter().GetResult();
        OdsOutputEvidence officeIMO = Inspect("OfficeIMO", scale, officeIMOBytes);
        OdsOutputEvidence openStandardLibrary = Inspect("OpenStandardLibrary", scale, openStandardLibraryBytes);

        if (officeIMO.RecordCount != openStandardLibrary.RecordCount ||
            officeIMO.ContentLength != openStandardLibrary.ContentLength) {
            throw new InvalidOperationException(
                $"ODS/{scale.Name} semantic output differs: " +
                $"OfficeIMO {officeIMO.RecordCount} cells/{officeIMO.ContentLength} characters; " +
                $"OpenStandardLibrary {openStandardLibrary.RecordCount} cells/{openStandardLibrary.ContentLength} characters.");
        }

        return new OdsComparisonReport(scale.Name, officeIMO, openStandardLibrary);
    }

    // Reads the package with OfficeIMO and compares every populated cell, by position, with the corpus value.
    internal static OdsOutputEvidence Inspect(string implementation, OdsComparisonScale scale, byte[] package) {
        OdsDocument document = OdsDocument.Load(new MemoryStream(package, writable: false));
        OdsSheet sheet = document.Sheets.Single();
        long rowIndex = 0;
        long rowCount = 0;
        long cellCount = 0;
        long contentLength = 0;
        foreach (OdsRowRun row in sheet.RowRuns) {
            if (!row.CellRuns.Any(cell => cell.Value.ToString().Length > 0)) {
                rowIndex += row.RepeatCount;
                continue;
            }
            for (long repeat = 0; repeat < row.RepeatCount; repeat++, rowIndex++) {
                rowCount++;
                long columnIndex = 0;
                foreach (OdsCellRun cell in row.CellRuns) {
                    string value = cell.Value.ToString();
                    if (value.Length == 0) {
                        columnIndex += cell.RepeatCount;
                        continue;
                    }
                    for (long cellRepeat = 0; cellRepeat < cell.RepeatCount; cellRepeat++, columnIndex++) {
                        RequireCell(implementation, scale, rowIndex, columnIndex, value);
                        cellCount++;
                        contentLength += value.Length;
                    }
                }
            }
        }

        RequireCount(implementation, "rows", scale.Rows, rowCount);
        RequireCount(implementation, "cells", checked((long)scale.Rows * scale.Columns), cellCount);
        return new OdsOutputEvidence(implementation, package.LongLength, cellCount, contentLength);
    }

    // Compares every cell OpenStandardLibrary read from the package, by position, with the corpus value.
    internal static long InspectOpenStandardLibrary(OdsComparisonScale scale, byte[] package) {
        return InspectOpenStandardLibraryAsync(scale, package).GetAwaiter().GetResult();
    }

    private static async Task<long> InspectOpenStandardLibraryAsync(OdsComparisonScale scale, byte[] package) {
        await using var spreadsheet = new global::OoxSpreadsheet.Spreadsheet();
        OslSpreadsheet.Models.oWorkbook workbook = await spreadsheet.ImportOdsFileAsync(package).ConfigureAwait(false);
        long cellCount = 0;
        foreach (OslSpreadsheet.Models.oCell cell in workbook.Sheets.Single().Cells) {
            RequireCell("OpenStandardLibrary", scale, cell.Row - 1, cell.Column - 1, cell.Value);
            cellCount++;
        }
        RequireCount("OpenStandardLibrary", "cells", checked((long)scale.Rows * scale.Columns), cellCount);
        return cellCount;
    }

    private static void RequireCell(string implementation, OdsComparisonScale scale, long row, long column, string actual) {
        if (row < 0 || row >= scale.Rows || column < 0 || column >= scale.Columns) {
            throw new InvalidOperationException(
                $"{implementation} returned '{actual}' at row {row}, column {column}, outside the {scale.Rows}x{scale.Columns} corpus.");
        }
        string expected = OdsComparisonCorpus.Cell((int)row, (int)column);
        if (!string.Equals(actual, expected, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{implementation} returned '{actual}' at row {row}, column {column}; expected '{expected}'.");
        }
    }

    private static void RequireCount(string implementation, string contract, long expected, long actual) {
        if (actual != expected) {
            throw new InvalidOperationException(
                $"{implementation} produced {actual} {contract}; expected {expected}.");
        }
    }
}

internal sealed record OdsOutputEvidence(
    string Implementation,
    long OutputBytes,
    long RecordCount,
    long ContentLength);

internal sealed record OdsComparisonReport(
    string Scale,
    OdsOutputEvidence OfficeIMO,
    OdsOutputEvidence OpenStandardLibrary);
