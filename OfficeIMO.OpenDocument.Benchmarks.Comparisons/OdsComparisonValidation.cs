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

    private static OdsOutputEvidence Inspect(string implementation, OdsComparisonScale scale, byte[] package) {
        OdsDocument document = OdsDocument.Load(new MemoryStream(package, writable: false));
        OdsSheet sheet = document.Sheets.Single();
        long rowCount = 0;
        long cellCount = 0;
        long contentLength = 0;
        string? first = null;
        string? last = null;
        foreach (OdsRowRun row in sheet.RowRuns) {
            bool populatedRow = false;
            foreach (OdsCellRun cell in row.CellRuns) {
                string value = cell.Value.ToString();
                if (value.Length == 0) continue;
                populatedRow = true;
                long expandedCells = checked(row.RepeatCount * cell.RepeatCount);
                cellCount += expandedCells;
                contentLength += checked(expandedCells * value.Length);
                first ??= value;
                last = value;
            }
            if (populatedRow) rowCount += row.RepeatCount;
        }

        RequireCount(implementation, "rows", scale.Rows, rowCount);
        RequireCount(implementation, "cells", checked((long)scale.Rows * scale.Columns), cellCount);
        RequireMarker(implementation, first, OdsComparisonCorpus.Cell(0, 0));
        RequireMarker(
            implementation,
            last,
            OdsComparisonCorpus.Cell(scale.Rows - 1, scale.Columns - 1));
        return new OdsOutputEvidence(implementation, package.LongLength, cellCount, contentLength);
    }

    private static void RequireCount(string implementation, string contract, long expected, long actual) {
        if (actual != expected) {
            throw new InvalidOperationException(
                $"{implementation} produced {actual} {contract}; expected {expected}.");
        }
    }

    private static void RequireMarker(string implementation, string? actual, string expected) {
        if (!string.Equals(actual, expected, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{implementation} did not preserve required marker '{expected}'. Actual: '{actual}'.");
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
