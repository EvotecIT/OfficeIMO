using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class ExcelLibraryComparisonRunner {
    private static bool _captureValidationWorkbook;
    private static byte[]? _capturedValidationWorkbookBytes;

    private static void ValidateWriteScenarioOutputs(
        string scenario,
        IReadOnlyList<LibraryComparisonCase> cases) {
        if (!IsWorkbookWriteScenario(scenario)) {
            return;
        }

        var snapshots = new List<WorkbookSemanticSnapshot>(cases.Count);
        foreach (var comparisonCase in cases) {
            byte[] bytes;
            int outputMetric;
            _captureValidationWorkbook = true;
            _capturedValidationWorkbookBytes = null;
            try {
                outputMetric = comparisonCase.Action();
                bytes = _capturedValidationWorkbookBytes
                    ?? throw new InvalidOperationException("The write action did not expose the generated workbook bytes.");
            } catch (Exception exception) {
                throw new InvalidOperationException(
                    $"{scenario} / {comparisonCase.Library} failed its untimed workbook preflight.",
                    exception);
            } finally {
                _captureValidationWorkbook = false;
                _capturedValidationWorkbookBytes = null;
            }

            if (outputMetric != bytes.Length) {
                throw new InvalidOperationException(
                    $"{scenario} / {comparisonCase.Library} reported {outputMetric.ToString(CultureInfo.InvariantCulture)} bytes but generated {bytes.Length.ToString(CultureInfo.InvariantCulture)} bytes.");
            }

            snapshots.Add(CreateSemanticSnapshot(scenario, comparisonCase.Library, bytes));
        }

        var expected = snapshots[0];
        foreach (var actual in snapshots.Skip(1)) {
            if (expected.Cells.SequenceEqual(actual.Cells, StringComparer.Ordinal)) {
                continue;
            }

            int mismatchIndex = FindFirstMismatch(expected.Cells, actual.Cells);
            string expectedValue = mismatchIndex < expected.Cells.Count ? expected.Cells[mismatchIndex] : "<missing>";
            string actualValue = mismatchIndex < actual.Cells.Count ? actual.Cells[mismatchIndex] : "<missing>";
            throw new InvalidOperationException(
                $"Scenario '{scenario}' produced different workbook data in {expected.Library} and {actual.Library}. "
                + $"First mismatch at semantic item {mismatchIndex.ToString(CultureInfo.InvariantCulture)}: "
                + $"{expected.Library}={expectedValue}; {actual.Library}={actualValue}.");
        }
    }

    private static bool IsWorkbookWriteScenario(string scenario)
        => scenario.StartsWith("write-", StringComparison.Ordinal)
           || scenario.StartsWith("report-", StringComparison.Ordinal)
           || scenario.StartsWith("realworld-", StringComparison.Ordinal)
           || string.Equals(scenario, "append-plain-rows", StringComparison.Ordinal)
           || string.Equals(scenario, "copy-worksheet-package", StringComparison.Ordinal)
           || string.Equals(scenario, "autofit-existing", StringComparison.Ordinal)
           || string.Equals(scenario, "write-text-heavy-default", StringComparison.Ordinal);

    private static WorkbookSemanticSnapshot CreateSemanticSnapshot(string scenario, string library, byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        var validationErrors = new OpenXmlValidator().Validate(document).Take(5).ToArray();
        if (validationErrors.Length > 0) {
            string details = string.Join(
                " | ",
                validationErrors.Select(error => $"{error.Path?.XPath}: {error.Description}"));
            throw new InvalidOperationException(
                $"{scenario} / {library} generated an invalid Open XML workbook: {details}");
        }

        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException($"{scenario} / {library} generated a workbook without a workbook part.");
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
            .Elements<SharedStringItem>()
            .Select(static item => item.InnerText)
            .ToArray() ?? [];
        var cells = new List<string>();
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException($"{scenario} / {library} generated a workbook part without workbook XML.");
        ValidateFeatureContract(scenario, library, CreateFeatureSnapshot(workbookPart));
        var sheets = workbook.Sheets?.Elements<Sheet>().ToArray() ?? [];
        foreach (var sheet in sheets) {
            if (sheet.State?.Value is { } state && state != SheetStateValues.Visible) {
                continue;
            }

            string sheetName = sheet.Name?.Value ?? string.Empty;
            if (IsFeatureWorkbookScenario(scenario)
                && !string.Equals(sheetName, "Data", StringComparison.Ordinal)) {
                continue;
            }

            cells.Add("SHEET:" + sheetName);
            if (sheet.Id?.Value is not string relationshipId) {
                throw new InvalidOperationException($"{scenario} / {library} generated sheet '{sheetName}' without a relationship id.");
            }

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
            var worksheet = worksheetPart.Worksheet
                ?? throw new InvalidOperationException($"{scenario} / {library} generated sheet '{sheetName}' without worksheet XML.");
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                continue;
            }

            uint inferredRowIndex = 0;
            foreach (var row in sheetData.Elements<Row>()) {
                uint rowIndex = row.RowIndex?.Value ?? inferredRowIndex + 1;
                inferredRowIndex = rowIndex;
                int inferredColumnIndex = 0;
                foreach (var cell in row.Elements<Cell>()) {
                    string coordinate;
                    if (!string.IsNullOrWhiteSpace(cell.CellReference?.Value)) {
                        coordinate = cell.CellReference!.Value!;
                        inferredColumnIndex = ParseColumnIndex(coordinate);
                    } else {
                        inferredColumnIndex++;
                        coordinate = GetColumnName(inferredColumnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture);
                    }

                    string? formula = cell.CellFormula?.Text;
                    string? value = GetSemanticCellValue(cell, sharedStrings);
                    if (formula == null && value == "TEXT:" && IsSparseNullScenario(scenario)) {
                        continue;
                    }

                    if (formula == null && value == null) {
                        continue;
                    }

                    cells.Add(sheetName + "!" + coordinate + "="
                        + (formula == null ? string.Empty : "FORMULA:" + formula)
                        + (value == null ? string.Empty : "|VALUE:" + value));
                }
            }
        }

        return new WorkbookSemanticSnapshot(library, cells);
    }

    private static WorkbookFeatureSnapshot CreateFeatureSnapshot(WorkbookPart workbookPart) {
        int tables = 0;
        int autoFilters = 0;
        int frozenPanes = 0;
        int conditionalFormats = 0;
        int dataValidations = 0;
        int charts = 0;
        int pivotTables = 0;
        int customWidthColumns = 0;

        foreach (var worksheetPart in workbookPart.WorksheetParts) {
            if (worksheetPart.Worksheet is not { } worksheet) {
                continue;
            }

            tables += worksheetPart.TableDefinitionParts.Count();
            autoFilters += worksheet.Descendants<AutoFilter>().Count();
            autoFilters += worksheetPart.TableDefinitionParts.Count(part => part.Table?.AutoFilter != null);
            frozenPanes += worksheet.Descendants<Pane>().Count(static pane =>
                pane.State?.Value == PaneStateValues.Frozen
                || pane.State?.Value == PaneStateValues.FrozenSplit);
            conditionalFormats += worksheet.Descendants<ConditionalFormatting>().Count();
            dataValidations += worksheet.Descendants<DataValidation>().Count();
            charts += worksheetPart.DrawingsPart?.ChartParts.Count() ?? 0;
            pivotTables += worksheetPart.PivotTableParts.Count();
            customWidthColumns += worksheet.Descendants<Column>().Count(static column => column.Width?.HasValue == true);
        }

        return new WorkbookFeatureSnapshot(
            tables,
            autoFilters,
            frozenPanes,
            conditionalFormats,
            dataValidations,
            charts,
            pivotTables,
            customWidthColumns);
    }

    private static void ValidateFeatureContract(
        string scenario,
        string library,
        WorkbookFeatureSnapshot features) {
        WorkbookFeatureRequirement requirement = scenario switch {
            ReportWorkbookScenario or ReportWorkbookDataTableScenario
                => new(Tables: true, AutoFilters: true, FrozenPanes: true, ConditionalFormats: true,
                    DataValidations: true, Charts: true, PivotTables: true, CustomWidthColumns: true),
            ReportWorkbookCoreScenario or ReportWorkbookDataTableCoreScenario
                => new(Tables: true, AutoFilters: true, FrozenPanes: true, ConditionalFormats: true,
                    DataValidations: true, CustomWidthColumns: true),
            RealWorldReportScenario or RealWorldReportChartFirstScenario or RealWorldReportShuffledColumnsScenario
                or RealWorldReportExtraColumnScenario or RealWorldReportPostMutationScenario
                => new(Tables: true, AutoFilters: true, FrozenPanes: true, ConditionalFormats: true,
                    DataValidations: true, Charts: true, PivotTables: true, CustomWidthColumns: true),
            RealWorldReportNoAutoFitScenario
                => new(Tables: true, AutoFilters: true, FrozenPanes: true, ConditionalFormats: true,
                    DataValidations: true, Charts: true, PivotTables: true),
            RealWorldReportCoreScenario
                => new(Tables: true, AutoFilters: true, FrozenPanes: true, ConditionalFormats: true,
                    DataValidations: true, CustomWidthColumns: true),
            RealWorldFreezePanesScenario => new(FrozenPanes: true),
            RealWorldAutoFilterScenario => new(AutoFilters: true),
            RealWorldConditionalFormattingScenario => new(ConditionalFormats: true),
            RealWorldDataValidationScenario => new(DataValidations: true),
            RealWorldChartsScenario => new(Charts: true),
            RealWorldPivotTableScenario => new(PivotTables: true),
            _ => default
        };

        var missing = new List<string>();
        AddMissingFeature(missing, requirement.Tables, features.Tables, "table");
        AddMissingFeature(missing, requirement.AutoFilters, features.AutoFilters, "AutoFilter");
        AddMissingFeature(missing, requirement.FrozenPanes, features.FrozenPanes, "frozen pane");
        AddMissingFeature(missing, requirement.ConditionalFormats, features.ConditionalFormats, "conditional formatting");
        AddMissingFeature(missing, requirement.DataValidations, features.DataValidations, "data validation");
        AddMissingFeature(missing, requirement.Charts, features.Charts, "chart");
        AddMissingFeature(missing, requirement.PivotTables, features.PivotTables, "pivot table");
        AddMissingFeature(missing, requirement.CustomWidthColumns, features.CustomWidthColumns, "custom column width");
        if (missing.Count > 0) {
            throw new InvalidOperationException(
                $"{scenario} / {library} omitted required workbook structures: {string.Join(", ", missing)}.");
        }
    }

    private static void AddMissingFeature(List<string> missing, bool required, int count, string name) {
        if (required && count == 0) {
            missing.Add(name);
        }
    }

    private static bool IsSparseNullScenario(string scenario)
        => scenario.Contains("sparse", StringComparison.Ordinal);

    private static bool IsFeatureWorkbookScenario(string scenario)
        => scenario.StartsWith("report-", StringComparison.Ordinal)
           || scenario.StartsWith("realworld-", StringComparison.Ordinal);

    private static string? GetSemanticCellValue(Cell cell, IReadOnlyList<string> sharedStrings) {
        if (cell.DataType?.Value == CellValues.InlineString) {
            return "TEXT:" + (cell.InlineString?.InnerText ?? string.Empty);
        }

        string? raw = cell.CellValue?.Text;
        if (raw == null) {
            return null;
        }

        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.SharedString) {
            return GetSharedString(raw, sharedStrings);
        }

        if (dataType == CellValues.Boolean) {
            return raw == "1" ? "BOOLEAN:true" : "BOOLEAN:false";
        }

        if (dataType == CellValues.String) {
            return "TEXT:" + raw;
        }

        if (dataType == CellValues.Error) {
            return "ERROR:" + raw;
        }

        return "NUMBER:" + NormalizeNumber(raw);
    }

    private static string GetSharedString(string raw, IReadOnlyList<string> sharedStrings) {
        if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
            || index < 0
            || index >= sharedStrings.Count) {
            throw new InvalidOperationException($"Workbook contains invalid shared-string index '{raw}'.");
        }

        return "TEXT:" + sharedStrings[index];
    }

    private static string NormalizeNumber(string raw)
        => decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal value)
            ? decimal.Round(value, 8, MidpointRounding.AwayFromZero).ToString("G29", CultureInfo.InvariantCulture)
            : raw;

    private static int ParseColumnIndex(string cellReference) {
        int index = 0;
        foreach (char character in cellReference) {
            if (!char.IsLetter(character)) {
                break;
            }

            index = checked((index * 26) + (char.ToUpperInvariant(character) - 'A' + 1));
        }

        return index;
    }

    private static string GetColumnName(int columnIndex) {
        Span<char> buffer = stackalloc char[8];
        int position = buffer.Length;
        int current = columnIndex;
        do {
            current--;
            buffer[--position] = (char)('A' + (current % 26));
            current /= 26;
        } while (current > 0);

        return new string(buffer[position..]);
    }

    private static int FindFirstMismatch(IReadOnlyList<string> expected, IReadOnlyList<string> actual) {
        int count = Math.Min(expected.Count, actual.Count);
        for (int i = 0; i < count; i++) {
            if (!string.Equals(expected[i], actual[i], StringComparison.Ordinal)) {
                return i;
            }
        }

        return count;
    }

    private sealed record WorkbookSemanticSnapshot(string Library, IReadOnlyList<string> Cells);

    private readonly record struct WorkbookFeatureSnapshot(
        int Tables,
        int AutoFilters,
        int FrozenPanes,
        int ConditionalFormats,
        int DataValidations,
        int Charts,
        int PivotTables,
        int CustomWidthColumns);

    private readonly record struct WorkbookFeatureRequirement(
        bool Tables = false,
        bool AutoFilters = false,
        bool FrozenPanes = false,
        bool ConditionalFormats = false,
        bool DataValidations = false,
        bool Charts = false,
        bool PivotTables = false,
        bool CustomWidthColumns = false);
}
