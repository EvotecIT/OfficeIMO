using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelLibraryComparisonRunner {
    internal const int DefaultWarmupIterations = 1;
    internal const int DefaultMeasuredIterations = 3;

    private const int DefaultRowCount = 2500;
    private const int SparseLastRow = 100_001;
#if DEBUG
    private const string BuildConfiguration = "Debug";
#else
    private const string BuildConfiguration = "Release";
#endif

    internal static string WriteComparison(
        string outputPath,
        int rowCount = DefaultRowCount,
        bool includeLegacyEpPlus = true,
        IReadOnlyCollection<string>? scenarioFilters = null,
        int warmupIterations = DefaultWarmupIterations,
        int measuredIterations = DefaultMeasuredIterations) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }

        if (rowCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        }

        if (warmupIterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(warmupIterations));
        }

        if (measuredIterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(measuredIterations));
        }

        ConfigureEpPlusLicense();
        var scenarioFilter = BuildScenarioFilter(scenarioFilters);

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        byte[] officeImoWorkbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        byte[] closedXmlWorkbookBytes = CreateClosedXmlWorkbookBytes(rows);
        byte[] epPlusWorkbookBytes = CreateEpPlusWorkbookBytes(rows);
        byte[] formulaWorkbookBytes = CreateFormulaWorkbookBytes(rowCount);
        byte[] sharedStringWorkbookBytes = CreateSharedStringWorkbookBytes(rowCount);
        byte[] sparseWorkbookBytes = CreateSparseWorkbookBytes(SparseLastRow);
        string dataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(rowCount);
        int topDataRows = Math.Min(rowCount, 100);
        string topDataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(topDataRows);
        string sparseRange = $"A1:A{SparseLastRow}";
        var scenarios = new List<ExcelLibraryComparisonScenario>();

        AddScenarioGroup(scenarios, scenarioFilter, "write-bulk-report", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert objects, add table, autofit, save.", () => OfficeImoWriteBulkReport(rows)),
            new LibraryComparisonCase("ClosedXML", "Insert table, apply table style, autofit, save.", () => ClosedXmlWriteBulkReport(rows)),
            new LibraryComparisonCase("EPPlus", "Manual row population, add table, autofit, save.", () => EpPlusWriteBulkReport(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "append-plain-rows", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Append prepared plain cells with CellValues parallel mode.", () => OfficeImoAppendPlainRows(rows)),
            new LibraryComparisonCase("ClosedXML", "Append equivalent row/cell values.", () => ClosedXmlAppendPlainRows(rows)),
            new LibraryComparisonCase("EPPlus", "Append equivalent row/cell values.", () => EpPlusAppendPlainRows(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1 range with automatic execution policy.", () => OfficeImoReadRange(officeImoWorkbookBytes, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Iterate used data cells from workbook.", () => ClosedXmlReadRange(closedXmlWorkbookBytes)),
            new LibraryComparisonCase("EPPlus", "Iterate used data cells from workbook.", () => EpPlusReadRange(epPlusWorkbookBytes))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-top-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read the first 100 data rows from a larger sheet.", () => OfficeImoReadRange(officeImoWorkbookBytes, topDataRange)),
            new LibraryComparisonCase("ClosedXML", "Read the first 100 data rows from a larger sheet.", () => ClosedXmlReadRange(closedXmlWorkbookBytes, topDataRows)),
            new LibraryComparisonCase("EPPlus", "Read the first 100 data rows from a larger sheet.", () => EpPlusReadRange(epPlusWorkbookBytes, topDataRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-range-stream", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream A1 range in row chunks with automatic execution policy.", () => OfficeImoReadRangeStream(officeImoWorkbookBytes, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Iterate used data cells row-by-row.", () => ClosedXmlReadRangeStream(closedXmlWorkbookBytes)),
            new LibraryComparisonCase("EPPlus", "Iterate used data cells row-by-row.", () => EpPlusReadRangeStream(epPlusWorkbookBytes))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-top-range-stream", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream the first 100 data rows from a larger sheet.", () => OfficeImoReadRangeStream(officeImoWorkbookBytes, topDataRange)),
            new LibraryComparisonCase("ClosedXML", "Read the first 100 data rows from a larger sheet row-by-row.", () => ClosedXmlReadRangeStream(closedXmlWorkbookBytes, topDataRows)),
            new LibraryComparisonCase("EPPlus", "Read the first 100 data rows from a larger sheet row-by-row.", () => EpPlusReadRangeStream(epPlusWorkbookBytes, topDataRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "large-sparse-column-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1:A100001 with only first and last rows populated.", () => OfficeImoReadSparseColumn(sparseWorkbookBytes, sparseRange, SparseLastRow)),
            new LibraryComparisonCase("ClosedXML", "Read A1:A100001 with only first and last rows populated.", () => ClosedXmlReadSparseColumn(sparseWorkbookBytes, SparseLastRow)),
            new LibraryComparisonCase("EPPlus", "Read A1:A100001 with only first and last rows populated.", () => EpPlusReadSparseColumn(sparseWorkbookBytes, SparseLastRow))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "large-sparse-row-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1:A100001 as rows with only first and last rows populated.", () => OfficeImoReadSparseRows(sparseWorkbookBytes, sparseRange, SparseLastRow)),
            new LibraryComparisonCase("ClosedXML", "Read A1:A100001 as rows with only first and last rows populated.", () => ClosedXmlReadSparseRows(sparseWorkbookBytes, SparseLastRow)),
            new LibraryComparisonCase("EPPlus", "Read A1:A100001 as rows with only first and last rows populated.", () => EpPlusReadSparseRows(sparseWorkbookBytes, SparseLastRow))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-objects", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Typed materialization with ReadObjects<T>.", () => OfficeImoReadObjects(officeImoWorkbookBytes, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Manual typed materialization from worksheet rows.", () => ClosedXmlReadObjects(closedXmlWorkbookBytes)),
            new LibraryComparisonCase("EPPlus", "Manual typed materialization from worksheet rows.", () => EpPlusReadObjects(epPlusWorkbookBytes))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "autofit-existing", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Load existing workbook, autofit columns, save.", () => OfficeImoAutoFitExisting(officeImoWorkbookBytes)),
            new LibraryComparisonCase("ClosedXML", "Load existing workbook, autofit columns, save.", () => ClosedXmlAutoFitExisting(closedXmlWorkbookBytes)),
            new LibraryComparisonCase("EPPlus", "Load existing workbook, autofit columns, save.", () => EpPlusAutoFitExisting(epPlusWorkbookBytes))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "large-shared-strings", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write repeated and distinct text-heavy cells.", () => OfficeImoWriteSharedStrings(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Write repeated and distinct text-heavy cells.", () => ClosedXmlWriteSharedStrings(rowCount)),
            new LibraryComparisonCase("EPPlus", "Write repeated and distinct text-heavy cells.", () => EpPlusWriteSharedStrings(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "formula-heavy-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read formula text with cached formula results disabled.", () => OfficeImoReadFormulaText(formulaWorkbookBytes, rowCount)),
            new LibraryComparisonCase("ClosedXML", "Read formula A1 text from formula cells.", () => ClosedXmlReadFormulaText(formulaWorkbookBytes, rowCount)),
            new LibraryComparisonCase("EPPlus", "Read formula text from formula cells.", () => EpPlusReadFormulaText(formulaWorkbookBytes, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "shared-string-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read repeated shared string payload.", () => OfficeImoReadSharedStrings(sharedStringWorkbookBytes, rowCount)),
            new LibraryComparisonCase("ClosedXML", "Read repeated shared string payload.", () => ClosedXmlReadSharedStrings(sharedStringWorkbookBytes, rowCount)),
            new LibraryComparisonCase("EPPlus", "Read repeated shared string payload.", () => EpPlusReadSharedStrings(sharedStringWorkbookBytes, rowCount))
        ]);

        if (scenarios.Count == 0) {
            throw new ArgumentException("No comparison scenarios matched the requested --scenario filter.");
        }

        var profile = new ExcelLibraryComparisonProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            BuildConfiguration = BuildConfiguration,
            RowCount = rowCount,
            WarmupIterations = warmupIterations,
            MeasuredIterations = measuredIterations,
            Notes = "Local opt-in comparison. Not intended for CI gating.",
            Scenarios = scenarios
        };

        if (includeLegacyEpPlus) {
            profile.Scenarios.AddRange(RunLegacyEpPlusComparison(rowCount, scenarioFilter, warmupIterations, measuredIterations));
        }

        ValidateComparableReadMetrics(profile.Scenarios);

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(outputPath, JsonSerializer.Serialize(profile, options));
        return outputPath;
    }

    private static HashSet<string>? BuildScenarioFilter(IReadOnlyCollection<string>? scenarioFilters) {
        if (scenarioFilters == null || scenarioFilters.Count == 0) {
            return null;
        }

        var filter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string scenario in scenarioFilters) {
            if (!string.IsNullOrWhiteSpace(scenario)) {
                filter.Add(scenario.Trim());
            }
        }

        return filter.Count == 0 ? null : filter;
    }

    private static void AddScenarioGroup(
        List<ExcelLibraryComparisonScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        string scenario,
        int warmupIterations,
        int measuredIterations,
        IReadOnlyList<LibraryComparisonCase> cases) {
        if (scenarioFilter != null && !scenarioFilter.Contains(scenario)) {
            return;
        }

        Console.WriteLine($"Running {scenario} comparison group...");
        var measurements = BenchmarkMeasurement.MeasureGroup(
            warmupIterations,
            measuredIterations,
            cases.Select(c => c.Action).ToArray());

        for (int i = 0; i < cases.Count; i++) {
            var comparisonCase = cases[i];
            var measurement = measurements[i];
            Console.WriteLine(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"{scenario} / {comparisonCase.Library}: avg {measurement.AverageMilliseconds:F2} ms, median {measurement.MedianMilliseconds:F2} ms"));

            scenarios.Add(new ExcelLibraryComparisonScenario {
                Scenario = scenario,
                Library = comparisonCase.Library,
                Notes = comparisonCase.Notes,
                OutputMetric = measurement.OutputMetric,
                AverageMilliseconds = measurement.AverageMilliseconds,
                MedianMilliseconds = measurement.MedianMilliseconds,
                SamplesMilliseconds = measurement.SamplesMilliseconds.ToList()
            });
        }
    }

    private static IReadOnlyList<ExcelLibraryComparisonScenario> RunLegacyEpPlusComparison(
        int rowCount,
        IReadOnlySet<string>? scenarioFilter,
        int warmupIterations,
        int measuredIterations) {
        string repositoryRoot = FindRepositoryRoot();
        string legacyProjectPath = Path.Combine(repositoryRoot, "OfficeIMO.Excel.Benchmarks.LegacyEpPlus", "OfficeIMO.Excel.Benchmarks.LegacyEpPlus.csproj");
        if (!File.Exists(legacyProjectPath)) {
            throw new FileNotFoundException("Legacy EPPlus comparison project was not found.", legacyProjectPath);
        }

        string outputPath = Path.Combine(Path.GetTempPath(), "officeimo.excel.legacy-epplus-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + ".json");
        try {
            using var process = new System.Diagnostics.Process {
                StartInfo = new System.Diagnostics.ProcessStartInfo {
                    FileName = "dotnet",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.StartInfo.ArgumentList.Add("run");
            process.StartInfo.ArgumentList.Add("--configuration");
            process.StartInfo.ArgumentList.Add(BuildConfiguration);
            process.StartInfo.ArgumentList.Add("--framework");
            process.StartInfo.ArgumentList.Add("net8.0");
            process.StartInfo.ArgumentList.Add("--project");
            process.StartInfo.ArgumentList.Add(legacyProjectPath);
            process.StartInfo.ArgumentList.Add("--");
            process.StartInfo.ArgumentList.Add(outputPath);
            process.StartInfo.ArgumentList.Add("--rows");
            process.StartInfo.ArgumentList.Add(rowCount.ToString(CultureInfo.InvariantCulture));
            process.StartInfo.ArgumentList.Add("--warmup");
            process.StartInfo.ArgumentList.Add(warmupIterations.ToString(CultureInfo.InvariantCulture));
            process.StartInfo.ArgumentList.Add("--iterations");
            process.StartInfo.ArgumentList.Add(measuredIterations.ToString(CultureInfo.InvariantCulture));
            if (scenarioFilter != null) {
                foreach (string scenario in scenarioFilter) {
                    process.StartInfo.ArgumentList.Add("--scenario");
                    process.StartInfo.ArgumentList.Add(scenario);
                }
            }

            Console.WriteLine("Running legacy EPPlus comparison in an isolated process...");
            process.Start();
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (!string.IsNullOrWhiteSpace(standardOutput)) {
                Console.Write(standardOutput);
            }

            if (process.ExitCode != 0) {
                if (!string.IsNullOrWhiteSpace(standardError)) {
                    Console.Error.Write(standardError);
                }

                throw new InvalidOperationException($"Legacy EPPlus comparison failed with exit code {process.ExitCode}.");
            }

            if (!File.Exists(outputPath)) {
                throw new FileNotFoundException("Legacy EPPlus comparison did not produce an output file.", outputPath);
            }

            var profile = JsonSerializer.Deserialize<ExcelLibraryComparisonProfile>(File.ReadAllText(outputPath));
            return profile?.Scenarios ?? [];
        } finally {
            try {
                if (File.Exists(outputPath)) {
                    File.Delete(outputPath);
                }
            } catch {
                // Best-effort cleanup for the temporary subprocess JSON.
            }
        }
    }

    private static string FindRepositoryRoot() {
        DirectoryInfo? current = new(AppContext.BaseDirectory);
        while (current != null) {
            if (File.Exists(Path.Combine(current.FullName, "OfficeIMO.sln"))) {
                return current.FullName;
            }

            current = current.Parent;
        }

        current = new DirectoryInfo(Directory.GetCurrentDirectory());
        while (current != null) {
            if (File.Exists(Path.Combine(current.FullName, "OfficeIMO.sln"))) {
                return current.FullName;
            }

            current = current.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    private static void ConfigureEpPlusLicense() {
        if (ExcelPackage.License.LicenseType == EPPlusLicenseType.NonCommercialPersonal
            || ExcelPackage.License.LicenseType == EPPlusLicenseType.NonCommercialOrganization
            || ExcelPackage.License.LicenseType == EPPlusLicenseType.Commercial) {
            return;
        }

        ExcelPackage.License.SetNonCommercialOrganization("OfficeIMO local benchmarks");
    }

    private static int OfficeImoWriteBulkReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.PopulateOfficeImoWorksheet(sheet, rows);
        }

        return checked((int)stream.Length);
    }

    private static int ClosedXmlWriteBulkReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(worksheet, rows);
            workbook.SaveAs(stream);
        }

        return checked((int)stream.Length);
    }

    private static int EpPlusWriteBulkReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusWorksheet(worksheet, rows, includeTable: true, autoFit: true);
            package.Save();
        }

        return checked((int)stream.Length);
    }

    private static int OfficeImoAppendPlainRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValues(new[] {
                (1, 1, (object)"Id"),
                (1, 2, (object)"Region"),
                (1, 3, (object)"Owner"),
                (1, 4, (object)"Amount")
            }, ExecutionMode.Sequential);

            var cells = rows.SelectMany((row, index) => new[] {
                (index + 2, 1, (object)row.Id),
                (index + 2, 2, (object)row.Region),
                (index + 2, 3, (object)row.Owner),
                (index + 2, 4, (object)row.Amount)
            }).ToArray();
            sheet.CellValues(cells, ExecutionMode.Parallel);
        }

        return checked((int)stream.Length);
    }

    private static int ClosedXmlAppendPlainRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteHeaders(worksheet);
            for (int i = 0; i < rows.Count; i++) {
                var row = rows[i];
                int r = i + 2;
                worksheet.Cell(r, 1).Value = row.Id;
                worksheet.Cell(r, 2).Value = row.Region;
                worksheet.Cell(r, 3).Value = row.Owner;
                worksheet.Cell(r, 4).Value = row.Amount;
            }

            workbook.SaveAs(stream);
        }

        return checked((int)stream.Length);
    }

    private static int EpPlusAppendPlainRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteAppendHeaders(worksheet);
            for (int i = 0; i < rows.Count; i++) {
                var row = rows[i];
                int r = i + 2;
                worksheet.Cells[r, 1].Value = row.Id;
                worksheet.Cells[r, 2].Value = row.Region;
                worksheet.Cells[r, 3].Value = row.Owner;
                worksheet.Cells[r, 4].Value = row.Amount;
            }

            package.Save();
        }

        return checked((int)stream.Length);
    }

    private static int OfficeImoReadRange(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        object?[,] values = reader.GetSheet("Data").ReadRange(dataRange);
        int metric = AddSalesHeadersMetric(0);
        for (int row = 1; row < values.GetLength(0); row++) {
            metric = AddSalesRangeMetric(
                metric,
                Convert.ToInt32(values[row, 0], CultureInfo.InvariantCulture),
                Convert.ToString(values[row, 1], CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(values[row, 2], CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(values[row, 3]),
                Convert.ToDouble(values[row, 4], CultureInfo.InvariantCulture),
                Convert.ToInt32(values[row, 5], CultureInfo.InvariantCulture),
                Convert.ToBoolean(values[row, 6], CultureInfo.InvariantCulture),
                Convert.ToString(values[row, 7], CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int ClosedXmlReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        if (maxDataRows != int.MaxValue) {
            lastRow = Math.Min(lastRow, maxDataRows + 1);
        }

        int metric = AddSalesHeadersMetric(0);

        for (int row = 2; row <= lastRow; row++) {
            metric = AddSalesRangeMetric(
                metric,
                worksheet.Cell(row, 1).GetValue<int>(),
                worksheet.Cell(row, 2).GetValue<string>(),
                worksheet.Cell(row, 3).GetValue<string>(),
                worksheet.Cell(row, 4).GetValue<DateTime>(),
                worksheet.Cell(row, 5).GetValue<double>(),
                worksheet.Cell(row, 6).GetValue<int>(),
                worksheet.Cell(row, 7).GetValue<bool>(),
                worksheet.Cell(row, 8).GetValue<string>());
        }

        return metric;
    }

    private static int EpPlusReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int lastRow = worksheet.Dimension?.End.Row ?? 0;
        if (maxDataRows != int.MaxValue) {
            lastRow = Math.Min(lastRow, maxDataRows + 1);
        }

        int metric = AddSalesHeadersMetric(0);

        for (int row = 2; row <= lastRow; row++) {
            metric = AddSalesRangeMetric(
                metric,
                Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
                Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(worksheet.Cells[row, 4].Value),
                Convert.ToDouble(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
                Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
                Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
                Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int OfficeImoReadRangeStream(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;

        foreach (var chunk in reader.GetSheet("Data").ReadRangeStream(dataRange, chunkRows: 512)) {
            for (int rowOffset = 0; rowOffset < chunk.RowCount; rowOffset++) {
                int absoluteRow = chunk.StartRow + rowOffset;
                object?[] values = chunk.Rows[rowOffset];
                if (absoluteRow == 1) {
                    metric = AddSalesHeadersMetric(metric);
                    continue;
                }

                metric = AddSalesRangeMetric(
                    metric,
                    Convert.ToInt32(values[0], CultureInfo.InvariantCulture),
                    Convert.ToString(values[1], CultureInfo.InvariantCulture) ?? string.Empty,
                    Convert.ToString(values[2], CultureInfo.InvariantCulture) ?? string.Empty,
                    ReadDateCell(values[3]),
                    Convert.ToDouble(values[4], CultureInfo.InvariantCulture),
                    Convert.ToInt32(values[5], CultureInfo.InvariantCulture),
                    Convert.ToBoolean(values[6], CultureInfo.InvariantCulture),
                    Convert.ToString(values[7], CultureInfo.InvariantCulture) ?? string.Empty);
            }
        }

        return metric;
    }

    private static int ClosedXmlReadRangeStream(byte[] workbookBytes, int maxDataRows = int.MaxValue) => ClosedXmlReadRange(workbookBytes, maxDataRows);

    private static int EpPlusReadRangeStream(byte[] workbookBytes, int maxDataRows = int.MaxValue) => EpPlusReadRange(workbookBytes, maxDataRows);

    private static int OfficeImoReadSparseColumn(byte[] workbookBytes, string sparseRange, int expectedRows) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;
        int rowIndex = 0;

        foreach (object? value in reader.GetSheet("Data").ReadColumn(sparseRange)) {
            rowIndex++;
            metric = AddSparseMetric(metric, rowIndex, expectedRows, value);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse column rows, got {rowIndex}.");
        }

        return metric;
    }

    private static int OfficeImoReadSparseRows(byte[] workbookBytes, string sparseRange, int expectedRows) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;
        int rowIndex = 0;

        foreach (object?[]? row in reader.GetSheet("Data").ReadRows(sparseRange)) {
            rowIndex++;
            metric = AddSparseMetric(metric, rowIndex, expectedRows, row?[0]);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse rows, got {rowIndex}.");
        }

        return metric;
    }

    private static int ClosedXmlReadSparseColumn(byte[] workbookBytes, int expectedRows) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int metric = 0;

        for (int row = 1; row <= expectedRows; row++) {
            metric = AddSparseMetric(metric, row, expectedRows, worksheet.Cell(row, 1).GetString());
        }

        return metric;
    }

    private static int ClosedXmlReadSparseRows(byte[] workbookBytes, int expectedRows) => ClosedXmlReadSparseColumn(workbookBytes, expectedRows);

    private static int EpPlusReadSparseColumn(byte[] workbookBytes, int expectedRows) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int metric = 0;

        for (int row = 1; row <= expectedRows; row++) {
            metric = AddSparseMetric(metric, row, expectedRows, worksheet.Cells[row, 1].Text);
        }

        return metric;
    }

    private static int EpPlusReadSparseRows(byte[] workbookBytes, int expectedRows) => EpPlusReadSparseColumn(workbookBytes, expectedRows);

    private static int OfficeImoReadObjects(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;
        foreach (var row in reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange)) {
            metric = AddSalesRecordMetric(metric, row);
        }

        return metric;
    }

    private static int ClosedXmlReadObjects(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int metric = 0;

        for (int row = 2; row <= lastRow; row++) {
            var record = new ReadSalesRecord {
                Id = worksheet.Cell(row, 1).GetValue<int>(),
                Region = worksheet.Cell(row, 2).GetValue<string>(),
                Owner = worksheet.Cell(row, 3).GetValue<string>(),
                CreatedOn = worksheet.Cell(row, 4).GetValue<DateTime>(),
                Amount = worksheet.Cell(row, 5).GetValue<double>(),
                Units = worksheet.Cell(row, 6).GetValue<int>(),
                Active = worksheet.Cell(row, 7).GetValue<bool>(),
                Notes = worksheet.Cell(row, 8).GetValue<string>()
            };
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int EpPlusReadObjects(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int lastRow = worksheet.Dimension?.End.Row ?? 0;
        int metric = 0;

        for (int row = 2; row <= lastRow; row++) {
            var record = new ReadSalesRecord {
                Id = Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
                Region = Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                Owner = Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                CreatedOn = worksheet.Cells[row, 4].Value is DateTime date ? date : DateTime.FromOADate(Convert.ToDouble(worksheet.Cells[row, 4].Value, CultureInfo.InvariantCulture)),
                Amount = Convert.ToDouble(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
                Units = Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
                Active = Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
                Notes = Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty
            };
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int OfficeImoAutoFitExisting(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Load(input)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            document.GetSheet("Data").AutoFitColumns();
            document.Save(output);
        }

        return checked((int)output.Length);
    }

    private static int ClosedXmlAutoFitExisting(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using (var workbook = new XLWorkbook(input)) {
            workbook.Worksheet("Data").ColumnsUsed().AdjustToContents();
            workbook.SaveAs(output);
        }

        return checked((int)output.Length);
    }

    private static int EpPlusAutoFitExisting(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using (var package = new ExcelPackage(input)) {
            var worksheet = package.Workbook.Worksheets["Data"];
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            package.SaveAs(output);
        }

        return checked((int)output.Length);
    }

    private static int OfficeImoWriteSharedStrings(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Strings");
            sheet.CellValues(BuildSharedStringCells(rowCount), ExecutionMode.Parallel);
        }

        return checked((int)stream.Length);
    }

    private static int ClosedXmlWriteSharedStrings(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Strings");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = "Repeated value " + (row % 12);
                worksheet.Cell(row, 2).Value = "Distinct value " + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                worksheet.Cell(row, 3).Value = "Long segment " + new string((char)('A' + (row % 26)), 48);
            }

            workbook.SaveAs(stream);
        }

        return checked((int)stream.Length);
    }

    private static int EpPlusWriteSharedStrings(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Strings");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = "Repeated value " + (row % 12);
                worksheet.Cells[row, 2].Value = "Distinct value " + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                worksheet.Cells[row, 3].Value = "Long segment " + new string((char)('A' + (row % 26)), 48);
            }

            package.Save();
        }

        return checked((int)stream.Length);
    }

    private static int OfficeImoReadFormulaText(byte[] workbookBytes, int rowCount) {
        using var reader = ExcelDocumentReader.Open(workbookBytes, new ExcelReadOptions { UseCachedFormulaResult = false });
        object?[,] values = reader.GetSheet("Formulas").ReadRange($"D2:D{rowCount + 1}", ExecutionMode.Sequential);
        int metric = 0;
        for (int row = 0; row < values.GetLength(0); row++) {
            metric = AddStringMetric(metric, Convert.ToString(values[row, 0], CultureInfo.InvariantCulture));
        }

        return metric;
    }

    private static int ClosedXmlReadFormulaText(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Formulas");
        int metric = 0;
        for (int row = 2; row <= rowCount + 1; row++) {
            metric = AddStringMetric(metric, worksheet.Cell(row, 4).FormulaA1);
        }

        return metric;
    }

    private static int EpPlusReadFormulaText(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Formulas"];
        int metric = 0;
        for (int row = 2; row <= rowCount + 1; row++) {
            metric = AddStringMetric(metric, worksheet.Cells[row, 4].Formula);
        }

        return metric;
    }

    private static int OfficeImoReadSharedStrings(byte[] workbookBytes, int rowCount) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        object?[,] values = reader.GetSheet("Strings").ReadRange($"A1:C{rowCount}");
        int metric = 0;
        for (int row = 0; row < values.GetLength(0); row++) {
            for (int col = 0; col < values.GetLength(1); col++) {
                metric = AddStringMetric(metric, Convert.ToString(values[row, col], CultureInfo.InvariantCulture));
            }
        }

        return metric;
    }

    private static int ClosedXmlReadSharedStrings(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Strings");
        int metric = 0;
        for (int row = 1; row <= rowCount; row++) {
            for (int col = 1; col <= 3; col++) {
                metric = AddStringMetric(metric, worksheet.Cell(row, col).GetString());
            }
        }

        return metric;
    }

    private static int EpPlusReadSharedStrings(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Strings"];
        int metric = 0;
        for (int row = 1; row <= rowCount; row++) {
            for (int col = 1; col <= 3; col++) {
                metric = AddStringMetric(metric, worksheet.Cells[row, col].Text);
            }
        }

        return metric;
    }

    private static byte[] CreateClosedXmlWorkbookBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(worksheet, rows);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static byte[] CreateEpPlusWorkbookBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusWorksheet(worksheet, rows, includeTable: true, autoFit: true);
            package.Save();
        }

        return stream.ToArray();
    }

    private static byte[] CreateFormulaWorkbookBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Formulas");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(1, 3, "C");
            sheet.CellValue(1, 4, "Total");
            for (int row = 2; row <= rowCount + 1; row++) {
                sheet.CellValue(row, 1, (double)row);
                sheet.CellValue(row, 2, (double)(row * 2));
                sheet.CellValue(row, 3, (double)(row * 3));
                sheet.CellFormula(row, 4, $"SUM(A{row}:C{row})");
            }
        }

        return stream.ToArray();
    }

    private static byte[] CreateSharedStringWorkbookBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Strings");
            sheet.CellValues(BuildSharedStringCells(rowCount), ExecutionMode.Parallel);
        }

        return stream.ToArray();
    }

    private static byte[] CreateSparseWorkbookBytes(int lastRow) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Header");
            sheet.CellValue(lastRow, 1, "Tail");
        }

        return stream.ToArray();
    }

    private static (int Row, int Column, object Value)[] BuildSharedStringCells(int rowCount) {
        var cells = new (int Row, int Column, object Value)[rowCount * 3];
        int offset = 0;
        for (int row = 1; row <= rowCount; row++) {
            cells[offset++] = (row, 1, "Repeated value " + (row % 12));
            cells[offset++] = (row, 2, "Distinct value " + row.ToString(System.Globalization.CultureInfo.InvariantCulture));
            cells[offset++] = (row, 3, "Long segment " + new string((char)('A' + (row % 26)), 48));
        }

        return cells;
    }

    private static void PopulateEpPlusWorksheet(ExcelWorksheet worksheet, IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeTable, bool autoFit) {
        WriteHeaders(worksheet);
        for (int i = 0; i < rows.Count; i++) {
            var row = rows[i];
            int r = i + 2;
            worksheet.Cells[r, 1].Value = row.Id;
            worksheet.Cells[r, 2].Value = row.Region;
            worksheet.Cells[r, 3].Value = row.Owner;
            worksheet.Cells[r, 4].Value = row.CreatedOn;
            worksheet.Cells[r, 5].Value = row.Amount;
            worksheet.Cells[r, 6].Value = row.Units;
            worksheet.Cells[r, 7].Value = row.Active;
            worksheet.Cells[r, 8].Value = row.Notes;
        }

        if (includeTable) {
            var table = worksheet.Tables.Add(worksheet.Cells[1, 1, rows.Count + 1, 8], "SalesData");
            table.TableStyle = TableStyles.Medium2;
        }

        if (autoFit && worksheet.Dimension != null) {
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }
    }

    private static void WriteHeaders(IXLWorksheet worksheet) {
        worksheet.Cell(1, 1).Value = "Id";
        worksheet.Cell(1, 2).Value = "Region";
        worksheet.Cell(1, 3).Value = "Owner";
        worksheet.Cell(1, 4).Value = "Amount";
    }

    private static void WriteHeaders(ExcelWorksheet worksheet) {
        worksheet.Cells[1, 1].Value = "Id";
        worksheet.Cells[1, 2].Value = "Region";
        worksheet.Cells[1, 3].Value = "Owner";
        worksheet.Cells[1, 4].Value = "CreatedOn";
        worksheet.Cells[1, 5].Value = "Amount";
        worksheet.Cells[1, 6].Value = "Units";
        worksheet.Cells[1, 7].Value = "Active";
        worksheet.Cells[1, 8].Value = "Notes";
    }

    private static void WriteAppendHeaders(ExcelWorksheet worksheet) {
        worksheet.Cells[1, 1].Value = "Id";
        worksheet.Cells[1, 2].Value = "Region";
        worksheet.Cells[1, 3].Value = "Owner";
        worksheet.Cells[1, 4].Value = "Amount";
    }

    private static void ValidateComparableReadMetrics(IEnumerable<ExcelLibraryComparisonScenario> scenarios) {
        foreach (var group in scenarios.Where(scenario => IsComparableReadScenario(scenario.Scenario)).GroupBy(scenario => scenario.Scenario)) {
            var metrics = group.Select(scenario => scenario.OutputMetric).Distinct().ToArray();
            if (metrics.Length == 1) {
                continue;
            }

            string details = string.Join(", ", group.Select(scenario => scenario.Library + "=" + scenario.OutputMetric.ToString(CultureInfo.InvariantCulture)));
            throw new InvalidOperationException($"Scenario '{group.Key}' read different values across libraries: {details}.");
        }
    }

    private static bool IsComparableReadScenario(string scenario)
        => string.Equals(scenario, "read-range", StringComparison.Ordinal)
           || string.Equals(scenario, "read-range-stream", StringComparison.Ordinal)
           || string.Equals(scenario, "read-top-range", StringComparison.Ordinal)
           || string.Equals(scenario, "read-top-range-stream", StringComparison.Ordinal)
           || string.Equals(scenario, "large-sparse-column-read", StringComparison.Ordinal)
           || string.Equals(scenario, "large-sparse-row-read", StringComparison.Ordinal)
           || string.Equals(scenario, "read-objects", StringComparison.Ordinal)
           || string.Equals(scenario, "formula-heavy-read", StringComparison.Ordinal)
           || string.Equals(scenario, "shared-string-read", StringComparison.Ordinal);

    private static int AddSalesRecordMetric(int metric, ReadSalesRecord record) {
        metric = AddIntMetric(metric, record.Id);
        metric = AddStringMetric(metric, record.Region);
        metric = AddStringMetric(metric, record.Owner);
        metric = AddIntMetric(metric, record.CreatedOn.DayOfYear);
        metric = AddDoubleMetric(metric, record.Amount);
        metric = AddIntMetric(metric, record.Units);
        metric = AddIntMetric(metric, record.Active ? 1 : 0);
        return AddStringMetric(metric, record.Notes);
    }

    private static int AddSalesHeadersMetric(int metric) {
        metric = AddStringMetric(metric, "Id");
        metric = AddStringMetric(metric, "Region");
        metric = AddStringMetric(metric, "Owner");
        metric = AddStringMetric(metric, "CreatedOn");
        metric = AddStringMetric(metric, "Amount");
        metric = AddStringMetric(metric, "Units");
        metric = AddStringMetric(metric, "Active");
        return AddStringMetric(metric, "Notes");
    }

    private static int AddSalesRangeMetric(
        int metric,
        int id,
        string region,
        string owner,
        DateTime createdOn,
        double amount,
        int units,
        bool active,
        string notes) {
        metric = AddIntMetric(metric, id);
        metric = AddStringMetric(metric, region);
        metric = AddStringMetric(metric, owner);
        metric = AddIntMetric(metric, createdOn.DayOfYear);
        metric = AddDoubleMetric(metric, amount);
        metric = AddIntMetric(metric, units);
        metric = AddIntMetric(metric, active ? 1 : 0);
        return AddStringMetric(metric, notes);
    }

    private static DateTime ReadDateCell(object? value)
        => value is DateTime dateTime
            ? dateTime
            : DateTime.FromOADate(Convert.ToDouble(value, CultureInfo.InvariantCulture));

    private static int AddIntMetric(int metric, int value) {
        unchecked {
            return (metric * 397) ^ value;
        }
    }

    private static int AddDoubleMetric(int metric, double value) {
        unchecked {
            return AddIntMetric(metric, (int)Math.Round(value * 100, MidpointRounding.AwayFromZero));
        }
    }

    private static int AddStringMetric(int metric, string? value) {
        unchecked {
            int result = metric;
            if (value == null) {
                return result * 397;
            }

            for (int i = 0; i < value.Length; i++) {
                result = (result * 397) ^ value[i];
            }

            return result;
        }
    }

    private static int AddSparseMetric(int metric, int rowIndex, int expectedRows, object? value) {
        string? text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (string.IsNullOrEmpty(text)) {
            text = null;
        }

        if (rowIndex == 1) {
            if (!string.Equals(text, "Header", StringComparison.Ordinal)) {
                throw new InvalidOperationException("Sparse read did not return the first row value.");
            }

            return AddStringMetric(metric, text);
        }

        if (rowIndex == expectedRows) {
            if (!string.Equals(text, "Tail", StringComparison.Ordinal)) {
                throw new InvalidOperationException("Sparse read did not return the last row value.");
            }

            return AddStringMetric(metric, text);
        }

        if (text != null) {
            throw new InvalidOperationException($"Sparse read returned an unexpected value at row {rowIndex}.");
        }

        return AddStringMetric(metric, null);
    }

    private sealed class ExcelLibraryComparisonProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public string BuildConfiguration { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public int WarmupIterations { get; init; }
        public int MeasuredIterations { get; init; }
        public string Notes { get; init; } = string.Empty;
        public List<ExcelLibraryComparisonScenario> Scenarios { get; init; } = [];
    }

    private sealed class ExcelLibraryComparisonScenario {
        public string Scenario { get; init; } = string.Empty;
        public string Library { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
        public int OutputMetric { get; init; }
        public double AverageMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
    }

    private sealed record LibraryComparisonCase(string Library, string Notes, Func<int> Action);

    private sealed class ReadSalesRecord {
        public int Id { get; set; }
        public string Region { get; set; } = string.Empty;
        public string Owner { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
        public double Amount { get; set; }
        public int Units { get; set; }
        public bool Active { get; set; }
        public string Notes { get; set; } = string.Empty;
    }
}
