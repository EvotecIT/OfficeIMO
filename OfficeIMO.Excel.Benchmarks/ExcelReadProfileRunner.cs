using System.Data;
using System.Text.Json;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelReadProfileRunner {
    private const int DefaultRowCount = 2500;
    private const int SparseLastRow = 100_001;
    private const int HeaderLookupIterations = 10_000;
    private const int HeaderOperationIterations = 100;
    private const int LoadedTextLookupIterations = 20;
#if DEBUG
    private const string BuildConfiguration = "Debug";
#else
    private const string BuildConfiguration = "Release";
#endif
    internal const int DefaultWarmupIterations = 2;
    internal const int DefaultMeasuredIterations = 5;

    internal static string WriteProfile(
        string outputPath,
        int rowCount = DefaultRowCount,
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

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        byte[] workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        string dataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(rowCount);
        using var loadedHeaderDocument = ExcelDocument.Load(new MemoryStream(workbookBytes, writable: false), readOnly: true);
        var loadedHeaderSheet = loadedHeaderDocument.GetSheet("Data");
        using var loadedHeaderOpsStream = new MemoryStream();
        loadedHeaderOpsStream.Write(workbookBytes, 0, workbookBytes.Length);
        loadedHeaderOpsStream.Position = 0;
        using var loadedHeaderOpsDocument = ExcelDocument.Load(loadedHeaderOpsStream);
        var loadedHeaderOpsSheet = loadedHeaderOpsDocument.GetSheet("Data");
        byte[] mixedTypeWorkbookBytes = CreateMixedTypeWorkbookBytes(rowCount);
        byte[] sparseWorkbookBytes = CreateSparseWorkbookBytes(SparseLastRow);
        byte[] sharedStringHeavyWorkbookBytes = CreateSharedStringHeavyWorkbookBytes(rowCount);
        using var loadedSharedStringHeaderDocument = ExcelDocument.Load(new MemoryStream(sharedStringHeavyWorkbookBytes, writable: false), readOnly: true);
        var loadedSharedStringHeaderSheet = loadedSharedStringHeaderDocument.GetSheet("Data");
        string sparseRange = $"A1:A{SparseLastRow}";

        List<ExcelReadProfileScenario> scenarios = [];
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadObjects", "Automatic execution policy.", () => OfficeImoReadObjects(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadObjects.Sequential", "Forced sequential range conversion.", () => OfficeImoReadObjects(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadObjects.Parallel", "Forced parallel range conversion.", () => OfficeImoReadObjects(workbookBytes, dataRange, ExecutionMode.Parallel))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadObjectsAs", "Typed object materialization with automatic execution policy.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadObjectsAs.Sequential", "Typed object materialization with forced sequential range conversion.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadObjectsAs.Parallel", "Typed object materialization with forced parallel range conversion.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, ExecutionMode.Parallel)),
            new ReadProfileCase("ReadObjectsAs.CustomConverterFallback", "Typed object materialization through a custom converter hook that falls back to built-in conversion.", () => OfficeImoReadObjectsAsCustomConverterFallback(workbookBytes, dataRange)),
            new ReadProfileCase("ReadObjectsAs.CustomConverterCultureFallback", "Typed object materialization through a custom converter fallback with non-invariant culture.", () => OfficeImoReadObjectsAsCustomConverterCultureFallback(workbookBytes, dataRange))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadObjectsStreamAs", "Streaming typed object materialization without full result buffering.", () => OfficeImoReadObjectsStreamAs(workbookBytes, dataRange)),
            new ReadProfileCase("ReadObjectsStreamAs.CustomConverterFallback", "Streaming typed object materialization through a custom converter hook that falls back to built-in conversion.", () => OfficeImoReadObjectsStreamAsCustomConverterFallback(workbookBytes, dataRange)),
            new ReadProfileCase("ReadObjectsStreamAs.CustomConverterCultureFallback", "Streaming typed object materialization through a custom converter fallback with non-invariant culture.", () => OfficeImoReadObjectsStreamAsCustomConverterCultureFallback(workbookBytes, dataRange))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadRange", "Dense 2D array read with automatic execution policy.", () => OfficeImoReadRange(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadRange.Sequential", "Dense 2D array read with forced sequential conversion.", () => OfficeImoReadRange(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadRange.Parallel", "Dense 2D array read with forced parallel conversion.", () => OfficeImoReadRange(workbookBytes, dataRange, ExecutionMode.Parallel))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadRangeAsDataTable", "Automatic execution policy.", () => OfficeImoReadDataTable(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadRangeAsDataTable.Sequential", "Forced sequential range conversion.", () => OfficeImoReadDataTable(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadRangeAsDataTable.Parallel", "Forced parallel range conversion.", () => OfficeImoReadDataTable(workbookBytes, dataRange, ExecutionMode.Parallel)),
            new ReadProfileCase("ReadRangeAsDataTable.HeadersNoInference", "Header row with type inference disabled.", () => OfficeImoReadDataTableHeadersNoInference(workbookBytes, dataRange)),
            new ReadProfileCase("ReadRangeAsDataTable.NoHeadersNoInference", "Generated columns with type inference disabled.", () => OfficeImoReadDataTableNoHeadersNoInference(workbookBytes, dataRange)),
            new ReadProfileCase("ReadRangeAsDataTable.MixedTypeInference", "Infer object columns from mixed-type data.", () => OfficeImoReadDataTable(mixedTypeWorkbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadRangeAsDataTable.CustomConverterFallback", "Sequential DataTable read through a custom converter hook that falls back to built-in conversion.", () => OfficeImoReadDataTableCustomConverterFallback(workbookBytes, dataRange)),
            new ReadProfileCase("ReadRangeAsDataTable.CustomConverterCultureFallback", "Sequential DataTable read through a custom converter fallback with non-invariant culture.", () => OfficeImoReadDataTableCustomConverterCultureFallback(workbookBytes, dataRange))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadRangeStream", "Streaming row chunks with automatic execution policy.", () => OfficeImoReadRangeStream(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadRangeStream.Sequential", "Streaming row chunks with forced sequential conversion.", () => OfficeImoReadRangeStream(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadRangeStream.Parallel", "Streaming row chunks with forced parallel conversion.", () => OfficeImoReadRangeStream(workbookBytes, dataRange, ExecutionMode.Parallel)),
            new ReadProfileCase("ReadRangeStream.CustomConverterFallback", "Streaming row chunks through a custom converter hook that falls back to built-in conversion.", () => OfficeImoReadRangeStreamCustomConverterFallback(workbookBytes, dataRange)),
            new ReadProfileCase("ReadRangeStream.CustomConverterCultureFallback", "Streaming row chunks through a custom converter fallback with non-invariant culture.", () => OfficeImoReadRangeStreamCustomConverterCultureFallback(workbookBytes, dataRange))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadRows.Dense", "Dense row enumeration over the same workbook payload.", () => OfficeImoReadRowsDense(workbookBytes, dataRange)),
            new ReadProfileCase("ReadRows.Dense.CustomConverterFallback", "Dense row enumeration through a custom converter hook that falls back to built-in conversion.", () => OfficeImoReadRowsDenseCustomConverterFallback(workbookBytes, dataRange)),
            new ReadProfileCase("ReadRows.Dense.CustomConverterCultureFallback", "Dense row enumeration through a custom converter fallback with non-invariant culture.", () => OfficeImoReadRowsDenseCustomConverterCultureFallback(workbookBytes, dataRange))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadColumn.Dense", "Dense single-column enumeration over the same workbook payload.", () => OfficeImoReadDenseColumn(workbookBytes, $"A1:A{rowCount + 1}")),
            new ReadProfileCase("ReadColumn.Dense.CustomConverterFallback", "Dense single-column enumeration through a custom converter hook that falls back to built-in conversion.", () => OfficeImoReadDenseColumnCustomConverterFallback(workbookBytes, $"A1:A{rowCount + 1}")),
            new ReadProfileCase("ReadColumn.Dense.CustomConverterCultureFallback", "Dense single-column enumeration through a custom converter fallback with non-invariant culture.", () => OfficeImoReadDenseColumnCustomConverterCultureFallback(workbookBytes, $"A1:A{rowCount + 1}"))
        ], warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "GetUsedRangeA1.Cached", "Repeated used-range lookup through ExcelDocumentReader after the worksheet dimension has been resolved.", () => OfficeImoGetUsedRangeCached(workbookBytes), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "GetHeaderMap.LargeUsedRange", "Build the header lookup map from the first used row of a larger worksheet.", () => OfficeImoGetHeaderMap(workbookBytes), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "GetHeaderMap.LoadedRebuild", "Rebuild the header lookup map on an already loaded worksheet after cache invalidation.", () => OfficeImoGetHeaderMapLoadedRebuild(loadedHeaderSheet), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "GetHeaderMap.LoadedSharedStrings", "Rebuild headers when the worksheet has many unique shared strings below the header row.", () => OfficeImoGetHeaderMapLoadedRebuild(loadedSharedStringHeaderSheet), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "TryGetColumnIndexByHeader.Cached", "Repeated cached header lookup without exposing the mutable header map.", () => OfficeImoTryGetHeaderLookupLoaded(loadedHeaderSheet), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "HeaderOps.SetByHeader.Cached", "Repeated header-based data-row updates through the normal SetByHeader API.", () => OfficeImoSetByHeaderLoaded(loadedHeaderOpsSheet, rowCount), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "LoadedText.FindFirst.SharedStrings", "Repeated loaded-sheet text search over cells backed by many shared strings.", () => OfficeImoFindFirstSharedStringLoaded(loadedSharedStringHeaderSheet, rowCount), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "LoadedText.ReplaceAll.SharedStrings", "Batch replacement over a loaded worksheet with many shared-string cells.", () => OfficeImoReplaceAllSharedStrings(sharedStringHeavyWorkbookBytes, rowCount), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "HeaderOps.AutoFilterByHeaders.BatchMap", "Repeated multi-header AutoFilter updates using one internal cached header map per operation.", () => OfficeImoAutoFilterByHeadersLoaded(loadedHeaderOpsSheet), warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadColumn.LargeSparse", "Sparse A1:A100001 read with only the first and last rows populated.", () => OfficeImoReadSparseColumn(sparseWorkbookBytes, sparseRange, SparseLastRow)),
            new ReadProfileCase("ReadColumn.LargeSparse.CustomConverterFallback", "Sparse A1:A100001 read through a custom converter hook that falls back to built-in conversion.", () => OfficeImoReadSparseColumnCustomConverterFallback(sparseWorkbookBytes, sparseRange, SparseLastRow))
        ], warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "ReadRows.LargeSparse", "Sparse A1:A100001 row read with only the first and last rows populated.", () => OfficeImoReadSparseRows(sparseWorkbookBytes, sparseRange, SparseLastRow), warmupIterations, measuredIterations));
        scenarios.Add(Measure("ClosedXML", "ReadRows", "Worksheet row iteration over the same workbook payload.", () => ClosedXmlReadRows(workbookBytes), warmupIterations, measuredIterations));

        var profile = new ExcelReadProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            BuildConfiguration = BuildConfiguration,
            RowCount = rowCount,
            WarmupIterations = warmupIterations,
            MeasuredIterations = measuredIterations,
            Scenarios = scenarios
        };

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(outputPath, JsonSerializer.Serialize(profile, options));
        return outputPath;
    }

    private static ExcelReadProfileScenario Measure(
        string library,
        string name,
        string notes,
        Func<int> action,
        int warmupIterations,
        int measuredIterations) {
        var measurement = BenchmarkMeasurement.Measure(warmupIterations, measuredIterations, action);

        return new ExcelReadProfileScenario {
            Library = library,
            Name = name,
            Notes = notes,
            OutputMetric = measurement.OutputMetric,
            AverageMilliseconds = measurement.AverageMilliseconds,
            MedianMilliseconds = measurement.MedianMilliseconds,
            SamplesMilliseconds = measurement.SamplesMilliseconds.ToList()
        };
    }

    private static IReadOnlyList<ExcelReadProfileScenario> MeasureGroup(
        string library,
        IReadOnlyList<ReadProfileCase> cases,
        int warmupIterations,
        int measuredIterations) {
        if (cases.Count == 0) {
            return [];
        }

        for (int i = 0; i < cases.Count; i++) {
            if (cases[i].Action == null) {
                throw new ArgumentNullException(nameof(cases));
            }
        }

        var measurements = BenchmarkMeasurement.MeasureGroup(
            warmupIterations,
            measuredIterations,
            cases.Select(c => c.Action).ToArray());

        var scenarios = new List<ExcelReadProfileScenario>(cases.Count);
        for (int i = 0; i < cases.Count; i++) {
            var measurement = measurements[i];
            scenarios.Add(new ExcelReadProfileScenario {
                Library = library,
                Name = cases[i].Name,
                Notes = cases[i].Notes,
                OutputMetric = measurement.OutputMetric,
                AverageMilliseconds = measurement.AverageMilliseconds,
                MedianMilliseconds = measurement.MedianMilliseconds,
                SamplesMilliseconds = measurement.SamplesMilliseconds.ToList()
            });
        }

        return scenarios;
    }

    private static int OfficeImoReadObjects(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return reader.GetSheet("Data").ReadObjects(dataRange, mode).Count();
    }

    private static int OfficeImoReadObjectsAs(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange, mode).Count();
    }

    private static int OfficeImoReadObjectsAsCustomConverterFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange, ExecutionMode.Sequential).Count();
    }

    private static int OfficeImoReadObjectsAsCustomConverterCultureFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange, ExecutionMode.Sequential).Count();
    }

    private static int OfficeImoReadObjectsStreamAs(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return reader.GetSheet("Data").ReadObjectsStream<ReadSalesRecord>(dataRange).Count();
    }

    private static int OfficeImoReadObjectsStreamAsCustomConverterFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return reader.GetSheet("Data").ReadObjectsStream<ReadSalesRecord>(dataRange).Count();
    }

    private static int OfficeImoReadObjectsStreamAsCustomConverterCultureFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return reader.GetSheet("Data").ReadObjectsStream<ReadSalesRecord>(dataRange).Count();
    }

    private static int OfficeImoReadRange(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        var values = reader.GetSheet("Data").ReadRange(dataRange, mode);
        return values.GetLength(0) * values.GetLength(1);
    }

    private static int OfficeImoReadDataTable(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true, mode: mode).Rows.Count;
    }

    private static int OfficeImoReadDataTableHeadersNoInference(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes, new ExcelReadOptions { InferDataTableColumnTypes = false });
        DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true, mode: ExecutionMode.Sequential);
        return table.Rows.Count + table.Columns.Count;
    }

    private static int OfficeImoReadDataTableNoHeadersNoInference(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes, new ExcelReadOptions { InferDataTableColumnTypes = false });
        DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: false, mode: ExecutionMode.Sequential);
        return table.Rows.Count + table.Columns.Count;
    }

    private static int OfficeImoReadDataTableCustomConverterFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true, mode: ExecutionMode.Sequential);
        return table.Rows.Count + table.Columns.Count;
    }

    private static int OfficeImoReadDataTableCustomConverterCultureFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true, mode: ExecutionMode.Sequential);
        return table.Rows.Count + table.Columns.Count;
    }

    private static int OfficeImoReadRangeStream(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return CountRangeStreamRows(reader.GetSheet("Data"), dataRange, mode);
    }

    private static int OfficeImoReadRangeStreamCustomConverterFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return CountRangeStreamRows(reader.GetSheet("Data"), dataRange, ExecutionMode.Sequential);
    }

    private static int OfficeImoReadRangeStreamCustomConverterCultureFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return CountRangeStreamRows(reader.GetSheet("Data"), dataRange, ExecutionMode.Sequential);
    }

    private static int CountRangeStreamRows(ExcelSheetReader sheet, string dataRange, ExecutionMode? mode) {
        int rows = 0;
        foreach (var chunk in sheet.ReadRangeStream(dataRange, chunkRows: 512, mode: mode)) {
            rows += chunk.RowCount;
        }

        return rows;
    }

    private static int OfficeImoReadRowsDense(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return CountDenseRows(reader.GetSheet("Data"), dataRange);
    }

    private static int OfficeImoReadRowsDenseCustomConverterFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return CountDenseRows(reader.GetSheet("Data"), dataRange);
    }

    private static int OfficeImoReadRowsDenseCustomConverterCultureFallback(byte[] workbookBytes, string dataRange) {
        var options = new ExcelReadOptions {
            Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return CountDenseRows(reader.GetSheet("Data"), dataRange);
    }

    private static int CountDenseRows(ExcelSheetReader sheet, string dataRange) {
        int rows = 0;
        int populatedCells = 0;
        foreach (object?[]? row in sheet.ReadRows(dataRange)) {
            rows++;
            if (row == null) {
                continue;
            }

            for (int i = 0; i < row.Length; i++) {
                if (row[i] != null) {
                    populatedCells++;
                }
            }
        }

        return rows + populatedCells;
    }

    private static int OfficeImoReadDenseColumn(byte[] workbookBytes, string columnRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return CountColumnValues(reader.GetSheet("Data"), columnRange);
    }

    private static int OfficeImoReadDenseColumnCustomConverterFallback(byte[] workbookBytes, string columnRange) {
        var options = new ExcelReadOptions {
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return CountColumnValues(reader.GetSheet("Data"), columnRange);
    }

    private static int OfficeImoReadDenseColumnCustomConverterCultureFallback(byte[] workbookBytes, string columnRange) {
        var options = new ExcelReadOptions {
            Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return CountColumnValues(reader.GetSheet("Data"), columnRange);
    }

    private static int CountColumnValues(ExcelSheetReader sheet, string columnRange) {
        int rows = 0;
        int populated = 0;
        foreach (object? value in sheet.ReadColumn(columnRange)) {
            rows++;
            if (value != null) {
                populated++;
            }
        }

        return rows + populated;
    }

    private static int OfficeImoGetUsedRangeCached(byte[] workbookBytes) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        var sheet = reader.GetSheet("Data");
        int total = 0;
        for (int i = 0; i < HeaderLookupIterations; i++) {
            total += sheet.GetUsedRangeA1().Length;
        }

        return total;
    }

    private static int OfficeImoGetHeaderMap(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var document = ExcelDocument.Load(stream, readOnly: true);
        var map = document.GetSheet("Data").GetHeaderMap();
        return map.Count + map["Id"];
    }

    private static int OfficeImoGetHeaderMapLoadedRebuild(ExcelSheet sheet) {
        sheet.ClearHeaderCache();
        var map = sheet.GetHeaderMap();
        return map.Count + map["Id"];
    }

    private static int OfficeImoTryGetHeaderLookupLoaded(ExcelSheet sheet) {
        _ = sheet.GetHeaderMap();
        int total = 0;
        for (int i = 0; i < HeaderLookupIterations; i++) {
            if (sheet.TryGetColumnIndexByHeader("Amount", out int columnIndex)) {
                total += columnIndex;
            }
        }

        return total;
    }

    private static int OfficeImoSetByHeaderLoaded(ExcelSheet sheet, int rowCount) {
        _ = sheet.GetHeaderMap();
        int total = 0;
        for (int i = 0; i < HeaderLookupIterations; i++) {
            int row = (i % rowCount) + 2;
            sheet.SetByHeader(row, "Amount", i);
            total += row;
        }

        return total;
    }

    private static int OfficeImoFindFirstSharedStringLoaded(ExcelSheet sheet, int rowCount) {
        string expectedAddress = $"D{rowCount + 1}";
        string expectedText = $"Unique note payload {rowCount - 1:000000}";
        int total = 0;
        for (int i = 0; i < LoadedTextLookupIterations; i++) {
            string? address = sheet.FindFirst(expectedText);
            if (!string.Equals(address, expectedAddress, StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Expected {expectedAddress}, got {address ?? "<null>"}.");
            }

            total += expectedAddress.Length;
        }

        return total;
    }

    private static int OfficeImoReplaceAllSharedStrings(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream();
        stream.Write(workbookBytes, 0, workbookBytes.Length);
        stream.Position = 0;

        using var document = ExcelDocument.Load(stream);
        int replaced = document.GetSheet("Data").ReplaceAll("Unique note payload", "Processed note payload");
        if (replaced != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} replacements, got {replaced}.");
        }

        return replaced;
    }

    private static int OfficeImoAutoFilterByHeadersLoaded(ExcelSheet sheet) {
        int appliedFilters = 0;
        for (int i = 0; i < HeaderOperationIterations; i++) {
            sheet.AutoFilterByHeadersEquals(
                ("Id", new[] { "1" }),
                ("Region", new[] { "North" }),
                ("Owner", new[] { "Owner 1", "Owner 2" }),
                ("CreatedOn", new[] { "2024-01-01" }),
                ("Amount", new[] { "100" }),
                ("Units", new[] { "3" }),
                ("Active", new[] { "TRUE" }),
                ("Notes", new[] { "Note 1" }));
            appliedFilters += 8;
        }

        return appliedFilters;
    }

    private static int OfficeImoReadSparseColumn(byte[] workbookBytes, string sparseRange, int expectedRows) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return CountSparseColumn(reader.GetSheet("Data"), sparseRange, expectedRows);
    }

    private static int OfficeImoReadSparseColumnCustomConverterFallback(byte[] workbookBytes, string sparseRange, int expectedRows) {
        var options = new ExcelReadOptions {
            CellValueConverter = static _ => ExcelCellValue.NotHandled
        };
        using var reader = ExcelDocumentReader.Open(workbookBytes, options);
        return CountSparseColumn(reader.GetSheet("Data"), sparseRange, expectedRows);
    }

    private static int CountSparseColumn(ExcelSheetReader sheet, string sparseRange, int expectedRows) {
        int rowIndex = 0;

        foreach (object? value in sheet.ReadColumn(sparseRange)) {
            rowIndex++;
            ValidateSparseCell(rowIndex, expectedRows, value);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse column rows, got {rowIndex}.");
        }

        return rowIndex;
    }

    private static int OfficeImoReadSparseRows(byte[] workbookBytes, string sparseRange, int expectedRows) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int rowIndex = 0;

        foreach (object?[]? row in reader.GetSheet("Data").ReadRows(sparseRange)) {
            rowIndex++;
            ValidateSparseCell(rowIndex, expectedRows, row?[0]);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse rows, got {rowIndex}.");
        }

        return rowIndex;
    }

    private static int ClosedXmlReadRows(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int count = 0;

        for (int row = 2; row <= lastRow; row++) {
            _ = worksheet.Cell(row, 1).GetValue<int>();
            _ = worksheet.Cell(row, 5).GetValue<double>();
            count++;
        }

        return count;
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

    private static byte[] CreateMixedTypeWorkbookBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Id");
            sheet.CellValue(1, 2, "Region");
            sheet.CellValue(1, 3, "Owner");
            sheet.CellValue(1, 4, "CreatedOn");
            sheet.CellValue(1, 5, "Amount");
            sheet.CellValue(1, 6, "Units");
            sheet.CellValue(1, 7, "Active");
            sheet.CellValue(1, 8, "Notes");

            for (int i = 0; i < rowCount; i++) {
                int row = i + 2;
                bool even = (i & 1) == 0;
                sheet.CellValue(row, 1, even ? i + 1 : $"ID-{i + 1}");
                sheet.CellValue(row, 2, even ? "North" : i);
                sheet.CellValue(row, 3, even ? $"Owner {i % 13}" : i + 1000);
                sheet.CellValue(row, 4, even ? new DateTime(2024, 1, 1).AddDays(i) : $"2024-{(i % 12) + 1:00}");
                sheet.CellValue(row, 5, even ? i * 1.25D : $"Amount {i}");
                sheet.CellValue(row, 6, even ? i % 17 : $"Units {i % 17}");
                sheet.CellValue(row, 7, even ? i % 2 == 0 : $"Active {i % 2}");
                sheet.CellValue(row, 8, even ? $"Note {i}" : i * 2.0D);
            }
        }

        return stream.ToArray();
    }

    private static byte[] CreateSharedStringHeavyWorkbookBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Id");
            sheet.CellValue(1, 2, "Region");
            sheet.CellValue(1, 3, "Owner");
            sheet.CellValue(1, 4, "Notes");

            for (int i = 0; i < rowCount; i++) {
                int row = i + 2;
                sheet.CellValue(row, 1, $"ID-{i + 1:000000}");
                sheet.CellValue(row, 2, $"Region {i:000000}");
                sheet.CellValue(row, 3, $"Owner {i:000000}");
                sheet.CellValue(row, 4, $"Unique note payload {i:000000}");
            }
        }

        return stream.ToArray();
    }

    private static void ValidateSparseCell(int rowIndex, int expectedRows, object? value) {
        if (rowIndex == 1) {
            if (!Equals("Header", value)) {
                throw new InvalidOperationException("Sparse read did not return the first row value.");
            }

            return;
        }

        if (rowIndex == expectedRows) {
            if (!Equals("Tail", value)) {
                throw new InvalidOperationException("Sparse read did not return the last row value.");
            }

            return;
        }

        if (value != null) {
            throw new InvalidOperationException($"Sparse read returned an unexpected value at row {rowIndex}.");
        }
    }

    private sealed class ExcelReadProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public string BuildConfiguration { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public int WarmupIterations { get; init; }
        public int MeasuredIterations { get; init; }
        public List<ExcelReadProfileScenario> Scenarios { get; init; } = [];
    }

    private sealed class ExcelReadProfileScenario {
        public string Library { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
        public int OutputMetric { get; init; }
        public double AverageMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
    }

    private sealed record ReadProfileCase(string Name, string Notes, Func<int> Action);

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
