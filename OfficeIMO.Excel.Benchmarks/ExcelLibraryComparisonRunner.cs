using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml;
using ClosedXML.Excel;
using LargeXlsx;
using MiniExcelApi = MiniExcelLibs.MiniExcel;
using MiniExcelConfiguration = MiniExcelLibs.OpenXml.OpenXmlConfiguration;
using MiniExcelTableStyles = MiniExcelLibs.OpenXml.TableStyles;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using Sylvan.Data.Excel;
using SylvanExcelDataReader = Sylvan.Data.Excel.ExcelDataReader;
using SylvanExcelDataWriter = Sylvan.Data.Excel.ExcelDataWriter;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class ExcelLibraryComparisonRunner {
    internal const int DefaultWarmupIterations = 3;
    internal const int DefaultMeasuredIterations = 5;

    private const int DefaultRowCount = 2500;
    private const int SparseLastRow = 100_001;
    private const int HelloWorldColumnCount = 10;
    private const int BlogStringColumnCount = 20;
    private const string HelloWorldValue = "HelloWorld";
    private const string DenseHelloWorldReadRangeScenario = "dense-helloworld-read-range";
    private const string DenseHelloWorldReadStreamScenario = "dense-helloworld-read-stream";
    private const string LegacyMiniExcelHelloWorldReadRangeScenario = "miniexcel-helloworld-read-range";
    private const string LegacyMiniExcelHelloWorldReadStreamScenario = "miniexcel-helloworld-read-stream";
    private static readonly string[] HelloWorldColumnNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
    private static readonly string[] BlogStringColumnNames = Enumerable.Range(1, BlogStringColumnCount).Select(static index => "C" + index.ToString(CultureInfo.InvariantCulture)).ToArray();
    private static readonly string[] PowerShellMixedColumnNames = ["Id", "Name", "Department", "Region", "IsEnabled", "Created", "Score", "Owner", "TicketCount", "Notes"];
    private static readonly string[] PowerShellWideColumnNames;
    private static readonly string[] PowerShellMixedRegions = ["NA", "EU", "APAC", "LATAM"];
    private static readonly ExcelTabularWriteOptions CompactTabularWriteOptions = new() {
        IncludeCellReferences = false,
        UseSharedStrings = false
    };
    private static readonly XlsxStyle LargeXlsxDateTimeStyle = XlsxStyle.Default.With(new XlsxNumberFormat("yyyy-mm-dd hh:mm"));
    private static IReadOnlySet<string>? _libraryFilter;
#if DEBUG
    private const string BuildConfiguration = "Debug";
#else
    private const string BuildConfiguration = "Release";
#endif

    static ExcelLibraryComparisonRunner() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        PowerShellWideColumnNames = new[] { "Id", "Name", "Created", "Enabled" }
            .Concat(Enumerable.Range(1, 36).Select(static index => "Metric" + index.ToString(CultureInfo.InvariantCulture)))
            .ToArray();
    }

    internal static string WriteComparison(
        string outputPath,
        int rowCount = DefaultRowCount,
        bool includeLegacyEpPlus = true,
        IReadOnlyCollection<string>? scenarioFilters = null,
        int warmupIterations = DefaultWarmupIterations,
        int measuredIterations = DefaultMeasuredIterations,
        IReadOnlyCollection<string>? libraryFilters = null) {
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
        _libraryFilter = BuildLibraryFilter(libraryFilters);
        var scenarioFilter = BuildScenarioFilter(scenarioFilters);
        if (ContainsOnlyHelloWorldScenarios(scenarioFilter)) {
            var helloWorldScenarios = new List<ExcelLibraryComparisonScenario>();
            AddHelloWorldScenarioGroups(helloWorldScenarios, scenarioFilter, rowCount, warmupIterations, measuredIterations);

            var helloWorldProfile = CreateComparisonProfile(
                rowCount,
                warmupIterations,
                measuredIterations,
                "Dense HelloWorld grid comparison. Generates an A1:J(row count) workbook filled with HelloWorld and compares dense and streaming reads.",
                helloWorldScenarios);

            return WriteProfile(outputPath, helloWorldProfile);
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        var firstTableRows = rows.Take(rowCount / 2).ToList();
        var secondTableRows = rows.Skip(rowCount / 2).ToList();
        var salesDataTable = CreateSalesDataTable(rows, "SalesData");
        var salesCells = BuildSalesCells(rows);
        var objectColumnSalesDataTable = CreateObjectColumnSalesDataTable(rows, "SalesData");
        var typedObjectRows = CreateTypedObjectRows(rows);
        var dictionaryRows = CreateDictionaryRows(rows);
        var blogStringRows = CreateBlogStringRows(rowCount);
        var powerShellMixedRows = CreatePowerShellMixedRows(rowCount);
        var powerShellObjectMixedRows = CreatePowerShellObjectMixedRows(powerShellMixedRows);
        var powerShellMixedDataTable = CreatePowerShellMixedDataTable(powerShellMixedRows, "PowerShellMixed");
        var powerShellWideRows = CreatePowerShellWideRows(rowCount);
        var powerShellObjectWideRows = CreatePowerShellObjectMixedRows(powerShellWideRows);
        var powerShellWideDataTable = CreatePowerShellWideDataTable(powerShellWideRows, "PowerShellWide");
        var salesDataSet = CreateSalesDataSet(firstTableRows, secondTableRows);
        var sparseDataSet = CreateSparseDataSet(rowCount);
        var officeImoWorkbookBytes = new Lazy<byte[]>(() => ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows));
        var closedXmlWorkbookBytes = new Lazy<byte[]>(() => CreateClosedXmlWorkbookBytes(rows));
        var epPlusWorkbookBytes = new Lazy<byte[]>(() => CreateEpPlusWorkbookBytes(rows));
        var formulaWorkbookBytes = new Lazy<byte[]>(() => CreateFormulaWorkbookBytes(rowCount));
        var sharedStringWorkbookBytes = new Lazy<byte[]>(() => CreateSharedStringWorkbookBytes(rowCount));
        var sparseWorkbookBytes = new Lazy<byte[]>(() => CreateSparseWorkbookBytes(SparseLastRow));
        string dataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(rowCount);
        string firstColumnRange = $"A1:A{(rowCount + 1).ToString(CultureInfo.InvariantCulture)}";
        int topDataRows = Math.Min(rowCount, 100);
        string topDataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(topDataRows);
        int bottomDataRows = Math.Min(rowCount, 100);
        int bottomDataRowsToSkip = rowCount - bottomDataRows;
        int bottomFirstWorksheetRow = bottomDataRowsToSkip + 2;
        string bottomDataRange = $"A{bottomFirstWorksheetRow.ToString(CultureInfo.InvariantCulture)}:H{(rowCount + 1).ToString(CultureInfo.InvariantCulture)}";
        string sparseRange = $"A1:A{SparseLastRow}";
        var scenarios = new List<ExcelLibraryComparisonScenario>();

        AddScenarioGroup(scenarios, scenarioFilter, "write-bulk-report", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert objects, add table, autofit, save.", () => OfficeImoWriteBulkReport(rows)),
            new LibraryComparisonCase("ClosedXML", "Insert table, apply table style, autofit, save.", () => ClosedXmlWriteBulkReport(rows)),
            new LibraryComparisonCase("EPPlus", "Manual row population, add table, autofit, save.", () => EpPlusWriteBulkReport(rows)),
            new LibraryComparisonCase("MiniExcel", "Streaming object export with table styling and auto-width configuration.", () => MiniExcelWriteBulkReport(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-dataset-tables", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert a prepared DataSet through the normal workbook API and save.", () => OfficeImoWriteDataSetTables(salesDataSet)),
            new LibraryComparisonCase("ClosedXML", "Import prepared DataTables as two styled worksheet tables and save.", () => ClosedXmlWriteDataSetTables(salesDataSet)),
            new LibraryComparisonCase("EPPlus", "Import prepared DataTables as two styled worksheet tables and save.", () => EpPlusWriteDataSetTables(salesDataSet)),
            new LibraryComparisonCase("MiniExcel", "Streaming DataSet export with one sheet per table.", () => MiniExcelWriteDataSetTables(salesDataSet))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-dataset-tables-autofit", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert a prepared DataSet through the normal workbook API with AutoFit and save.", () => OfficeImoWriteDataSetTables(salesDataSet, autoFit: true)),
            new LibraryComparisonCase("ClosedXML", "Import prepared DataTables as two styled worksheet tables, adjust columns, and save.", () => ClosedXmlWriteDataSetTables(salesDataSet, autoFit: true)),
            new LibraryComparisonCase("EPPlus", "Import prepared DataTables as two styled worksheet tables, autofit columns, and save.", () => EpPlusWriteDataSetTables(salesDataSet, autoFit: true)),
            new LibraryComparisonCase("MiniExcel", "Streaming DataSet export with MiniExcel auto-width configuration.", () => MiniExcelWriteDataSetTables(salesDataSet, autoFit: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-dataset-headerless-tables", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert a prepared DataSet as headerless tables through the normal workbook API and save.", () => OfficeImoWriteDataSetTables(salesDataSet, includeHeaders: false)),
            new LibraryComparisonCase("ClosedXML", "Import prepared DataTables as headerless styled worksheet tables and save.", () => ClosedXmlWriteDataSetTables(salesDataSet, includeHeaders: false)),
            new LibraryComparisonCase("EPPlus", "Import prepared DataTables as headerless styled worksheet tables and save.", () => EpPlusWriteDataSetTables(salesDataSet, includeHeaders: false)),
            new LibraryComparisonCase("MiniExcel", "Streaming DataSet export without header rows.", () => MiniExcelWriteDataSetTables(salesDataSet, includeHeaders: false))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-dataset-sparse-tables", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert a sparse prepared DataSet through the normal workbook API and save.", () => OfficeImoWriteDataSetTables(sparseDataSet)),
            new LibraryComparisonCase("ClosedXML", "Import sparse prepared DataTables as styled worksheet tables and save.", () => ClosedXmlWriteDataSetTables(sparseDataSet)),
            new LibraryComparisonCase("EPPlus", "Import sparse prepared DataTables as styled worksheet tables and save.", () => EpPlusWriteDataSetTables(sparseDataSet)),
            new LibraryComparisonCase("MiniExcel", "Streaming sparse DataSet export with one sheet per table.", () => MiniExcelWriteDataSetTables(sparseDataSet))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-dataset-sparse-direct-export", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a sparse prepared DataSet through the static direct export API.", () => OfficeImoWriteDataSetDirectExport(sparseDataSet)),
            new LibraryComparisonCase("ClosedXML", "Import sparse prepared DataTables as styled worksheet tables and save.", () => ClosedXmlWriteDataSetTables(sparseDataSet)),
            new LibraryComparisonCase("EPPlus", "Import sparse prepared DataTables as styled worksheet tables and save.", () => EpPlusWriteDataSetTables(sparseDataSet)),
            new LibraryComparisonCase("MiniExcel", "Streaming sparse DataSet export with one sheet per table.", () => MiniExcelWriteDataSetTables(sparseDataSet)),
            new LibraryComparisonCase("LargeXlsx", "Streaming sparse DataSet export with one sheet per table.", () => LargeXlsxWriteDataSetPlain(sparseDataSet))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datatable-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert a prepared DataTable through the normal worksheet API and save.", () => OfficeImoWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("ClosedXML", "Import a prepared DataTable and save.", () => ClosedXmlWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("EPPlus", "Import a prepared DataTable and save.", () => EpPlusWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("MiniExcel", "Streaming DataTable export and save.", () => MiniExcelWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("LargeXlsx", "Streaming typed DataTable rows and save.", () => LargeXlsxWriteDataTable(salesDataTable))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datatable-table-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert a prepared DataTable as a styled table through the normal worksheet API and save.", () => OfficeImoWriteDataTableAsTable(salesDataTable)),
            new LibraryComparisonCase("ClosedXML", "Import a prepared DataTable as a styled worksheet table and save.", () => ClosedXmlWriteDataTable(salesDataTable, includeTable: true)),
            new LibraryComparisonCase("EPPlus", "Import a prepared DataTable as a styled worksheet table and save.", () => EpPlusWriteDataTable(salesDataTable, includeTable: true)),
            new LibraryComparisonCase("MiniExcel", "Streaming DataTable export with table styling configuration.", () => MiniExcelWriteDataTable(salesDataTable, includeTable: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datatable-object-table-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert an object-typed DataTable as a styled table through the normal worksheet API and save.", () => OfficeImoWriteDataTableAsTable(objectColumnSalesDataTable)),
            new LibraryComparisonCase("ClosedXML", "Import an object-typed DataTable as a styled worksheet table and save.", () => ClosedXmlWriteDataTable(objectColumnSalesDataTable, includeTable: true)),
            new LibraryComparisonCase("EPPlus", "Import an object-typed DataTable as a styled worksheet table and save.", () => EpPlusWriteDataTable(objectColumnSalesDataTable, includeTable: true)),
            new LibraryComparisonCase("MiniExcel", "Streaming object-typed DataTable export with table styling configuration.", () => MiniExcelWriteDataTable(objectColumnSalesDataTable, includeTable: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "build-object-datatable-dictionaries", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Build a DataTable from dictionary rows matching the normalized PowerShell object shape.", () => ObjectDataTableBuilder.FromObjects(dictionaryRows, "SalesData").Rows.Count)
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "build-object-datatable-typed", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Build a DataTable from typed object rows through the normal object projection API.", () => ObjectDataTableBuilder.FromObjects(typedObjectRows, "SalesData").Rows.Count)
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-dictionary-objects-table-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Build a DataTable from normalized dictionary rows, insert it as a styled table, and save.", () => OfficeImoWriteDataTableAsTable(ObjectDataTableBuilder.FromObjects(dictionaryRows, "SalesData")))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datareader-table", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream a DataTable-backed IDataReader as a styled table through the normal worksheet API and save.", () => OfficeImoWriteDataReaderTable(salesDataTable)),
            new LibraryComparisonCase("ClosedXML", "Import the same prepared data as a styled worksheet table and save.", () => ClosedXmlWriteDataTable(salesDataTable, includeTable: true)),
            new LibraryComparisonCase("EPPlus", "Import the same prepared data as a styled worksheet table and save.", () => EpPlusWriteDataTable(salesDataTable, includeTable: true)),
            new LibraryComparisonCase("MiniExcel", "Stream the same DataTable-backed IDataReader and save.", () => MiniExcelWriteDataReaderTable(salesDataTable))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datareader-table-autofit", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream a DataTable-backed IDataReader as a styled table, AutoFit, and save.", () => OfficeImoWriteDataReaderTable(salesDataTable, autoFit: true)),
            new LibraryComparisonCase("ClosedXML", "Import the same prepared data as a styled worksheet table, adjust columns, and save.", () => ClosedXmlWriteDataTable(salesDataTable, includeTable: true, autoFit: true)),
            new LibraryComparisonCase("EPPlus", "Import the same prepared data as a styled worksheet table, autofit columns, and save.", () => EpPlusWriteDataTable(salesDataTable, includeTable: true, autoFit: true)),
            new LibraryComparisonCase("MiniExcel", "Stream the same DataTable-backed IDataReader with table styling and auto-width configuration.", () => MiniExcelWriteDataReaderTable(salesDataTable, autoFit: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datareader-plain", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream a DataTable-backed IDataReader as plain worksheet rows through the normal worksheet API and save.", () => OfficeImoWriteDataReaderPlain(salesDataTable)),
            new LibraryComparisonCase("ClosedXML", "Import the same prepared data as plain worksheet rows and save.", () => ClosedXmlWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("EPPlus", "Import the same prepared data as plain worksheet rows and save.", () => EpPlusWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("MiniExcel", "Stream the same DataTable-backed IDataReader as plain worksheet rows and save.", () => MiniExcelWriteDataReaderPlain(salesDataTable)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Stream the same DataTable-backed DbDataReader through ExcelDataWriter and save.", () => SylvanWriteDataReaderPlain(salesDataTable)),
            new LibraryComparisonCase("LargeXlsx", "Stream the same DataTable-backed IDataReader as plain worksheet rows and save.", () => LargeXlsxWriteDataReaderPlain(salesDataTable))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datareader-direct-package", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write the prepared DataTable-backed IDataReader through the package-native OfficeIMO API.", () => OfficeImoWriteDataReaderDirectPackage(salesDataTable)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Stream the same prepared DataTable-backed reader through ExcelDataWriter.", () => SylvanWriteDataReaderPlain(salesDataTable)),
            new LibraryComparisonCase("LargeXlsx", "Stream the same prepared DataTable-backed IDataReader as plain worksheet rows.", () => LargeXlsxWriteDataReaderPlain(salesDataTable))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-datareader-compact-package", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a compact contiguous package through the package-native OfficeIMO DataReader API.", () => OfficeImoWriteDataReaderCompactPackage(salesDataTable)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Stream the same reader using implicit contiguous cell positions.", () => SylvanWriteDataReaderPlain(salesDataTable)),
            new LibraryComparisonCase("LargeXlsx", "Stream the same reader with cell references disabled.", () => LargeXlsxWriteDataReaderPlainCompact(salesDataTable))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalues-rectangle-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a prepared complete A1 rectangle with CellValues and save.", () => OfficeImoWriteCellValuesRectangle(salesCells)),
            new LibraryComparisonCase("ClosedXML", "Write the same complete A1 rectangle and save.", () => ClosedXmlWriteSalesRows(rows, includeAllColumns: true)),
            new LibraryComparisonCase("EPPlus", "Write the same complete A1 rectangle and save.", () => EpPlusWriteSalesRows(rows, includeAllColumns: true)),
            new LibraryComparisonCase("MiniExcel", "Streaming typed row export with the same columns and headers.", () => MiniExcelWriteSalesRows(rows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming typed row export with the same columns and headers.", () => LargeXlsxWriteSalesRows(rows, includeAllColumns: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalues-sparse-rectangle-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a sparse A1 rectangle with CellValues and save.", () => OfficeImoWriteCellValuesSparseRectangle(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign sparse object-typed cells with null blanks one by one and save.", () => ClosedXmlWriteCellValueObjectSparse(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign sparse object-typed cells with null blanks one by one and save.", () => EpPlusWriteCellValueObjectSparse(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalues-headerless-rectangle-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a headerless A1 rectangle with CellValues and save.", () => OfficeImoWriteCellValuesHeaderlessRectangle(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Write the same headerless rectangle and save.", () => ClosedXmlWriteHeaderlessMixedRows(rowCount)),
            new LibraryComparisonCase("EPPlus", "Write the same headerless rectangle and save.", () => EpPlusWriteHeaderlessMixedRows(rowCount)),
            new LibraryComparisonCase("LargeXlsx", "Stream the same headerless typed rectangle and save.", () => LargeXlsxWriteHeaderlessMixedRows(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-blog-2023-20-string-columns", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert a normalized 20-column all-string DTO workload inspired by the 2023 blog benchmark and save.", () => OfficeImoWriteBlogStringRows(blogStringRows)),
            new LibraryComparisonCase("ClosedXML", "Write the same normalized 20-column all-string DTO workload and save.", () => ClosedXmlWriteBlogStringRows(blogStringRows)),
            new LibraryComparisonCase("MiniExcel", "Streaming export of the same normalized 20-column all-string DTO workload and save.", () => MiniExcelWriteBlogStringRows(blogStringRows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming write of the same normalized 20-column all-string DTO workload and save.", () => LargeXlsxWriteBlogStringRows(blogStringRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-strings", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign repeated and distinct text-heavy cells one by one and save.", () => OfficeImoWriteCellValueStrings(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign repeated and distinct text-heavy cells one by one and save.", () => ClosedXmlWriteSharedStrings(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign repeated and distinct text-heavy cells one by one and save.", () => EpPlusWriteSharedStrings(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-strings-repeated", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign low-cardinality repeated text cells one by one and save.", () => OfficeImoWriteCellValueRepeatedStrings(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign low-cardinality repeated text cells one by one and save.", () => ClosedXmlWriteCellValueRepeatedStrings(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign low-cardinality repeated text cells one by one and save.", () => EpPlusWriteCellValueRepeatedStrings(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-strings-distinct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign high-cardinality distinct text cells one by one and save.", () => OfficeImoWriteCellValueDistinctStrings(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign high-cardinality distinct text cells one by one and save.", () => ClosedXmlWriteCellValueDistinctStrings(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign high-cardinality distinct text cells one by one and save.", () => EpPlusWriteCellValueDistinctStrings(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-empty-strings", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign empty and non-empty text cells one by one and save.", () => OfficeImoWriteCellValueEmptyStrings(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign empty and non-empty text cells one by one and save.", () => ClosedXmlWriteCellValueEmptyStrings(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign empty and non-empty text cells one by one and save.", () => EpPlusWriteCellValueEmptyStrings(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-numbers", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign numeric cells one by one and save.", () => OfficeImoWriteCellValueNumbers(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign numeric cells one by one and save.", () => ClosedXmlWriteCellValueNumbers(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign numeric cells one by one and save.", () => EpPlusWriteCellValueNumbers(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-scalars", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign decimal and boolean cells one by one and save.", () => OfficeImoWriteCellValueScalars(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign decimal and boolean cells one by one and save.", () => ClosedXmlWriteCellValueScalars(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign decimal and boolean cells one by one and save.", () => EpPlusWriteCellValueScalars(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-temporal", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign date and duration cells one by one and save.", () => OfficeImoWriteCellValueTemporal(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign date and duration cells one by one and save.", () => ClosedXmlWriteCellValueTemporal(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign date and duration cells one by one and save.", () => EpPlusWriteCellValueTemporal(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-object-mixed", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign mixed object-typed cells one by one and save.", () => OfficeImoWriteCellValueObjectMixed(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign mixed object-typed cells one by one and save.", () => ClosedXmlWriteCellValueObjectMixed(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign mixed object-typed cells one by one and save.", () => EpPlusWriteCellValueObjectMixed(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-object-sparse", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign sparse object-typed cells with null blanks one by one and save.", () => OfficeImoWriteCellValueObjectSparse(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign sparse object-typed cells with null blanks one by one and save.", () => ClosedXmlWriteCellValueObjectSparse(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign sparse object-typed cells with null blanks one by one and save.", () => EpPlusWriteCellValueObjectSparse(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellvalue-object-sparse-batch", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign sparse object-typed cells inside one worksheet Batch and save.", () => OfficeImoWriteCellValueObjectSparseBatch(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign sparse object-typed cells with null blanks one by one and save.", () => ClosedXmlWriteCellValueObjectSparse(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign sparse object-typed cells with null blanks one by one and save.", () => EpPlusWriteCellValueObjectSparse(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-cellformula", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Assign numeric cells and row formulas one by one and save.", () => OfficeImoWriteCellFormula(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Assign numeric cells and row formulas one by one and save.", () => ClosedXmlWriteCellFormula(rowCount)),
            new LibraryComparisonCase("EPPlus", "Assign numeric cells and row formulas one by one and save.", () => EpPlusWriteCellFormula(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-insertobjects-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert typed objects through the normal worksheet API and save.", () => OfficeImoWriteInsertObjects(rows)),
            new LibraryComparisonCase("ClosedXML", "Insert the same typed objects and save.", () => ClosedXmlWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("EPPlus", "Import the same typed objects and save.", () => EpPlusWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("MiniExcel", "Streaming typed object export and save.", () => MiniExcelWriteSalesRows(rows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming typed object export and save.", () => LargeXlsxWriteSalesRows(rows, includeAllColumns: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-objects-direct-package", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write typed rows through the package-native OfficeIMO API.", () => OfficeImoWriteObjectsDirectPackage(rows)),
            new LibraryComparisonCase("LargeXlsx", "Stream the same typed rows through its package-native writer.", () => LargeXlsxWriteSalesRows(rows, includeAllColumns: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-typed-rows-compact-package", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write typed rows as a compact contiguous package through the package-native OfficeIMO row writer.", () => OfficeImoWriteTypedRowsCompactPackage(rows)),
            new LibraryComparisonCase("LargeXlsx", "Stream the same typed rows with cell references disabled.", () => LargeXlsxWriteSalesRowsCompact(rows, includeAllColumns: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-insertobjects-autofitcolumnsfor-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert typed objects, AutoFit all exported columns through AutoFitColumnsFor, and save.", () => OfficeImoWriteInsertObjectsAutoFitColumnsFor(rows)),
            new LibraryComparisonCase("ClosedXML", "Insert the same typed objects, adjust columns, and save.", () => ClosedXmlWriteDataTable(salesDataTable, autoFit: true)),
            new LibraryComparisonCase("EPPlus", "Import the same typed objects, autofit columns, and save.", () => EpPlusWriteDataTable(salesDataTable, autoFit: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-insertobjects-partial-autofitcolumnsfor-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert typed objects, AutoFit selected exported columns through AutoFitColumnsFor, and save.", () => OfficeImoWriteInsertObjectsPartialAutoFitColumnsFor(rows)),
            new LibraryComparisonCase("ClosedXML", "Insert the same typed objects, adjust selected columns, and save.", () => ClosedXmlWriteDataTablePartialAutoFit(salesDataTable, [1, 3, 6, 8])),
            new LibraryComparisonCase("EPPlus", "Import the same typed objects, autofit selected columns, and save.", () => EpPlusWriteDataTablePartialAutoFit(salesDataTable, [1, 3, 6, 8]))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-insertobjects-flat-dictionaries-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert flat dictionary rows through the normal worksheet API and save.", () => OfficeImoWriteInsertDictionaryObjects(dictionaryRows)),
            new LibraryComparisonCase("ClosedXML", "Import the same prepared data and save.", () => ClosedXmlWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("EPPlus", "Import the same prepared data and save.", () => EpPlusWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("MiniExcel", "Streaming typed row export with the same values and save.", () => MiniExcelWriteSalesRows(rows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming typed row export with the same values and save.", () => LargeXlsxWriteSalesRows(rows, includeAllColumns: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert flat dictionary rows, AutoFit all exported columns through AutoFitColumnsFor, and save.", () => OfficeImoWriteInsertDictionaryObjectsAutoFitColumnsFor(dictionaryRows)),
            new LibraryComparisonCase("ClosedXML", "Import the same prepared data, adjust columns, and save.", () => ClosedXmlWriteDataTable(salesDataTable, autoFit: true)),
            new LibraryComparisonCase("EPPlus", "Import the same prepared data, autofit columns, and save.", () => EpPlusWriteDataTable(salesDataTable, autoFit: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-powershell-mixed-objects-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert PowerShell-like mixed dictionary objects through the normal worksheet API and save.", () => OfficeImoWriteInsertPowerShellMixedObjects(powerShellMixedRows)),
            new LibraryComparisonCase("ClosedXML", "Import the same mixed typed data and save.", () => ClosedXmlWriteDataTable(powerShellMixedDataTable)),
            new LibraryComparisonCase("EPPlus", "Import the same mixed typed data and save.", () => EpPlusWriteDataTable(powerShellMixedDataTable)),
            new LibraryComparisonCase("MiniExcel", "Streaming mixed dictionary object export and save.", () => MiniExcelWriteDictionaryObjects(powerShellMixedRows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming mixed dictionary rows and save.", () => LargeXlsxWritePowerShellMixedRows(powerShellMixedRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-powershell-psobject-mixed-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert PSObject-like mixed rows through the normal worksheet API and save.", () => OfficeImoWriteInsertPowerShellObjectMixedObjects(powerShellObjectMixedRows)),
            new LibraryComparisonCase("ClosedXML", "Import the same mixed typed data and save.", () => ClosedXmlWriteDataTable(powerShellMixedDataTable)),
            new LibraryComparisonCase("EPPlus", "Import the same mixed typed data and save.", () => EpPlusWriteDataTable(powerShellMixedDataTable)),
            new LibraryComparisonCase("MiniExcel", "Streaming mixed dictionary object export and save.", () => MiniExcelWriteDictionaryObjects(powerShellMixedRows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming mixed dictionary rows and save.", () => LargeXlsxWritePowerShellMixedRows(powerShellMixedRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-powershell-psobject-wide-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Insert PSObject-like wide telemetry rows through the normal worksheet API and save.", () => OfficeImoWriteInsertPowerShellObjectWideObjects(powerShellObjectWideRows)),
            new LibraryComparisonCase("ClosedXML", "Import the same wide typed data and save.", () => ClosedXmlWriteDataTable(powerShellWideDataTable)),
            new LibraryComparisonCase("EPPlus", "Import the same wide typed data and save.", () => EpPlusWriteDataTable(powerShellWideDataTable)),
            new LibraryComparisonCase("MiniExcel", "Streaming wide dictionary object export and save.", () => MiniExcelWriteDictionaryObjects(powerShellWideRows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming wide dictionary rows and save.", () => LargeXlsxWritePowerShellWideRows(powerShellWideRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "write-fluent-rowsfrom-direct", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write typed rows through the fluent RowsFrom API and save.", () => OfficeImoWriteFluentRowsFrom(rows)),
            new LibraryComparisonCase("ClosedXML", "Insert the same typed rows and save.", () => ClosedXmlWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("EPPlus", "Import the same typed rows and save.", () => EpPlusWriteDataTable(salesDataTable)),
            new LibraryComparisonCase("MiniExcel", "Streaming typed row export and save.", () => MiniExcelWriteSalesRows(rows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming typed row export and save.", () => LargeXlsxWriteSalesRows(rows, includeAllColumns: true))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "append-plain-rows", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Append prepared plain cells with CellValues parallel mode.", () => OfficeImoAppendPlainRows(rows)),
            new LibraryComparisonCase("ClosedXML", "Append equivalent row/cell values.", () => ClosedXmlAppendPlainRows(rows)),
            new LibraryComparisonCase("EPPlus", "Append equivalent row/cell values.", () => EpPlusAppendPlainRows(rows)),
            new LibraryComparisonCase("MiniExcel", "Streaming export of equivalent four-column row/cell values.", () => MiniExcelAppendPlainRows(rows)),
            new LibraryComparisonCase("LargeXlsx", "Streaming export of equivalent four-column row/cell values.", () => LargeXlsxWriteSalesRows(rows, includeAllColumns: false))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "copy-worksheet-package", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Copy one worksheet between workbooks with package mode, avoiding row-object materialization.", () => OfficeImoCopyWorksheetFromPackage(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("OfficeIMO.Excel Values", "Copy one worksheet between workbooks through the reader/writer values fallback.", () => OfficeImoCopyWorksheetFromValues(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ClosedXML", "Copy one worksheet between workbooks with the library worksheet-copy API.", () => ClosedXmlCopyWorksheet(closedXmlWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Copy one worksheet between workbooks with the library worksheet-copy API.", () => EpPlusCopyWorksheet(epPlusWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1 range with automatic execution policy.", () => OfficeImoReadRange(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Iterate used data cells from the same workbook payload.", () => ClosedXmlReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Iterate used data cells from the same workbook payload.", () => EpPlusReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("MiniExcel", "Stream used data rows from the same workbook payload.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan over the same workbook payload.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan over the same workbook payload.", () => SylvanReadRange(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-used-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Discover and materialize the worksheet used range with the public one-pass reader.", () => OfficeImoReadUsedRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ClosedXML", "Resolve and iterate used data cells from the same workbook payload.", () => ClosedXmlReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Resolve and iterate used data cells from the same workbook payload.", () => EpPlusReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("MiniExcel", "Stream used data rows from the same workbook payload.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan over the same workbook payload.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan over the same workbook payload.", () => SylvanReadRange(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-datareader", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Forward-only IDataReader scan over the requested A1 range.", () => OfficeImoReadDataReader(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan over the same workbook payload.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan over the same workbook payload.", () => SylvanReadRange(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-datareader-readonly", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read every row through IDataReader.Read without copying values into a consumer buffer.", () => OfficeImoReadDataReaderRowsOnly(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ExcelDataReader", "Read every row through IExcelDataReader.Read without copying values into a consumer buffer.", () => ExcelDataReaderReadRowsOnly(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Read every row through DbDataReader.Read without copying values into a consumer buffer.", () => SylvanReadRowsOnly(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-datareader-first-column", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read every row from the full IDataReader range while accessing only the Id column.", () => OfficeImoReadDataReaderFirstColumn(officeImoWorkbookBytes.Value, dataRange, rowCount)),
            new LibraryComparisonCase("ExcelDataReader", "Read every row from the full IExcelDataReader worksheet while accessing only the Id column.", () => ExcelDataReaderReadRangeFirstColumn(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Read every row from the full DbDataReader worksheet while accessing only the Id column.", () => SylvanReadRangeFirstColumn(officeImoWorkbookBytes.Value, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-datareader-getvalues", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Forward-only IDataReader scan using GetValues over the requested A1 range.", () => OfficeImoReadDataReaderGetValues(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan using GetValues over the same workbook payload.", () => ExcelDataReaderReadRangeGetValues(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan using GetValues over the same workbook payload.", () => SylvanReadRangeGetValues(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-datareader-typed", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Forward-only IDataReader scan with typed field access over the requested A1 range.", () => OfficeImoReadDataReaderTyped(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan with typed field access over the same workbook payload.", () => ExcelDataReaderReadRangeTyped(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan with typed field access over the same workbook payload.", () => SylvanReadRange(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-range-decimal", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1 range with NumericAsDecimal enabled.", () => OfficeImoReadRangeDecimal(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Iterate used data cells and read Amount as decimal.", () => ClosedXmlReadRangeDecimal(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Iterate used data cells and read Amount as decimal.", () => EpPlusReadRangeDecimal(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("MiniExcel", "Stream used data rows and read Amount as decimal.", () => MiniExcelReadRangeDecimal(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only scan and read Amount as decimal.", () => ExcelDataReaderReadRangeDecimal(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only scan and read Amount as decimal.", () => SylvanReadRangeDecimal(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "enumerate-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Enumerate typed cells from the requested A1 range.", () => OfficeImoEnumerateRange(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Enumerate cells from the requested A1 range.", () => ClosedXmlEnumerateRange(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("EPPlus", "Enumerate cells from the requested A1 range.", () => EpPlusEnumerateRange(officeImoWorkbookBytes.Value, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "enumerate-top-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Enumerate the first 100 data rows from a larger sheet.", () => OfficeImoEnumerateRange(officeImoWorkbookBytes.Value, topDataRange)),
            new LibraryComparisonCase("ClosedXML", "Enumerate the first 100 data rows from the same larger sheet.", () => ClosedXmlEnumerateRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("EPPlus", "Enumerate the first 100 data rows from the same larger sheet.", () => EpPlusEnumerateRange(officeImoWorkbookBytes.Value, topDataRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "enumerate-cells", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Enumerate all typed cells from the worksheet.", () => OfficeImoEnumerateCells(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ClosedXML", "Enumerate all used cells from the worksheet.", () => ClosedXmlEnumerateRange(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("EPPlus", "Enumerate all used cells from the worksheet.", () => EpPlusEnumerateRange(officeImoWorkbookBytes.Value, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "enumerate-first-column-from-wide-sheet", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Enumerate A1:A(row count + header) from the same 8-column workbook.", () => OfficeImoEnumerateFirstColumn(officeImoWorkbookBytes.Value, firstColumnRange, rowCount)),
            new LibraryComparisonCase("ClosedXML", "Enumerate A1:A(row count + header) from the same 8-column workbook.", () => ClosedXmlEnumerateFirstColumn(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("EPPlus", "Enumerate A1:A(row count + header) from the same 8-column workbook.", () => EpPlusEnumerateFirstColumn(officeImoWorkbookBytes.Value, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-first-column-from-wide-sheet", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1:A(row count + header) from the same 8-column workbook.", () => OfficeImoReadFirstColumn(officeImoWorkbookBytes.Value, firstColumnRange, rowCount)),
            new LibraryComparisonCase("ClosedXML", "Read A1:A(row count + header) from the same 8-column workbook.", () => ClosedXmlReadFirstColumn(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("EPPlus", "Read A1:A(row count + header) from the same 8-column workbook.", () => EpPlusReadFirstColumn(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("MiniExcel", "Stream A1:A(row count + header) from the same 8-column workbook.", () => MiniExcelReadFirstColumn(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan of the first column from the same 8-column workbook.", () => ExcelDataReaderReadFirstColumn(officeImoWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan of the first column from the same 8-column workbook.", () => SylvanReadFirstColumn(officeImoWorkbookBytes.Value, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-top-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read the first 100 data rows from a larger sheet.", () => OfficeImoReadRange(officeImoWorkbookBytes.Value, topDataRange)),
            new LibraryComparisonCase("ClosedXML", "Read the first 100 data rows from the same larger sheet.", () => ClosedXmlReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("EPPlus", "Read the first 100 data rows from the same larger sheet.", () => EpPlusReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("MiniExcel", "Stream the first 100 data rows from the same larger sheet.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan of the first 100 data rows.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan of the first 100 data rows.", () => SylvanReadRange(officeImoWorkbookBytes.Value, topDataRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-bottom-range", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read the last 100 data rows from a larger sheet by requested range.", () => OfficeImoReadRange(officeImoWorkbookBytes.Value, bottomDataRange, rangeIncludesHeader: false)),
            new LibraryComparisonCase("ClosedXML", "Read the last 100 data rows from the same larger sheet.", () => ClosedXmlReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("EPPlus", "Read the last 100 data rows from the same larger sheet.", () => EpPlusReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("MiniExcel", "Stream the last 100 data rows from the same larger sheet.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan to the last 100 data rows.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan to the last 100 data rows.", () => SylvanReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-datatable", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Materialize A1 range as a DataTable.", () => OfficeImoReadDataTable(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Manual DataTable materialization from the same worksheet rows.", () => ClosedXmlReadDataTable(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Manual DataTable materialization from the same worksheet rows.", () => EpPlusReadDataTable(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("MiniExcel", "Materialize the same worksheet rows through QueryAsDataTable.", () => MiniExcelReadDataTable(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ExcelDataReader", "Manual DataTable materialization from IExcelDataReader.", () => ExcelDataReaderReadDataTable(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Materialize worksheet rows from DbDataReader into a DataTable.", () => SylvanReadDataTable(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-range-stream", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream A1 range in row chunks with automatic execution policy.", () => OfficeImoReadRangeStream(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Iterate used data cells row-by-row from the same workbook payload.", () => ClosedXmlReadRangeStream(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Iterate used data cells row-by-row from the same workbook payload.", () => EpPlusReadRangeStream(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("MiniExcel", "Stream used data rows with deferred execution from the same workbook payload.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader stream over used data rows.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader stream over used data rows.", () => SylvanReadRange(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-top-range-stream", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream the first 100 data rows from a larger sheet.", () => OfficeImoReadRangeStream(officeImoWorkbookBytes.Value, topDataRange)),
            new LibraryComparisonCase("ClosedXML", "Read the first 100 data rows from the same larger sheet row-by-row.", () => ClosedXmlReadRangeStream(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("EPPlus", "Read the first 100 data rows from the same larger sheet row-by-row.", () => EpPlusReadRangeStream(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("MiniExcel", "Stream the first 100 data rows with deferred execution from the same larger sheet.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader stream of the first 100 data rows.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader stream of the first 100 data rows.", () => SylvanReadRange(officeImoWorkbookBytes.Value, topDataRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-bottom-range-stream", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream the last 100 data rows from a larger sheet by requested range.", () => OfficeImoReadRangeStream(officeImoWorkbookBytes.Value, bottomDataRange, rangeIncludesHeader: false)),
            new LibraryComparisonCase("ClosedXML", "Read the last 100 data rows from the same larger sheet row-by-row.", () => ClosedXmlReadRangeStream(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("EPPlus", "Read the last 100 data rows from the same larger sheet row-by-row.", () => EpPlusReadRangeStream(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("MiniExcel", "Stream the last 100 data rows with deferred execution from the same larger sheet.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader stream to the last 100 data rows.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader stream to the last 100 data rows.", () => SylvanReadRange(officeImoWorkbookBytes.Value, bottomDataRows, bottomDataRowsToSkip))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-top-range-stream-small-chunks", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream the first 100 data rows from a larger sheet in small row chunks.", () => OfficeImoReadRangeStream(officeImoWorkbookBytes.Value, topDataRange, chunkRows: 10)),
            new LibraryComparisonCase("ClosedXML", "Read the first 100 data rows from the same larger sheet row-by-row.", () => ClosedXmlReadRangeStream(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("EPPlus", "Read the first 100 data rows from the same larger sheet row-by-row.", () => EpPlusReadRangeStream(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("MiniExcel", "Stream the first 100 data rows with deferred execution from the same larger sheet.", () => MiniExcelReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader stream of the first 100 data rows.", () => ExcelDataReaderReadRange(officeImoWorkbookBytes.Value, topDataRows)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader stream of the first 100 data rows.", () => SylvanReadRange(officeImoWorkbookBytes.Value, topDataRows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "large-sparse-column-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1:A100001 with only first and last rows populated.", () => OfficeImoReadSparseColumn(sparseWorkbookBytes.Value, sparseRange, SparseLastRow)),
            new LibraryComparisonCase("ClosedXML", "Read A1:A100001 with only first and last rows populated.", () => ClosedXmlReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("EPPlus", "Read A1:A100001 with only first and last rows populated.", () => EpPlusReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("MiniExcel", "Stream A1:A100001 with only first and last rows populated.", () => MiniExcelReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader sparse-column scan.", () => ExcelDataReaderReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader sparse-column scan.", () => SylvanReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "large-sparse-row-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1:A100001 as rows with only first and last rows populated.", () => OfficeImoReadSparseRows(sparseWorkbookBytes.Value, sparseRange, SparseLastRow)),
            new LibraryComparisonCase("ClosedXML", "Read A1:A100001 as rows with only first and last rows populated.", () => ClosedXmlReadSparseRows(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("EPPlus", "Read A1:A100001 as rows with only first and last rows populated.", () => EpPlusReadSparseRows(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("MiniExcel", "Stream A1:A100001 rows with only first and last rows populated.", () => MiniExcelReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader sparse-row scan.", () => ExcelDataReaderReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader sparse-row scan.", () => SylvanReadSparseColumn(sparseWorkbookBytes.Value, SparseLastRow))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-objects", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Typed materialization with ReadObjects<T>.", () => OfficeImoReadObjects(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Manual typed materialization from the same worksheet rows.", () => ClosedXmlReadObjects(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Manual typed materialization from the same worksheet rows.", () => EpPlusReadObjects(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("MiniExcel", "Typed row materialization through Query<T> from the same worksheet rows.", () => MiniExcelReadObjects(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ExcelDataReader", "Manual typed materialization from IExcelDataReader.", () => ExcelDataReaderReadObjects(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Manual typed materialization from DbDataReader.", () => SylvanReadObjects(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "read-objects-stream", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Streaming typed scan with ReadObjectsStream<T>.", () => OfficeImoReadObjectsStream(officeImoWorkbookBytes.Value, dataRange)),
            new LibraryComparisonCase("ClosedXML", "Manual row-by-row typed scan from the same worksheet rows.", () => ClosedXmlReadObjectsStream(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Manual row-by-row typed scan from the same worksheet rows.", () => EpPlusReadObjectsStream(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("MiniExcel", "Streaming typed scan through deferred Query<T> from the same worksheet rows.", () => MiniExcelReadObjectsStream(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only typed scan from IExcelDataReader.", () => ExcelDataReaderReadObjectsStream(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only typed scan from DbDataReader.", () => SylvanReadObjectsStream(officeImoWorkbookBytes.Value))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "autofit-existing", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Load existing workbook, autofit columns, save.", () => OfficeImoAutoFitExisting(officeImoWorkbookBytes.Value)),
            new LibraryComparisonCase("ClosedXML", "Load existing workbook, autofit columns, save.", () => ClosedXmlAutoFitExisting(closedXmlWorkbookBytes.Value)),
            new LibraryComparisonCase("EPPlus", "Load existing workbook, autofit columns, save.", () => EpPlusAutoFitExisting(epPlusWorkbookBytes.Value))
        ]);

        AddReportWorkbookScenarioGroups(scenarios, scenarioFilter, powerShellMixedRows, powerShellMixedDataTable, warmupIterations, measuredIterations);
        AddRealWorldScenarioGroups(scenarios, scenarioFilter, rows, warmupIterations, measuredIterations);

        AddScenarioGroup(scenarios, scenarioFilter, "large-shared-strings", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write repeated and distinct text-heavy cells.", () => OfficeImoWriteSharedStrings(rowCount)),
            new LibraryComparisonCase("ClosedXML", "Write repeated and distinct text-heavy cells.", () => ClosedXmlWriteSharedStrings(rowCount)),
            new LibraryComparisonCase("EPPlus", "Write repeated and distinct text-heavy cells.", () => EpPlusWriteSharedStrings(rowCount)),
            new LibraryComparisonCase("MiniExcel", "Streaming export of repeated and distinct text-heavy cells.", () => MiniExcelWriteSharedStrings(rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "formula-heavy-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read formula text with cached formula results disabled.", () => OfficeImoReadFormulaText(formulaWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("ClosedXML", "Read formula A1 text from formula cells.", () => ClosedXmlReadFormulaText(formulaWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("EPPlus", "Read formula text from formula cells.", () => EpPlusReadFormulaText(formulaWorkbookBytes.Value, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, "shared-string-read", warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read repeated shared string payload.", () => OfficeImoReadSharedStrings(sharedStringWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("ClosedXML", "Read repeated shared string payload.", () => ClosedXmlReadSharedStrings(sharedStringWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("EPPlus", "Read repeated shared string payload.", () => EpPlusReadSharedStrings(sharedStringWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("MiniExcel", "Stream repeated shared string payload.", () => MiniExcelReadSharedStrings(sharedStringWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader read of repeated shared string payload.", () => ExcelDataReaderReadSharedStrings(sharedStringWorkbookBytes.Value, rowCount)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader read of repeated shared string payload.", () => SylvanReadSharedStrings(sharedStringWorkbookBytes.Value, rowCount))
        ]);

        AddHelloWorldScenarioGroups(scenarios, scenarioFilter, rowCount, warmupIterations, measuredIterations);

        if (scenarios.Count == 0) {
            throw new ArgumentException("No comparison scenarios matched the requested --scenario filter.");
        }

        var profile = CreateComparisonProfile(
            rowCount,
            warmupIterations,
            measuredIterations,
            "Local opt-in comparison. Not intended for CI gating.",
            scenarios);

        if (includeLegacyEpPlus) {
            profile.Scenarios.AddRange(RunLegacyEpPlusComparison(rowCount, scenarioFilter, warmupIterations, measuredIterations));
        }

        return WriteProfile(outputPath, profile);
    }

    internal static string WritePackageProfile(
        string outputPath,
        int rowCount = DefaultRowCount,
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
        var firstTableRows = rows.Take(rowCount / 2).ToList();
        var secondTableRows = rows.Skip(rowCount / 2).ToList();
        var salesDataTable = CreateSalesDataTable(rows, "SalesData");
        var salesCells = BuildSalesCells(rows);
        var dictionaryRows = CreateDictionaryRows(rows);
        var legacyDictionaryRows = CreateLegacyDictionaryRows(rows);
        var blogStringRows = CreateBlogStringRows(rowCount);
        var powerShellMixedRows = CreatePowerShellMixedRows(rowCount);
        var powerShellObjectMixedRows = CreatePowerShellObjectMixedRows(powerShellMixedRows);
        var powerShellMixedDataTable = CreatePowerShellMixedDataTable(powerShellMixedRows, "PowerShellMixed");
        var powerShellWideRows = CreatePowerShellWideRows(rowCount);
        var powerShellObjectWideRows = CreatePowerShellObjectMixedRows(powerShellWideRows);
        var powerShellWideDataTable = CreatePowerShellWideDataTable(powerShellWideRows, "PowerShellWide");
        var salesDataSet = CreateSalesDataSet(firstTableRows, secondTableRows);
        var sparseDataSet = CreateSparseDataSet(rowCount);
        var officeImoWorkbookBytes = new Lazy<byte[]>(() => ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows));
        var closedXmlWorkbookBytes = new Lazy<byte[]>(() => CreateClosedXmlWorkbookBytes(rows));
        var epPlusWorkbookBytes = new Lazy<byte[]>(() => CreateEpPlusWorkbookBytes(rows));

        var scenarios = new List<ExcelPackageProfileScenario>();
        AddPackageProfileGroup(scenarios, scenarioFilter, "write-bulk-report", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert objects, add table, autofit, save.", () => OfficeImoWriteBulkReportBytes(rows)),
            new PackageProfileCase("ClosedXML", "Insert table, apply table style, autofit, save.", () => ClosedXmlWriteBulkReportBytes(rows)),
            new PackageProfileCase("EPPlus", "Manual row population, add table, autofit, save.", () => EpPlusWriteBulkReportBytes(rows)),
            new PackageProfileCase("MiniExcel", "Streaming object export with table styling and auto-width configuration.", () => MiniExcelWriteBulkReportBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-dataset-tables", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert a prepared DataSet through the normal workbook API and save.", () => OfficeImoWriteDataSetTablesBytes(salesDataSet)),
            new PackageProfileCase("ClosedXML", "Import prepared DataTables as two styled worksheet tables and save.", () => ClosedXmlWriteDataSetTablesBytes(salesDataSet)),
            new PackageProfileCase("EPPlus", "Import prepared DataTables as two styled worksheet tables and save.", () => EpPlusWriteDataSetTablesBytes(salesDataSet)),
            new PackageProfileCase("MiniExcel", "Streaming DataSet export with one sheet per table.", () => MiniExcelWriteDataSetTablesBytes(salesDataSet))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-dataset-sparse-tables", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert a sparse prepared DataSet through the normal workbook API and save.", () => OfficeImoWriteDataSetTablesBytes(sparseDataSet)),
            new PackageProfileCase("ClosedXML", "Import sparse prepared DataTables as styled worksheet tables and save.", () => ClosedXmlWriteDataSetTablesBytes(sparseDataSet)),
            new PackageProfileCase("EPPlus", "Import sparse prepared DataTables as styled worksheet tables and save.", () => EpPlusWriteDataSetTablesBytes(sparseDataSet)),
            new PackageProfileCase("MiniExcel", "Streaming sparse DataSet export with one sheet per table.", () => MiniExcelWriteDataSetTablesBytes(sparseDataSet))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-dataset-sparse-direct-export", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a sparse prepared DataSet through the static direct export API.", () => OfficeImoWriteDataSetDirectExportBytes(sparseDataSet)),
            new PackageProfileCase("ClosedXML", "Import sparse prepared DataTables as styled worksheet tables and save.", () => ClosedXmlWriteDataSetTablesBytes(sparseDataSet)),
            new PackageProfileCase("EPPlus", "Import sparse prepared DataTables as styled worksheet tables and save.", () => EpPlusWriteDataSetTablesBytes(sparseDataSet)),
            new PackageProfileCase("MiniExcel", "Streaming sparse DataSet export with one sheet per table.", () => MiniExcelWriteDataSetTablesBytes(sparseDataSet)),
            new PackageProfileCase("LargeXlsx", "Streaming sparse DataSet export with one sheet per table.", () => LargeXlsxWriteDataSetPlainBytes(sparseDataSet))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-dataset-tables-autofit", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert a prepared DataSet through the normal workbook API with AutoFit and save.", () => OfficeImoWriteDataSetTablesBytes(salesDataSet, autoFit: true)),
            new PackageProfileCase("ClosedXML", "Import prepared DataTables as two styled worksheet tables, adjust columns, and save.", () => ClosedXmlWriteDataSetTablesBytes(salesDataSet, autoFit: true)),
            new PackageProfileCase("EPPlus", "Import prepared DataTables as two styled worksheet tables, autofit columns, and save.", () => EpPlusWriteDataSetTablesBytes(salesDataSet, autoFit: true)),
            new PackageProfileCase("MiniExcel", "Streaming DataSet export with MiniExcel auto-width configuration.", () => MiniExcelWriteDataSetTablesBytes(salesDataSet, autoFit: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-datatable-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert a prepared DataTable through the normal worksheet API and save.", () => OfficeImoWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("ClosedXML", "Import a prepared DataTable and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("EPPlus", "Import a prepared DataTable and save.", () => EpPlusWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming DataTable export and save.", () => MiniExcelWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("LargeXlsx", "Streaming typed DataTable rows and save.", () => LargeXlsxWriteDataTableBytes(salesDataTable))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-datatable-table-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert a prepared DataTable as a styled table through the normal worksheet API and save.", () => OfficeImoWriteDataTableAsTableBytes(salesDataTable)),
            new PackageProfileCase("ClosedXML", "Import a prepared DataTable as a styled worksheet table and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable, includeTable: true)),
            new PackageProfileCase("EPPlus", "Import a prepared DataTable as a styled worksheet table and save.", () => EpPlusWriteDataTableBytes(salesDataTable, includeTable: true)),
            new PackageProfileCase("MiniExcel", "Streaming DataTable export with table styling configuration.", () => MiniExcelWriteDataTableBytes(salesDataTable, includeTable: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-datareader-table", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Stream a DataTable-backed IDataReader as a styled table through the normal worksheet API and save.", () => OfficeImoWriteDataReaderTableBytes(salesDataTable)),
            new PackageProfileCase("ClosedXML", "Import the same prepared data as a styled worksheet table and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable, includeTable: true)),
            new PackageProfileCase("EPPlus", "Import the same prepared data as a styled worksheet table and save.", () => EpPlusWriteDataTableBytes(salesDataTable, includeTable: true)),
            new PackageProfileCase("MiniExcel", "Stream the same DataTable-backed IDataReader and save.", () => MiniExcelWriteDataReaderTableBytes(salesDataTable))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-datareader-table-autofit", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Stream a DataTable-backed IDataReader as a styled table, AutoFit, and save.", () => OfficeImoWriteDataReaderTableBytes(salesDataTable, autoFit: true)),
            new PackageProfileCase("ClosedXML", "Import the same prepared data as a styled worksheet table, adjust columns, and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable, includeTable: true, autoFit: true)),
            new PackageProfileCase("EPPlus", "Import the same prepared data as a styled worksheet table, autofit columns, and save.", () => EpPlusWriteDataTableBytes(salesDataTable, includeTable: true, autoFit: true)),
            new PackageProfileCase("MiniExcel", "Stream the same DataTable-backed IDataReader with table styling and auto-width configuration.", () => MiniExcelWriteDataReaderTableBytes(salesDataTable, autoFit: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-datareader-plain", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Stream a DataTable-backed IDataReader as plain worksheet rows through the normal worksheet API and save.", () => OfficeImoWriteDataReaderPlainBytes(salesDataTable)),
            new PackageProfileCase("ClosedXML", "Import the same prepared data as plain worksheet rows and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("EPPlus", "Import the same prepared data as plain worksheet rows and save.", () => EpPlusWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("MiniExcel", "Stream the same DataTable-backed IDataReader as plain worksheet rows and save.", () => MiniExcelWriteDataReaderPlainBytes(salesDataTable)),
            new PackageProfileCase("Sylvan.Data.Excel", "Stream the same DataTable-backed DbDataReader through ExcelDataWriter and save.", () => SylvanWriteDataReaderPlainBytes(salesDataTable)),
            new PackageProfileCase("LargeXlsx", "Stream the same DataTable-backed IDataReader as plain worksheet rows and save.", () => LargeXlsxWriteDataReaderPlainBytes(salesDataTable))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-datareader-direct-package", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write the prepared DataTable-backed IDataReader through the package-native OfficeIMO API.", () => OfficeImoWriteDataReaderDirectPackageBytes(salesDataTable)),
            new PackageProfileCase("Sylvan.Data.Excel", "Stream the same prepared DataTable-backed reader through ExcelDataWriter.", () => SylvanWriteDataReaderPlainBytes(salesDataTable)),
            new PackageProfileCase("LargeXlsx", "Stream the same prepared DataTable-backed IDataReader as plain worksheet rows.", () => LargeXlsxWriteDataReaderPlainBytes(salesDataTable))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-datareader-compact-package", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a compact contiguous package through the package-native OfficeIMO DataReader API.", () => OfficeImoWriteDataReaderCompactPackageBytes(salesDataTable)),
            new PackageProfileCase("Sylvan.Data.Excel", "Stream the same reader using implicit contiguous cell positions.", () => SylvanWriteDataReaderPlainBytes(salesDataTable)),
            new PackageProfileCase("LargeXlsx", "Stream the same reader with cell references disabled.", () => LargeXlsxWriteDataReaderPlainBytes(salesDataTable, requireCellReferences: false))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalues-rectangle-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a prepared complete A1 rectangle with CellValues and save.", () => OfficeImoWriteCellValuesRectangleBytes(salesCells)),
            new PackageProfileCase("ClosedXML", "Write the same complete A1 rectangle and save.", () => ClosedXmlWriteSalesRowsBytes(rows, includeAllColumns: true)),
            new PackageProfileCase("EPPlus", "Write the same complete A1 rectangle and save.", () => EpPlusWriteSalesRowsBytes(rows, includeAllColumns: true)),
            new PackageProfileCase("MiniExcel", "Streaming typed row export with the same columns and headers.", () => MiniExcelWriteSalesRowsBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Streaming typed row export with the same columns and headers.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalues-headerless-rectangle-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a headerless A1 rectangle with CellValues and save.", () => OfficeImoWriteCellValuesHeaderlessRectangleBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Write the same headerless rectangle and save.", () => ClosedXmlWriteHeaderlessMixedRowsBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Write the same headerless rectangle and save.", () => EpPlusWriteHeaderlessMixedRowsBytes(rowCount)),
            new PackageProfileCase("LargeXlsx", "Stream the same headerless typed rectangle and save.", () => LargeXlsxWriteHeaderlessMixedRowsBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-blog-2023-20-string-columns", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert a normalized 20-column all-string DTO workload inspired by the 2023 blog benchmark and save.", () => OfficeImoWriteBlogStringRowsBytes(blogStringRows)),
            new PackageProfileCase("ClosedXML", "Write the same normalized 20-column all-string DTO workload and save.", () => ClosedXmlWriteBlogStringRowsBytes(blogStringRows)),
            new PackageProfileCase("MiniExcel", "Streaming export of the same normalized 20-column all-string DTO workload and save.", () => MiniExcelWriteBlogStringRowsBytes(blogStringRows)),
            new PackageProfileCase("LargeXlsx", "Streaming write of the same normalized 20-column all-string DTO workload and save.", () => LargeXlsxWriteBlogStringRowsBytes(blogStringRows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-strings", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign repeated and distinct text-heavy cells one by one and save.", () => OfficeImoWriteCellValueStringsBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign repeated and distinct text-heavy cells one by one and save.", () => ClosedXmlWriteSharedStringsBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign repeated and distinct text-heavy cells one by one and save.", () => EpPlusWriteSharedStringsBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-strings-repeated", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign low-cardinality repeated text cells one by one and save.", () => OfficeImoWriteCellValueRepeatedStringsBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign low-cardinality repeated text cells one by one and save.", () => ClosedXmlWriteCellValueRepeatedStringsBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign low-cardinality repeated text cells one by one and save.", () => EpPlusWriteCellValueRepeatedStringsBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-strings-distinct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign high-cardinality distinct text cells one by one and save.", () => OfficeImoWriteCellValueDistinctStringsBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign high-cardinality distinct text cells one by one and save.", () => ClosedXmlWriteCellValueDistinctStringsBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign high-cardinality distinct text cells one by one and save.", () => EpPlusWriteCellValueDistinctStringsBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-empty-strings", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign empty and non-empty text cells one by one and save.", () => OfficeImoWriteCellValueEmptyStringsBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign empty and non-empty text cells one by one and save.", () => ClosedXmlWriteCellValueEmptyStringsBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign empty and non-empty text cells one by one and save.", () => EpPlusWriteCellValueEmptyStringsBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-numbers", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign numeric cells one by one and save.", () => OfficeImoWriteCellValueNumbersBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign numeric cells one by one and save.", () => ClosedXmlWriteCellValueNumbersBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign numeric cells one by one and save.", () => EpPlusWriteCellValueNumbersBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-scalars", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign decimal and boolean cells one by one and save.", () => OfficeImoWriteCellValueScalarsBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign decimal and boolean cells one by one and save.", () => ClosedXmlWriteCellValueScalarsBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign decimal and boolean cells one by one and save.", () => EpPlusWriteCellValueScalarsBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-temporal", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign date and duration cells one by one and save.", () => OfficeImoWriteCellValueTemporalBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign date and duration cells one by one and save.", () => ClosedXmlWriteCellValueTemporalBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign date and duration cells one by one and save.", () => EpPlusWriteCellValueTemporalBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-object-mixed", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign mixed object-typed cells one by one and save.", () => OfficeImoWriteCellValueObjectMixedBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign mixed object-typed cells one by one and save.", () => ClosedXmlWriteCellValueObjectMixedBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign mixed object-typed cells one by one and save.", () => EpPlusWriteCellValueObjectMixedBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-object-sparse", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign sparse object-typed cells with null blanks one by one and save.", () => OfficeImoWriteCellValueObjectSparseBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign sparse object-typed cells with null blanks one by one and save.", () => ClosedXmlWriteCellValueObjectSparseBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign sparse object-typed cells with null blanks one by one and save.", () => EpPlusWriteCellValueObjectSparseBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellvalue-object-sparse-batch", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign sparse object-typed cells inside one worksheet Batch and save.", () => OfficeImoWriteCellValueObjectSparseBatchBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign sparse object-typed cells with null blanks one by one and save.", () => ClosedXmlWriteCellValueObjectSparseBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign sparse object-typed cells with null blanks one by one and save.", () => EpPlusWriteCellValueObjectSparseBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-cellformula", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Assign numeric cells and row formulas one by one and save.", () => OfficeImoWriteCellFormulaBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Assign numeric cells and row formulas one by one and save.", () => ClosedXmlWriteCellFormulaBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Assign numeric cells and row formulas one by one and save.", () => EpPlusWriteCellFormulaBytes(rowCount))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-insertobjects-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert typed objects through the normal worksheet API and save.", () => OfficeImoWriteInsertObjectsBytes(rows)),
            new PackageProfileCase("ClosedXML", "Insert the same typed objects and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("EPPlus", "Import the same typed objects and save.", () => EpPlusWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming typed object export and save.", () => MiniExcelWriteSalesRowsBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Streaming typed object export and save.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-objects-direct-package", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write typed rows through the package-native OfficeIMO API.", () => OfficeImoWriteObjectsDirectPackageBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Stream the same typed rows through its package-native writer.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-typed-rows-compact-package", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write typed rows as a compact contiguous package through the package-native OfficeIMO row writer.", () => OfficeImoWriteTypedRowsCompactPackageBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Stream the same typed rows with cell references disabled.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: true, requireCellReferences: false))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-insertobjects-autofitcolumnsfor-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert typed objects, AutoFit all exported columns through AutoFitColumnsFor, and save.", () => OfficeImoWriteInsertObjectsAutoFitColumnsForBytes(rows)),
            new PackageProfileCase("ClosedXML", "Insert the same typed objects, adjust columns, and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable, autoFit: true)),
            new PackageProfileCase("EPPlus", "Import the same typed objects, autofit columns, and save.", () => EpPlusWriteDataTableBytes(salesDataTable, autoFit: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-insertobjects-partial-autofitcolumnsfor-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert typed objects, AutoFit selected exported columns through AutoFitColumnsFor, and save.", () => OfficeImoWriteInsertObjectsPartialAutoFitColumnsForBytes(rows)),
            new PackageProfileCase("ClosedXML", "Insert the same typed objects, adjust selected columns, and save.", () => ClosedXmlWriteDataTablePartialAutoFitBytes(salesDataTable, [1, 3, 6, 8])),
            new PackageProfileCase("EPPlus", "Import the same typed objects, autofit selected columns, and save.", () => EpPlusWriteDataTablePartialAutoFitBytes(salesDataTable, [1, 3, 6, 8]))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-insertobjects-flat-dictionaries-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert flat dictionary rows through the normal worksheet API and save.", () => OfficeImoWriteInsertDictionaryObjectsBytes(dictionaryRows)),
            new PackageProfileCase("ClosedXML", "Import the same prepared data and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("EPPlus", "Import the same prepared data and save.", () => EpPlusWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming typed row export with the same values and save.", () => MiniExcelWriteSalesRowsBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Streaming typed row export with the same values and save.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-insertobjects-flat-dictionaries-autofitcolumnsfor-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert flat dictionary rows, AutoFit all exported columns through AutoFitColumnsFor, and save.", () => OfficeImoWriteInsertDictionaryObjectsAutoFitColumnsForBytes(dictionaryRows)),
            new PackageProfileCase("ClosedXML", "Import the same prepared data, adjust columns, and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable, autoFit: true)),
            new PackageProfileCase("EPPlus", "Import the same prepared data, autofit columns, and save.", () => EpPlusWriteDataTableBytes(salesDataTable, autoFit: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-insertobjects-legacy-dictionaries-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert legacy dictionary/hashtable rows through the normal worksheet API and save.", () => OfficeImoWriteInsertDictionaryObjectsBytes(legacyDictionaryRows)),
            new PackageProfileCase("ClosedXML", "Import the same prepared data and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("EPPlus", "Import the same prepared data and save.", () => EpPlusWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming typed row export with the same values and save.", () => MiniExcelWriteSalesRowsBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Streaming typed row export with the same values and save.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-powershell-mixed-objects-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert PowerShell-like mixed dictionary objects through the normal worksheet API and save.", () => OfficeImoWriteInsertPowerShellMixedObjectsBytes(powerShellMixedRows)),
            new PackageProfileCase("ClosedXML", "Import the same mixed typed data and save.", () => ClosedXmlWriteDataTableBytes(powerShellMixedDataTable)),
            new PackageProfileCase("EPPlus", "Import the same mixed typed data and save.", () => EpPlusWriteDataTableBytes(powerShellMixedDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming mixed dictionary object export and save.", () => MiniExcelWriteDictionaryObjectsBytes(powerShellMixedRows)),
            new PackageProfileCase("LargeXlsx", "Streaming mixed dictionary rows and save.", () => LargeXlsxWritePowerShellMixedRowsBytes(powerShellMixedRows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-powershell-psobject-mixed-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert PSObject-like mixed rows through the normal worksheet API and save.", () => OfficeImoWriteInsertPowerShellObjectMixedObjectsBytes(powerShellObjectMixedRows)),
            new PackageProfileCase("ClosedXML", "Import the same mixed typed data and save.", () => ClosedXmlWriteDataTableBytes(powerShellMixedDataTable)),
            new PackageProfileCase("EPPlus", "Import the same mixed typed data and save.", () => EpPlusWriteDataTableBytes(powerShellMixedDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming mixed dictionary object export and save.", () => MiniExcelWriteDictionaryObjectsBytes(powerShellMixedRows)),
            new PackageProfileCase("LargeXlsx", "Streaming mixed dictionary rows and save.", () => LargeXlsxWritePowerShellMixedRowsBytes(powerShellMixedRows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-powershell-psobject-wide-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Insert PSObject-like wide telemetry rows through the normal worksheet API and save.", () => OfficeImoWriteInsertPowerShellObjectWideObjectsBytes(powerShellObjectWideRows)),
            new PackageProfileCase("ClosedXML", "Import the same wide typed data and save.", () => ClosedXmlWriteDataTableBytes(powerShellWideDataTable)),
            new PackageProfileCase("EPPlus", "Import the same wide typed data and save.", () => EpPlusWriteDataTableBytes(powerShellWideDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming wide dictionary object export and save.", () => MiniExcelWriteDictionaryObjectsBytes(powerShellWideRows)),
            new PackageProfileCase("LargeXlsx", "Streaming wide dictionary rows and save.", () => LargeXlsxWritePowerShellWideRowsBytes(powerShellWideRows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "write-fluent-rowsfrom-direct", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write typed rows through the fluent RowsFrom API and save.", () => OfficeImoWriteFluentRowsFromBytes(rows)),
            new PackageProfileCase("ClosedXML", "Insert the same typed rows and save.", () => ClosedXmlWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("EPPlus", "Import the same typed rows and save.", () => EpPlusWriteDataTableBytes(salesDataTable)),
            new PackageProfileCase("MiniExcel", "Streaming typed row export and save.", () => MiniExcelWriteSalesRowsBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Streaming typed row export and save.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: true))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "append-plain-rows", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Append prepared plain cells with CellValues parallel mode.", () => OfficeImoAppendPlainRowsBytes(rows)),
            new PackageProfileCase("ClosedXML", "Append equivalent row/cell values.", () => ClosedXmlAppendPlainRowsBytes(rows)),
            new PackageProfileCase("EPPlus", "Append equivalent row/cell values.", () => EpPlusAppendPlainRowsBytes(rows)),
            new PackageProfileCase("MiniExcel", "Streaming export of equivalent four-column row/cell values.", () => MiniExcelAppendPlainRowsBytes(rows)),
            new PackageProfileCase("LargeXlsx", "Streaming export of equivalent four-column row/cell values.", () => LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns: false))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, "autofit-existing", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Load existing workbook, autofit columns, save.", () => OfficeImoAutoFitExistingBytes(officeImoWorkbookBytes.Value)),
            new PackageProfileCase("ClosedXML", "Load existing workbook, autofit columns, save.", () => ClosedXmlAutoFitExistingBytes(closedXmlWorkbookBytes.Value)),
            new PackageProfileCase("EPPlus", "Load existing workbook, autofit columns, save.", () => EpPlusAutoFitExistingBytes(epPlusWorkbookBytes.Value))
        ]);

        AddReportWorkbookPackageProfileGroups(scenarios, scenarioFilter, powerShellMixedRows, powerShellMixedDataTable, warmupIterations, measuredIterations);
        AddRealWorldPackageProfileGroups(scenarios, scenarioFilter, rows, warmupIterations, measuredIterations);

        AddPackageProfileGroup(scenarios, scenarioFilter, "large-shared-strings", warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write repeated and distinct text-heavy cells.", () => OfficeImoWriteSharedStringsBytes(rowCount)),
            new PackageProfileCase("ClosedXML", "Write repeated and distinct text-heavy cells.", () => ClosedXmlWriteSharedStringsBytes(rowCount)),
            new PackageProfileCase("EPPlus", "Write repeated and distinct text-heavy cells.", () => EpPlusWriteSharedStringsBytes(rowCount)),
            new PackageProfileCase("MiniExcel", "Streaming export of repeated and distinct text-heavy cells.", () => MiniExcelWriteSharedStringsBytes(rowCount))
        ]);

        if (scenarios.Count == 0) {
            throw new ArgumentException("No package-profile scenarios matched the requested --scenario filter.");
        }

        var profile = new ExcelPackageProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            BuildConfiguration = BuildConfiguration,
            RowCount = rowCount,
            WarmupIterations = warmupIterations,
            MeasuredIterations = measuredIterations,
            Notes = "Local opt-in package-size profile. Timed samples include workbook generation only; ZIP part analysis is performed after timing.",
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

    private static ExcelLibraryComparisonProfile CreateComparisonProfile(
        int rowCount,
        int warmupIterations,
        int measuredIterations,
        string notes,
        List<ExcelLibraryComparisonScenario> scenarios) {
        return new ExcelLibraryComparisonProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            BuildConfiguration = BuildConfiguration,
            RowCount = rowCount,
            WarmupIterations = warmupIterations,
            MeasuredIterations = measuredIterations,
            Notes = notes,
            Scenarios = scenarios
        };
    }

    private static string WriteProfile(string outputPath, ExcelLibraryComparisonProfile profile) {
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
                filter.Add(NormalizeScenarioName(scenario.Trim()));
            }
        }

        return filter.Count == 0 ? null : filter;
    }

    private static HashSet<string>? BuildLibraryFilter(IReadOnlyCollection<string>? libraryFilters) {
        if (libraryFilters == null || libraryFilters.Count == 0) {
            return null;
        }

        var filter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string library in libraryFilters) {
            if (!string.IsNullOrWhiteSpace(library)) {
                filter.Add(library.Trim());
            }
        }

        return filter.Count == 0 ? null : filter;
    }

    private static string NormalizeScenarioName(string scenario)
        => string.Equals(scenario, LegacyMiniExcelHelloWorldReadRangeScenario, StringComparison.OrdinalIgnoreCase)
            ? DenseHelloWorldReadRangeScenario
            : string.Equals(scenario, LegacyMiniExcelHelloWorldReadStreamScenario, StringComparison.OrdinalIgnoreCase)
                ? DenseHelloWorldReadStreamScenario
                : scenario;

    private static bool ContainsOnlyHelloWorldScenarios(IReadOnlySet<string>? scenarioFilter)
        => scenarioFilter != null
           && scenarioFilter.Count > 0
           && scenarioFilter.All(IsHelloWorldScenario);

    private static bool IsHelloWorldScenario(string scenario)
        => string.Equals(scenario, DenseHelloWorldReadRangeScenario, StringComparison.OrdinalIgnoreCase)
           || string.Equals(scenario, DenseHelloWorldReadStreamScenario, StringComparison.OrdinalIgnoreCase);

    private static void AddHelloWorldScenarioGroups(
        List<ExcelLibraryComparisonScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        int rowCount,
        int warmupIterations,
        int measuredIterations) {
        if (scenarioFilter == null
            || (!scenarioFilter.Contains(DenseHelloWorldReadRangeScenario)
                && !scenarioFilter.Contains(DenseHelloWorldReadStreamScenario))) {
            return;
        }

        Console.WriteLine($"Preparing dense HelloWorld workbook with {rowCount.ToString(CultureInfo.InvariantCulture)} rows x {HelloWorldColumnCount.ToString(CultureInfo.InvariantCulture)} columns...");
        byte[] workbookBytes = CreateHelloWorldWorkbookBytes(rowCount);

        AddScenarioGroup(scenarios, scenarioFilter, DenseHelloWorldReadRangeScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Read A1:J(row count) HelloWorld grid with ReadRange.", () => OfficeImoReadHelloWorldRange(workbookBytes, rowCount)),
            new LibraryComparisonCase("MiniExcel", "Stream A1:J(row count) HelloWorld grid with QueryRange.", () => MiniExcelReadHelloWorldRange(workbookBytes, rowCount)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader scan of A1:J(row count) HelloWorld grid.", () => ExcelDataReaderReadHelloWorldRange(workbookBytes, rowCount)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader scan of A1:J(row count) HelloWorld grid.", () => SylvanReadHelloWorldRange(workbookBytes, rowCount))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, DenseHelloWorldReadStreamScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Stream A1:J(row count) HelloWorld grid with ReadRangeStream.", () => OfficeImoReadHelloWorldStream(workbookBytes, rowCount)),
            new LibraryComparisonCase("MiniExcel", "Stream A1:J(row count) HelloWorld grid with QueryRange.", () => MiniExcelReadHelloWorldRange(workbookBytes, rowCount)),
            new LibraryComparisonCase("ExcelDataReader", "Forward-only IExcelDataReader stream of A1:J(row count) HelloWorld grid.", () => ExcelDataReaderReadHelloWorldRange(workbookBytes, rowCount)),
            new LibraryComparisonCase("Sylvan.Data.Excel", "Forward-only DbDataReader stream of A1:J(row count) HelloWorld grid.", () => SylvanReadHelloWorldRange(workbookBytes, rowCount))
        ]);
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

        IReadOnlyList<LibraryComparisonCase> selectedCases = _libraryFilter == null
            ? cases
            : cases.Where(comparisonCase => _libraryFilter.Contains(comparisonCase.Library)).ToArray();
        if (selectedCases.Count == 0) {
            return;
        }

        Console.WriteLine($"Running {scenario} comparison group...");
        var measurements = BenchmarkMeasurement.MeasureGroup(
            warmupIterations,
            measuredIterations,
            selectedCases.Select(c => c.Action).ToArray());

        for (int i = 0; i < selectedCases.Count; i++) {
            var comparisonCase = selectedCases[i];
            var measurement = measurements[i];
            Console.WriteLine(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"{scenario} / {comparisonCase.Library}: avg {measurement.AverageMilliseconds:F2} ms, median {measurement.MedianMilliseconds:F2} ms, alloc {measurement.AverageAllocatedBytes / 1024.0:F1} KB"));

            scenarios.Add(new ExcelLibraryComparisonScenario {
                Scenario = scenario,
                Library = comparisonCase.Library,
                Notes = comparisonCase.Notes,
                OutputMetric = measurement.OutputMetric,
                AverageMilliseconds = measurement.AverageMilliseconds,
                MedianMilliseconds = measurement.MedianMilliseconds,
                StandardDeviationMilliseconds = measurement.StandardDeviationMilliseconds,
                StandardErrorMilliseconds = measurement.StandardErrorMilliseconds,
                SamplesMilliseconds = measurement.SamplesMilliseconds.ToList(),
                AverageAllocatedBytes = measurement.AverageAllocatedBytes,
                MedianAllocatedBytes = measurement.MedianAllocatedBytes,
                SamplesAllocatedBytes = measurement.SamplesAllocatedBytes.ToList()
            });
        }
    }

    private static void AddPackageProfileGroup(
        List<ExcelPackageProfileScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        string scenario,
        int warmupIterations,
        int measuredIterations,
        IReadOnlyList<PackageProfileCase> cases) {
        if (scenarioFilter != null && !scenarioFilter.Contains(scenario)) {
            return;
        }

        Console.WriteLine($"Running {scenario} package profile group...");
        var measurements = BenchmarkMeasurement.MeasureGroup(
            warmupIterations,
            measuredIterations,
            cases.Select(c => new Func<int>(() => checked((int)c.CreatePackage().LongLength))).ToArray());

        for (int i = 0; i < cases.Count; i++) {
            var packageCase = cases[i];
            var measurement = measurements[i];
            byte[] packageBytes = packageCase.CreatePackage();
            var profile = AnalyzePackage(packageBytes);
            ValidatePackageProfile(scenario, packageCase.Library, profile);

            Console.WriteLine(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"{scenario} / {packageCase.Library}: avg {measurement.AverageMilliseconds:F2} ms, median {measurement.MedianMilliseconds:F2} ms, alloc {measurement.AverageAllocatedBytes / 1024.0:F1} KB, package {packageBytes.LongLength:N0} bytes"));

            scenarios.Add(new ExcelPackageProfileScenario {
                Scenario = scenario,
                Library = packageCase.Library,
                Notes = packageCase.Notes,
                AverageMilliseconds = measurement.AverageMilliseconds,
                MedianMilliseconds = measurement.MedianMilliseconds,
                StandardDeviationMilliseconds = measurement.StandardDeviationMilliseconds,
                StandardErrorMilliseconds = measurement.StandardErrorMilliseconds,
                SamplesMilliseconds = measurement.SamplesMilliseconds.ToList(),
                AverageAllocatedBytes = measurement.AverageAllocatedBytes,
                MedianAllocatedBytes = measurement.MedianAllocatedBytes,
                SamplesAllocatedBytes = measurement.SamplesAllocatedBytes.ToList(),
                Package = profile
            });
        }
    }

    private static ExcelPackageProfileSummary AnalyzePackage(byte[] packageBytes) {
        using var stream = new MemoryStream(packageBytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var parts = new List<ExcelPackagePartProfile>(archive.Entries.Count);
        var summary = new ExcelPackageProfileSummary {
            FileSizeBytes = packageBytes.LongLength
        };

        foreach (var entry in archive.Entries.OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)) {
            if (string.IsNullOrEmpty(entry.Name)) {
                continue;
            }

            string category = GetPackagePartCategory(entry.FullName);
            parts.Add(new ExcelPackagePartProfile {
                Name = entry.FullName,
                Category = category,
                CompressedBytes = entry.CompressedLength,
                UncompressedBytes = entry.Length
            });

            AddCategorySize(summary, category, entry.CompressedLength, entry.Length);

            if (entry.FullName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)) {
                (summary.SharedStringCount, summary.UniqueSharedStringCount) = CountSharedStrings(entry);
            } else if (entry.FullName.Equals("xl/styles.xml", StringComparison.OrdinalIgnoreCase)) {
                summary.CellStyleCount = CountCellStyles(entry);
            } else if (IsWorksheetPart(entry.FullName)) {
                var worksheetStats = CountWorksheetCells(entry);
                summary.WorksheetRowCount += worksheetStats.Rows;
                summary.WorksheetCellCount += worksheetStats.Cells;
            }
        }

        summary.PartCount = parts.Count;
        summary.Parts = parts
            .OrderByDescending(part => part.UncompressedBytes)
            .ThenBy(part => part.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return summary;
    }

    private static void ValidatePackageProfile(string scenario, string library, ExcelPackageProfileSummary summary) {
        if (summary.FileSizeBytes <= 0) {
            throw new InvalidOperationException($"Scenario '{scenario}' for {library} produced an empty workbook package.");
        }

        if (summary.PartCount <= 0) {
            throw new InvalidOperationException($"Scenario '{scenario}' for {library} produced a workbook package without ZIP parts.");
        }

        if (summary.WorkbookCompressedBytes <= 0) {
            throw new InvalidOperationException($"Scenario '{scenario}' for {library} produced a workbook package without a workbook part.");
        }

        if (summary.WorksheetRowCount <= 0 || summary.WorksheetCellCount <= 0) {
            throw new InvalidOperationException(
                $"Scenario '{scenario}' for {library} produced a workbook package without worksheet rows and cells.");
        }
    }

    private static string GetPackagePartCategory(string partName) {
        if (IsWorksheetPart(partName)) return "Worksheets";
        if (partName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)) return "SharedStrings";
        if (partName.Equals("xl/styles.xml", StringComparison.OrdinalIgnoreCase)) return "Styles";
        if (partName.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase)) return "Tables";
        if (partName.StartsWith("xl/workbook", StringComparison.OrdinalIgnoreCase)) return "Workbook";
        if (partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) return "Relationships";
        if (partName.StartsWith("docProps/", StringComparison.OrdinalIgnoreCase)) return "DocProps";
        return "Other";
    }

    private static bool IsWorksheetPart(string partName)
        => partName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
           && partName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    private static void AddCategorySize(ExcelPackageProfileSummary summary, string category, long compressedBytes, long uncompressedBytes) {
        switch (category) {
            case "Worksheets":
                summary.WorksheetCompressedBytes += compressedBytes;
                summary.WorksheetUncompressedBytes += uncompressedBytes;
                break;
            case "SharedStrings":
                summary.SharedStringsCompressedBytes += compressedBytes;
                summary.SharedStringsUncompressedBytes += uncompressedBytes;
                break;
            case "Styles":
                summary.StylesCompressedBytes += compressedBytes;
                summary.StylesUncompressedBytes += uncompressedBytes;
                break;
            case "Tables":
                summary.TablesCompressedBytes += compressedBytes;
                summary.TablesUncompressedBytes += uncompressedBytes;
                break;
            case "Workbook":
                summary.WorkbookCompressedBytes += compressedBytes;
                summary.WorkbookUncompressedBytes += uncompressedBytes;
                break;
            case "Relationships":
                summary.RelationshipsCompressedBytes += compressedBytes;
                summary.RelationshipsUncompressedBytes += uncompressedBytes;
                break;
            case "DocProps":
                summary.DocPropsCompressedBytes += compressedBytes;
                summary.DocPropsUncompressedBytes += uncompressedBytes;
                break;
            default:
                summary.OtherCompressedBytes += compressedBytes;
                summary.OtherUncompressedBytes += uncompressedBytes;
                break;
        }
    }

    private static (int Count, int UniqueCount) CountSharedStrings(ZipArchiveEntry entry) {
        var unique = new HashSet<string>(StringComparer.Ordinal);
        int count = 0;
        using Stream stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreComments = true, IgnoreWhitespace = true });

        while (reader.Read()) {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si") {
                continue;
            }

            count++;
            string text = reader.ReadInnerXml();
            unique.Add(text);
        }

        return (count, unique.Count);
    }

    private static int CountCellStyles(ZipArchiveEntry entry) {
        using Stream stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreComments = true, IgnoreWhitespace = true });

        while (reader.Read()) {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "cellXfs") {
                continue;
            }

            string? count = reader.GetAttribute("count");
            if (int.TryParse(count, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) {
                return parsed;
            }

            if (reader.IsEmptyElement) {
                return 0;
            }

            int styles = 0;
            int depth = reader.Depth;
            while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "cellXfs")) {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "xf") {
                    styles++;
                }
            }

            return styles;
        }

        return 0;
    }

    private static (int Rows, int Cells) CountWorksheetCells(ZipArchiveEntry entry) {
        int rows = 0;
        int cells = 0;
        using Stream stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreComments = true, IgnoreWhitespace = true });

        while (reader.Read()) {
            if (reader.NodeType != XmlNodeType.Element) {
                continue;
            }

            if (reader.LocalName == "row") {
                rows++;
            } else if (reader.LocalName == "c") {
                cells++;
            }
        }

        return (rows, cells);
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

    private static int ByteCount(byte[] bytes) => checked((int)bytes.LongLength);

    private static string RepeatedText(int row, int variant)
        => variant switch {
            0 => "Region " + (row % 8).ToString(CultureInfo.InvariantCulture),
            1 => "Owner " + (row % 16).ToString(CultureInfo.InvariantCulture),
            _ => "Status " + (row % 4).ToString(CultureInfo.InvariantCulture)
        };

    private static string DistinctText(int row, int variant)
        => variant switch {
            0 => "Invoice " + row.ToString(CultureInfo.InvariantCulture),
            1 => "Customer " + row.ToString(CultureInfo.InvariantCulture),
            _ => "Comment " + row.ToString(CultureInfo.InvariantCulture) + " " + new string((char)('A' + (row % 26)), 24)
        };

    private static int OfficeImoWriteBulkReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteBulkReportBytes(rows));

    private static byte[] OfficeImoWriteBulkReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.PopulateOfficeImoWorksheet(sheet, rows);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteBulkReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(ClosedXmlWriteBulkReportBytes(rows));

    private static byte[] ClosedXmlWriteBulkReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(worksheet, rows);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteBulkReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteBulkReportBytes(rows));

    private static byte[] EpPlusWriteBulkReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusWorksheet(worksheet, rows, includeTable: true, autoFit: true);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteDataSetTables(DataSet dataSet, bool autoFit = false, bool includeHeaders = true)
        => ByteCount(OfficeImoWriteDataSetTablesBytes(dataSet, autoFit, includeHeaders));

    private static byte[] OfficeImoWriteDataSetTablesBytes(DataSet dataSet, bool autoFit = false, bool includeHeaders = true) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.InsertDataSet(dataSet, includeHeaders: includeHeaders, autoFit: autoFit);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "DataSet comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteDataSetDirectExport(DataSet dataSet)
        => ByteCount(OfficeImoWriteDataSetDirectExportBytes(dataSet));

    private static byte[] OfficeImoWriteDataSetDirectExportBytes(DataSet dataSet) {
        using var stream = new MemoryStream();
        ExcelDocument.WriteDataSet(stream, dataSet);
        return stream.ToArray();
    }

    private static int OfficeImoWriteDataTable(DataTable dataTable)
        => ByteCount(OfficeImoWriteDataTableBytes(dataTable));

    private static byte[] OfficeImoWriteDataTableBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertDataTable(dataTable);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "DataTable comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteDataTableAsTable(DataTable dataTable)
        => ByteCount(OfficeImoWriteDataTableAsTableBytes(dataTable));

    private static byte[] OfficeImoWriteDataTableAsTableBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertDataTableAsTable(dataTable, tableName: "SalesData", style: TableStyle.TableStyleMedium2);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "DataTable table comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteDataReaderTable(DataTable dataTable, bool autoFit = false)
        => ByteCount(OfficeImoWriteDataReaderTableBytes(dataTable, autoFit));

    private static byte[] OfficeImoWriteDataReaderTableBytes(DataTable dataTable, bool autoFit = false) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream))
        using (var reader = dataTable.CreateDataReader()) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertDataReader(reader, tableName: "SalesData", style: TableStyle.TableStyleMedium2, autoFit: autoFit);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteDataReaderPlain(DataTable dataTable)
        => ByteCount(OfficeImoWriteDataReaderPlainBytes(dataTable));

    private static byte[] OfficeImoWriteDataReaderPlainBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream))
        using (var reader = dataTable.CreateDataReader()) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertDataReader(reader, createTable: false);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteDataReaderDirectPackage(DataTable dataTable)
        => ByteCount(OfficeImoWriteDataReaderDirectPackageBytes(dataTable));

    private static byte[] OfficeImoWriteDataReaderDirectPackageBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using var reader = dataTable.CreateDataReader();
        ExcelDocument.WriteDataReader(stream, reader);
        return stream.ToArray();
    }

    private static int OfficeImoWriteDataReaderCompactPackage(DataTable dataTable)
        => ByteCount(OfficeImoWriteDataReaderCompactPackageBytes(dataTable));

    private static byte[] OfficeImoWriteDataReaderCompactPackageBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using var reader = dataTable.CreateDataReader();
        ExcelDocument.WriteDataReader(stream, reader, CompactTabularWriteOptions);
        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValuesRectangle(IReadOnlyList<(int Row, int Column, object Value)> cells)
        => ByteCount(OfficeImoWriteCellValuesRectangleBytes(cells));

    private static byte[] OfficeImoWriteCellValuesRectangleBytes(IReadOnlyList<(int Row, int Column, object Value)> cells) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValues(cells);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "CellValues rectangle comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValuesSparseRectangle(int rowCount)
        => ByteCount(OfficeImoWriteCellValuesSparseRectangleBytes(rowCount));

    private static byte[] OfficeImoWriteCellValuesSparseRectangleBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("SparseObjects");
            sheet.CellValues(BuildSparseObjectCells(rowCount), ExecutionMode.Parallel);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "CellValues sparse rectangle comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValuesHeaderlessRectangle(int rowCount)
        => ByteCount(OfficeImoWriteCellValuesHeaderlessRectangleBytes(rowCount));

    private static byte[] OfficeImoWriteCellValuesHeaderlessRectangleBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Headerless");
            sheet.CellValues(BuildHeaderlessMixedCells(rowCount));
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "CellValues headerless rectangle comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteBlogStringRows(IReadOnlyList<BlogStringRow> rows)
        => ByteCount(OfficeImoWriteBlogStringRowsBytes(rows));

    private static byte[] OfficeImoWriteBlogStringRowsBytes(IReadOnlyList<BlogStringRow> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertObjects(rows);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "2023 blog-style 20 string column comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertObjects(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteInsertObjectsBytes(rows));

    private static byte[] OfficeImoWriteInsertObjectsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteObjectsDirectPackage(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteObjectsDirectPackageBytes(rows));

    private static byte[] OfficeImoWriteObjectsDirectPackageBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        ExcelDocument.WriteObjects(stream, rows, ExcelBenchmarkScenarioFactory.SalesTypedColumns);
        return stream.ToArray();
    }

    private static int OfficeImoWriteTypedRowsCompactPackage(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteTypedRowsCompactPackageBytes(rows));

    private static byte[] OfficeImoWriteTypedRowsCompactPackageBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        ExcelDocument.WriteRows(
            stream,
            rows,
            ExcelBenchmarkScenarioFactory.SalesColumnNames,
            static (writer, row) => writer
                .Write(row.Id)
                .Write(row.Region)
                .Write(row.Owner)
                .Write(row.CreatedOn)
                .Write(row.Amount)
                .Write(row.Units)
                .Write(row.Active)
                .Write(row.Notes),
            CompactTabularWriteOptions);
        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertObjectsAutoFitColumnsFor(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteInsertObjectsAutoFitColumnsForBytes(rows));

    private static byte[] OfficeImoWriteInsertObjectsAutoFitColumnsForBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            sheet.AutoFitColumnsFor(Enumerable.Range(1, 8));
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects AutoFitColumnsFor comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertObjectsPartialAutoFitColumnsFor(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteInsertObjectsPartialAutoFitColumnsForBytes(rows));

    private static byte[] OfficeImoWriteInsertObjectsPartialAutoFitColumnsForBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            sheet.AutoFitColumnsFor([1, 3, 6, 8]);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects partial AutoFitColumnsFor comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertDictionaryObjects(IReadOnlyList<object?> rows)
        => ByteCount(OfficeImoWriteInsertDictionaryObjectsBytes(rows));

    private static byte[] OfficeImoWriteInsertDictionaryObjectsBytes(IReadOnlyList<object?> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertObjects(rows);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects flat dictionary comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertDictionaryObjectsAutoFitColumnsFor(IReadOnlyList<object?> rows)
        => ByteCount(OfficeImoWriteInsertDictionaryObjectsAutoFitColumnsForBytes(rows));

    private static byte[] OfficeImoWriteInsertDictionaryObjectsAutoFitColumnsForBytes(IReadOnlyList<object?> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertObjects(rows);
            sheet.AutoFitColumnsFor(Enumerable.Range(1, 8));
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects flat dictionary AutoFitColumnsFor comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertPowerShellMixedObjects(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(OfficeImoWriteInsertPowerShellMixedObjectsBytes(rows));

    private static byte[] OfficeImoWriteInsertPowerShellMixedObjectsBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertObjects(rows);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects PowerShell mixed comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertPowerShellObjectMixedObjects(IReadOnlyList<System.Management.Automation.PSObject> rows)
        => ByteCount(OfficeImoWriteInsertPowerShellObjectMixedObjectsBytes(rows));

    private static byte[] OfficeImoWriteInsertPowerShellObjectMixedObjectsBytes(IReadOnlyList<System.Management.Automation.PSObject> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertObjects(rows);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects PSObject mixed comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteInsertPowerShellObjectWideObjects(IReadOnlyList<System.Management.Automation.PSObject> rows)
        => ByteCount(OfficeImoWriteInsertPowerShellObjectWideObjectsBytes(rows));

    private static byte[] OfficeImoWriteInsertPowerShellObjectWideObjectsBytes(IReadOnlyList<System.Management.Automation.PSObject> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.InsertObjects(rows);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "InsertObjects PSObject wide comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteFluentRowsFrom(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteFluentRowsFromBytes(rows));

    private static byte[] OfficeImoWriteFluentRowsFromBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.AsFluent()
                .Sheet("Data", sheet => sheet.RowsFrom(rows))
                .End()
                .Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "Fluent RowsFrom comparison");
        }

        return stream.ToArray();
    }

    private static void AssertOfficeImoDirectPackageWriter(ExcelDocument document, string scenario) {
        if (document.LastSaveDiagnostics.Writer != ExcelSavePackageWriter.DirectDataSetPackage) {
            throw new InvalidOperationException("OfficeIMO " + scenario + " did not use the direct DataSet package writer: " + document.LastSaveDiagnostics.FastPackageSkipReason);
        }
    }

    private static int ClosedXmlWriteDataSetTables(DataSet dataSet, bool autoFit = false, bool includeHeaders = true)
        => ByteCount(ClosedXmlWriteDataSetTablesBytes(dataSet, autoFit, includeHeaders));

    private static byte[] ClosedXmlWriteDataSetTablesBytes(DataSet dataSet, bool autoFit = false, bool includeHeaders = true) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            foreach (DataTable dataTable in dataSet.Tables) {
                var worksheet = workbook.Worksheets.Add(dataTable.TableName);
                var table = worksheet.Cell(1, 1).InsertTable(dataTable, dataTable.TableName, true);
                if (!includeHeaders) {
                    table.ShowHeaderRow = false;
                }

                ExcelBenchmarkScenarioFactory.StyleClosedXmlTable(table);
                if (autoFit) {
                    worksheet.Columns().AdjustToContents();
                }
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteDataTable(DataTable dataTable, bool includeTable = false, bool autoFit = false)
        => ByteCount(ClosedXmlWriteDataTableBytes(dataTable, includeTable, autoFit));

    private static byte[] ClosedXmlWriteDataTableBytes(DataTable dataTable, bool includeTable = false, bool autoFit = false) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            if (includeTable) {
                var table = worksheet.Cell(1, 1).InsertTable(dataTable, "SalesData", true);
                ExcelBenchmarkScenarioFactory.StyleClosedXmlTable(table);
            } else {
                for (int i = 0; i < dataTable.Columns.Count; i++) {
                    worksheet.Cell(1, i + 1).Value = dataTable.Columns[i].ColumnName;
                }

                worksheet.Cell(2, 1).InsertData(dataTable.Rows.Cast<DataRow>().Select(row => dataTable.Columns.Cast<DataColumn>().Select(column => row[column]).ToArray()));
            }

            if (autoFit) {
                worksheet.ColumnsUsed().AdjustToContents();
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteDataTablePartialAutoFit(DataTable dataTable, IReadOnlyList<int> columnIndexes)
        => ByteCount(ClosedXmlWriteDataTablePartialAutoFitBytes(dataTable, columnIndexes));

    private static byte[] ClosedXmlWriteDataTablePartialAutoFitBytes(DataTable dataTable, IReadOnlyList<int> columnIndexes) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            for (int i = 0; i < dataTable.Columns.Count; i++) {
                worksheet.Cell(1, i + 1).Value = dataTable.Columns[i].ColumnName;
            }

            worksheet.Cell(2, 1).InsertData(dataTable.Rows.Cast<DataRow>().Select(row => dataTable.Columns.Cast<DataColumn>().Select(column => row[column]).ToArray()));
            for (int i = 0; i < columnIndexes.Count; i++) {
                int columnIndex = columnIndexes[i];
                if (columnIndex > 0 && columnIndex <= dataTable.Columns.Count) {
                    worksheet.Column(columnIndex).AdjustToContents();
                }
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteDataSetTables(DataSet dataSet, bool autoFit = false, bool includeHeaders = true)
        => ByteCount(EpPlusWriteDataSetTablesBytes(dataSet, autoFit, includeHeaders));

    private static byte[] EpPlusWriteDataSetTablesBytes(DataSet dataSet, bool autoFit = false, bool includeHeaders = true) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            foreach (DataTable dataTable in dataSet.Tables) {
                var worksheet = package.Workbook.Worksheets.Add(dataTable.TableName);
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, includeHeaders, TableStyles.Medium2);
                if (autoFit && worksheet.Dimension != null) {
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteDataTable(DataTable dataTable, bool includeTable = false, bool autoFit = false)
        => ByteCount(EpPlusWriteDataTableBytes(dataTable, includeTable, autoFit));

    private static byte[] EpPlusWriteDataTableBytes(DataTable dataTable, bool includeTable = false, bool autoFit = false) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true, includeTable ? TableStyles.Medium2 : TableStyles.None);
            if (autoFit && worksheet.Dimension != null) {
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteDataTablePartialAutoFit(DataTable dataTable, IReadOnlyList<int> columnIndexes)
        => ByteCount(EpPlusWriteDataTablePartialAutoFitBytes(dataTable, columnIndexes));

    private static byte[] EpPlusWriteDataTablePartialAutoFitBytes(DataTable dataTable, IReadOnlyList<int> columnIndexes) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true, TableStyles.None);
            int lastRow = dataTable.Rows.Count + 1;
            for (int i = 0; i < columnIndexes.Count; i++) {
                int columnIndex = columnIndexes[i];
                if (columnIndex > 0 && columnIndex <= dataTable.Columns.Count) {
                    worksheet.Cells[1, columnIndex, lastRow, columnIndex].AutoFitColumns();
                }
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int MiniExcelWriteBulkReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(MiniExcelWriteBulkReportBytes(rows));

    private static byte[] MiniExcelWriteBulkReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        MiniExcelApi.SaveAs(
            stream,
            rows,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX,
            configuration: CreateMiniExcelConfiguration(includeTable: true, autoFit: true));
        return stream.ToArray();
    }

    private static int MiniExcelWriteDataSetTables(DataSet dataSet, bool autoFit = false, bool includeHeaders = true)
        => ByteCount(MiniExcelWriteDataSetTablesBytes(dataSet, autoFit, includeHeaders));

    private static byte[] MiniExcelWriteDataSetTablesBytes(DataSet dataSet, bool autoFit = false, bool includeHeaders = true) {
        using var stream = new MemoryStream();
        MiniExcelApi.SaveAs(
            stream,
            dataSet,
            printHeader: includeHeaders,
            excelType: MiniExcelLibs.ExcelType.XLSX,
            configuration: CreateMiniExcelConfiguration(includeTable: true, autoFit: autoFit));
        return stream.ToArray();
    }

    private static int MiniExcelWriteDataTable(DataTable dataTable, bool includeTable = false)
        => ByteCount(MiniExcelWriteDataTableBytes(dataTable, includeTable));

    private static byte[] MiniExcelWriteDataTableBytes(DataTable dataTable, bool includeTable = false) {
        using var stream = new MemoryStream();
        MiniExcelApi.SaveAs(
            stream,
            dataTable,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX,
            configuration: CreateMiniExcelConfiguration(includeTable: includeTable));
        return stream.ToArray();
    }

    private static int MiniExcelWriteDataReaderTable(DataTable dataTable, bool autoFit = false)
        => ByteCount(MiniExcelWriteDataReaderTableBytes(dataTable, autoFit));

    private static byte[] MiniExcelWriteDataReaderTableBytes(DataTable dataTable, bool autoFit = false) {
        return MiniExcelWriteDataReaderBytes(dataTable, includeTable: true, autoFit: autoFit);
    }

    private static int MiniExcelWriteDataReaderPlain(DataTable dataTable)
        => ByteCount(MiniExcelWriteDataReaderPlainBytes(dataTable));

    private static byte[] MiniExcelWriteDataReaderPlainBytes(DataTable dataTable) {
        return MiniExcelWriteDataReaderBytes(dataTable, includeTable: false, autoFit: false);
    }

    private static byte[] MiniExcelWriteDataReaderBytes(DataTable dataTable, bool includeTable, bool autoFit) {
        using var stream = new MemoryStream();
        using var reader = dataTable.CreateDataReader();
        MiniExcelApi.SaveAs(
            stream,
            reader,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX,
            configuration: CreateMiniExcelConfiguration(includeTable: includeTable, autoFit: autoFit));
        return stream.ToArray();
    }

    private static int SylvanWriteDataReaderPlain(DataTable dataTable)
        => ByteCount(SylvanWriteDataReaderPlainBytes(dataTable));

    private static byte[] SylvanWriteDataReaderPlainBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var writer = SylvanExcelDataWriter.Create(stream, ExcelWorkbookType.ExcelXml, new ExcelDataWriterOptions {
            OwnsStream = false
        }))
        using (var reader = dataTable.CreateDataReader()) {
            writer.Write(reader, "Data");
        }

        return stream.ToArray();
    }

    private static int LargeXlsxWriteDataSetPlain(DataSet dataSet)
        => ByteCount(LargeXlsxWriteDataSetPlainBytes(dataSet));

    private static byte[] LargeXlsxWriteDataSetPlainBytes(DataSet dataSet) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream)) {
            foreach (DataTable dataTable in dataSet.Tables) {
                WriteLargeXlsxDataTable(writer, dataTable.TableName, dataTable);
            }
        }

        return stream.ToArray();
    }

    private static int LargeXlsxWriteDataTable(DataTable dataTable)
        => ByteCount(LargeXlsxWriteDataTableBytes(dataTable));

    private static byte[] LargeXlsxWriteDataTableBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream)) {
            WriteLargeXlsxDataTable(writer, "Data", dataTable);
        }

        return stream.ToArray();
    }

    private static int LargeXlsxWriteDataReaderPlain(DataTable dataTable)
        => ByteCount(LargeXlsxWriteDataReaderPlainBytes(dataTable));

    private static int LargeXlsxWriteDataReaderPlainCompact(DataTable dataTable)
        => ByteCount(LargeXlsxWriteDataReaderPlainBytes(dataTable, requireCellReferences: false));

    private static byte[] LargeXlsxWriteDataReaderPlainBytes(DataTable dataTable, bool requireCellReferences = true) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream, requireCellReferences: requireCellReferences))
        using (var reader = dataTable.CreateDataReader()) {
            writer.BeginWorksheet("Data");
            writer.BeginRow();
            for (int field = 0; field < reader.FieldCount; field++) {
                writer.Write(reader.GetName(field));
            }

            while (reader.Read()) {
                writer.BeginRow();
                for (int field = 0; field < reader.FieldCount; field++) {
                    WriteLargeXlsxValue(writer, reader.GetValue(field));
                }
            }
        }

        return stream.ToArray();
    }

    private static int MiniExcelWriteSalesRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(MiniExcelWriteSalesRowsBytes(rows));

    private static byte[] MiniExcelWriteSalesRowsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        MiniExcelApi.SaveAs(stream, rows, sheetName: "Data", excelType: MiniExcelLibs.ExcelType.XLSX);
        return stream.ToArray();
    }

    private static int MiniExcelWriteDictionaryObjects(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(MiniExcelWriteDictionaryObjectsBytes(rows));

    private static byte[] MiniExcelWriteDictionaryObjectsBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        MiniExcelApi.SaveAs(stream, rows, sheetName: "Data", excelType: MiniExcelLibs.ExcelType.XLSX);
        return stream.ToArray();
    }

    private static int MiniExcelWriteBlogStringRows(IReadOnlyList<BlogStringRow> rows)
        => ByteCount(MiniExcelWriteBlogStringRowsBytes(rows));

    private static byte[] MiniExcelWriteBlogStringRowsBytes(IReadOnlyList<BlogStringRow> rows) {
        using var stream = new MemoryStream();
        MiniExcelApi.SaveAs(stream, rows, sheetName: "Data", excelType: MiniExcelLibs.ExcelType.XLSX);
        return stream.ToArray();
    }

    private static int LargeXlsxWriteSalesRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns)
        => ByteCount(LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns));

    private static int LargeXlsxWriteSalesRowsCompact(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns)
        => ByteCount(LargeXlsxWriteSalesRowsBytes(rows, includeAllColumns, requireCellReferences: false));

    private static byte[] LargeXlsxWriteSalesRowsBytes(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        bool includeAllColumns,
        bool requireCellReferences = true) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream, requireCellReferences: requireCellReferences)) {
            WriteLargeXlsxSalesRows(writer, rows, includeAllColumns);
        }

        return stream.ToArray();
    }

    private static int LargeXlsxWritePowerShellMixedRows(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(LargeXlsxWritePowerShellMixedRowsBytes(rows));

    private static byte[] LargeXlsxWritePowerShellMixedRowsBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream)) {
            WriteLargeXlsxPowerShellMixedRows(writer, rows);
        }

        return stream.ToArray();
    }

    private static int LargeXlsxWritePowerShellWideRows(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(LargeXlsxWritePowerShellWideRowsBytes(rows));

    private static byte[] LargeXlsxWritePowerShellWideRowsBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream)) {
            WriteLargeXlsxPowerShellWideRows(writer, rows);
        }

        return stream.ToArray();
    }

    private static int LargeXlsxWriteBlogStringRows(IReadOnlyList<BlogStringRow> rows)
        => ByteCount(LargeXlsxWriteBlogStringRowsBytes(rows));

    private static byte[] LargeXlsxWriteBlogStringRowsBytes(IReadOnlyList<BlogStringRow> rows) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream)) {
            WriteLargeXlsxBlogStringRows(writer, rows);
        }

        return stream.ToArray();
    }

    private static int LargeXlsxWriteHeaderlessMixedRows(int rowCount)
        => ByteCount(LargeXlsxWriteHeaderlessMixedRowsBytes(rowCount));

    private static byte[] LargeXlsxWriteHeaderlessMixedRowsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream)) {
            writer.BeginWorksheet("Headerless");
            for (int row = 1; row <= rowCount; row++) {
                writer.BeginRow()
                    .Write(row * 1.25d)
                    .Write(row % 2 == 0)
                    .Write("Item " + row.ToString(CultureInfo.InvariantCulture));
            }
        }

        return stream.ToArray();
    }

    private static int OfficeImoAppendPlainRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoAppendPlainRowsBytes(rows));

    private static byte[] OfficeImoAppendPlainRowsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
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

        return stream.ToArray();
    }

    private static int OfficeImoCopyWorksheetFromPackage(byte[] workbookBytes)
        => ByteCount(OfficeImoCopyWorksheetFromBytes(workbookBytes, ExcelWorksheetCopyMode.Package));

    private static int OfficeImoCopyWorksheetFromValues(byte[] workbookBytes)
        => ByteCount(OfficeImoCopyWorksheetFromBytes(workbookBytes, ExcelWorksheetCopyMode.Values));

    private static byte[] OfficeImoCopyWorksheetFromBytes(byte[] workbookBytes, ExcelWorksheetCopyMode copyMode) {
        using var sourceStream = new MemoryStream(workbookBytes, writable: false);
        using var sourceDocument = ExcelDocument.Load(sourceStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly });
        using var targetStream = new MemoryStream();
        using (var targetDocument = ExcelDocument.Create(targetStream)) {
            targetDocument.CopyWorksheetFrom(
                sourceDocument,
                "Data",
                "DataCopy",
                SheetNameValidationMode.Sanitize,
                new ExcelWorksheetCopyOptions { CopyMode = copyMode });
            targetDocument.Save(targetStream);
        }

        return targetStream.ToArray();
    }

    private static int ClosedXmlCopyWorksheet(byte[] workbookBytes)
        => ByteCount(ClosedXmlCopyWorksheetBytes(workbookBytes));

    private static byte[] ClosedXmlCopyWorksheetBytes(byte[] workbookBytes) {
        using var sourceStream = new MemoryStream(workbookBytes, writable: false);
        using var sourceWorkbook = new XLWorkbook(sourceStream);
        using var targetStream = new MemoryStream();
        using (var targetWorkbook = new XLWorkbook()) {
            sourceWorkbook.Worksheet("Data").CopyTo(targetWorkbook, "DataCopy");
            targetWorkbook.SaveAs(targetStream);
        }

        return targetStream.ToArray();
    }

    private static int EpPlusCopyWorksheet(byte[] workbookBytes)
        => ByteCount(EpPlusCopyWorksheetBytes(workbookBytes));

    private static byte[] EpPlusCopyWorksheetBytes(byte[] workbookBytes) {
        using var sourceStream = new MemoryStream(workbookBytes, writable: false);
        using var sourcePackage = new ExcelPackage(sourceStream);
        using var targetStream = new MemoryStream();
        using (var targetPackage = new ExcelPackage(targetStream)) {
            var sourceWorksheet = sourcePackage.Workbook.Worksheets["Data"];
            if (sourceWorksheet == null) {
                throw new InvalidOperationException("Source worksheet 'Data' was not found.");
            }

            targetPackage.Workbook.Worksheets.Add("DataCopy", sourceWorksheet);
            targetPackage.Save();
        }

        return targetStream.ToArray();
    }

    private static int ClosedXmlAppendPlainRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(ClosedXmlAppendPlainRowsBytes(rows));

    private static byte[] ClosedXmlAppendPlainRowsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
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

        return stream.ToArray();
    }

    private static int EpPlusAppendPlainRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusAppendPlainRowsBytes(rows));

    private static byte[] EpPlusAppendPlainRowsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
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

        return stream.ToArray();
    }

    private static int MiniExcelAppendPlainRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(MiniExcelAppendPlainRowsBytes(rows));

    private static byte[] MiniExcelAppendPlainRowsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        var values = rows.Select(row => new MiniExcelAppendRecord {
            Id = row.Id,
            Region = row.Region,
            Owner = row.Owner,
            Amount = row.Amount
        });
        MiniExcelApi.SaveAs(stream, values, sheetName: "Data", excelType: MiniExcelLibs.ExcelType.XLSX);
        return stream.ToArray();
    }

    private static int ClosedXmlWriteSalesRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns)
        => ByteCount(ClosedXmlWriteSalesRowsBytes(rows, includeAllColumns));

    private static byte[] ClosedXmlWriteSalesRowsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteSalesRows(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns)
        => ByteCount(EpPlusWriteSalesRowsBytes(rows, includeAllColumns));

    private static byte[] EpPlusWriteSalesRowsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoReadRange(byte[] workbookBytes, string dataRange, bool rangeIncludesHeader = true) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        object?[,] values = reader.GetSheet("Data").ReadRange(dataRange);
        int metric = AddSalesHeadersMetric(0);
        int firstDataRow = rangeIncludesHeader ? 1 : 0;
        for (int row = firstDataRow; row < values.GetLength(0); row++) {
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

    private static int OfficeImoReadUsedRange(byte[] workbookBytes) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        var sheet = reader.GetSheet("Data");
        object?[,] values = sheet.ReadUsedRange();
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

    private static int OfficeImoReadDataReader(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader(dataRange, schemaSampleRows: 0);
        int metric = AddSalesHeadersMetric(0);
        var values = new object[dataReader.FieldCount];

        while (dataReader.Read()) {
            dataReader.GetValues(values);
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

        return metric;
    }

    private static int OfficeImoReadDataReaderGetValues(byte[] workbookBytes, string dataRange)
        => OfficeImoReadDataReader(workbookBytes, dataRange);

    private static int OfficeImoReadDataReaderRowsOnly(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader(dataRange, schemaSampleRows: 0);
        int metric = AddSalesHeadersMetric(dataReader.FieldCount);
        int rowsRead = 0;
        while (dataReader.Read()) {
            rowsRead++;
        }

        return AddIntMetric(metric, rowsRead);
    }

    private static int OfficeImoReadDataReaderFirstColumn(byte[] workbookBytes, string dataRange, int rowCount) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader(dataRange, schemaSampleRows: 0);
        int metric = AddSalesHeadersMetric(dataReader.FieldCount);
        int rowsRead = 0;
        while (dataReader.Read()) {
            rowsRead++;
            metric = AddSalesIdDataMetric(metric, rowsRead, rowCount, dataReader.GetInt32(0));
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount.ToString(CultureInfo.InvariantCulture)} data rows, got {rowsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int OfficeImoReadDataReaderTyped(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader(dataRange, schemaSampleRows: 0);
        int metric = AddSalesHeadersMetric(0);

        while (dataReader.Read()) {
            metric = AddSalesRangeMetric(
                metric,
                dataReader.GetInt32(0),
                dataReader.GetString(1),
                dataReader.GetString(2),
                dataReader.GetDateTime(3),
                dataReader.GetDouble(4),
                dataReader.GetInt32(5),
                dataReader.GetBoolean(6),
                dataReader.GetString(7));
        }

        return metric;
    }

    private static int OfficeImoReadRangeDecimal(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes, new ExcelReadOptions { NumericAsDecimal = true });
        object?[,] values = reader.GetSheet("Data").ReadRange(dataRange);
        int metric = AddSalesHeadersMetric(0);
        for (int row = 1; row < values.GetLength(0); row++) {
            metric = AddSalesRangeDecimalMetric(
                metric,
                Convert.ToInt32(values[row, 0], CultureInfo.InvariantCulture),
                Convert.ToString(values[row, 1], CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(values[row, 2], CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(values[row, 3]),
                Convert.ToDecimal(values[row, 4], CultureInfo.InvariantCulture),
                Convert.ToInt32(values[row, 5], CultureInfo.InvariantCulture),
                Convert.ToBoolean(values[row, 6], CultureInfo.InvariantCulture),
                Convert.ToString(values[row, 7], CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int ClosedXmlReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue, int skipDataRows = 0) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int firstRow = 2 + skipDataRows;
        if (maxDataRows != int.MaxValue) {
            lastRow = Math.Min(lastRow, firstRow + maxDataRows - 1);
        }

        int metric = AddSalesHeadersMetric(0);

        for (int row = firstRow; row <= lastRow; row++) {
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

    private static int ClosedXmlReadRangeDecimal(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int metric = AddSalesHeadersMetric(0);

        for (int row = 2; row <= lastRow; row++) {
            metric = AddSalesRangeDecimalMetric(
                metric,
                worksheet.Cell(row, 1).GetValue<int>(),
                worksheet.Cell(row, 2).GetValue<string>(),
                worksheet.Cell(row, 3).GetValue<string>(),
                worksheet.Cell(row, 4).GetValue<DateTime>(),
                worksheet.Cell(row, 5).GetValue<decimal>(),
                worksheet.Cell(row, 6).GetValue<int>(),
                worksheet.Cell(row, 7).GetValue<bool>(),
                worksheet.Cell(row, 8).GetValue<string>());
        }

        return metric;
    }

    private static int EpPlusReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue, int skipDataRows = 0) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int lastRow = worksheet.Dimension?.End.Row ?? 0;
        int firstRow = 2 + skipDataRows;
        if (maxDataRows != int.MaxValue) {
            lastRow = Math.Min(lastRow, firstRow + maxDataRows - 1);
        }

        int metric = AddSalesHeadersMetric(0);

        for (int row = firstRow; row <= lastRow; row++) {
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

    private static int EpPlusReadRangeDecimal(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int lastRow = worksheet.Dimension?.End.Row ?? 0;
        int metric = AddSalesHeadersMetric(0);

        for (int row = 2; row <= lastRow; row++) {
            metric = AddSalesRangeDecimalMetric(
                metric,
                Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
                Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(worksheet.Cells[row, 4].Value),
                Convert.ToDecimal(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
                Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
                Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
                Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int OfficeImoEnumerateRange(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;
        foreach (var cell in reader.GetSheet("Data").EnumerateRange(dataRange)) {
            metric = AddSalesEnumeratedCellMetric(metric, cell.Row, cell.Column, cell.Value);
        }

        return metric;
    }

    private static int OfficeImoEnumerateCells(byte[] workbookBytes) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;
        foreach (var cell in reader.GetSheet("Data").EnumerateCells()) {
            metric = AddSalesEnumeratedCellMetric(metric, cell.Row, cell.Column, cell.Value);
        }

        return metric;
    }

    private static int OfficeImoEnumerateFirstColumn(byte[] workbookBytes, string dataRange, int rowCount) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int expectedRows = rowCount + 1;
        int metric = 0;
        int cellsRead = 0;
        foreach (var cell in reader.GetSheet("Data").EnumerateRange(dataRange)) {
            cellsRead++;
            metric = AddSalesIdColumnMetric(metric, cell.Row, rowCount, cell.Value);
        }

        if (cellsRead != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows.ToString(CultureInfo.InvariantCulture)} first-column cells, got {cellsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int ClosedXmlEnumerateRange(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int metric = 0;
        int lastRow = rowCount + 1;
        for (int row = 1; row <= lastRow; row++) {
            for (int column = 1; column <= 8; column++) {
                object? value = row == 1
                    ? worksheet.Cell(row, column).GetString()
                    : column switch {
                        1 or 6 => worksheet.Cell(row, column).GetValue<int>(),
                        4 => worksheet.Cell(row, column).GetValue<DateTime>(),
                        5 => worksheet.Cell(row, column).GetValue<double>(),
                        7 => worksheet.Cell(row, column).GetValue<bool>(),
                        _ => worksheet.Cell(row, column).GetValue<string>()
                    };
                metric = AddSalesEnumeratedCellMetric(metric, row, column, value);
            }
        }

        return metric;
    }

    private static int ClosedXmlEnumerateFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int metric = 0;
        int expectedRows = rowCount + 1;
        for (int row = 1; row <= expectedRows; row++) {
            object value = row == 1
                ? worksheet.Cell(row, 1).GetString()
                : worksheet.Cell(row, 1).GetValue<int>();
            metric = AddSalesIdColumnMetric(metric, row, rowCount, value);
        }

        return metric;
    }

    private static int EpPlusEnumerateRange(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int metric = 0;
        int lastRow = rowCount + 1;
        for (int row = 1; row <= lastRow; row++) {
            for (int column = 1; column <= 8; column++) {
                metric = AddSalesEnumeratedCellMetric(metric, row, column, worksheet.Cells[row, column].Value);
            }
        }

        return metric;
    }

    private static int EpPlusEnumerateFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int metric = 0;
        int expectedRows = rowCount + 1;
        for (int row = 1; row <= expectedRows; row++) {
            metric = AddSalesIdColumnMetric(metric, row, rowCount, worksheet.Cells[row, 1].Value);
        }

        return metric;
    }

    private static int MiniExcelReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue, int skipDataRows = 0) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        int metric = AddSalesHeadersMetric(0);
        IEnumerable<object> rows = MiniExcelApi.Query(
            stream,
            useHeaderRow: true,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX);
        if (skipDataRows > 0) {
            rows = rows.Skip(skipDataRows);
        }

        if (maxDataRows != int.MaxValue) {
            rows = rows.Take(maxDataRows);
        }

        foreach (var item in rows) {
            var row = (IDictionary<string, object?>)item;
            metric = AddSalesRangeMetric(
                metric,
                Convert.ToInt32(row["Id"], CultureInfo.InvariantCulture),
                Convert.ToString(row["Region"], CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(row["Owner"], CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(row["CreatedOn"]),
                Convert.ToDouble(row["Amount"], CultureInfo.InvariantCulture),
                Convert.ToInt32(row["Units"], CultureInfo.InvariantCulture),
                Convert.ToBoolean(row["Active"], CultureInfo.InvariantCulture),
                Convert.ToString(row["Notes"], CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int MiniExcelReadRangeDecimal(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        int metric = AddSalesHeadersMetric(0);
        IEnumerable<object> rows = MiniExcelApi.Query(
            stream,
            useHeaderRow: true,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX);

        foreach (var item in rows) {
            var row = (IDictionary<string, object?>)item;
            metric = AddSalesRangeDecimalMetric(
                metric,
                Convert.ToInt32(row["Id"], CultureInfo.InvariantCulture),
                Convert.ToString(row["Region"], CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(row["Owner"], CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(row["CreatedOn"]),
                Convert.ToDecimal(row["Amount"], CultureInfo.InvariantCulture),
                Convert.ToInt32(row["Units"], CultureInfo.InvariantCulture),
                Convert.ToBoolean(row["Active"], CultureInfo.InvariantCulture),
                Convert.ToString(row["Notes"], CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int SylvanReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue, int skipDataRows = 0) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(0);
        int rowsSkipped = 0;
        int rowsRead = 0;

        while (reader.Read()) {
            if (rowsSkipped < skipDataRows) {
                rowsSkipped++;
                continue;
            }

            if (rowsRead >= maxDataRows) {
                break;
            }

            rowsRead++;
            metric = AddSalesRangeMetric(
                metric,
                reader.GetInt32(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetDateTime(3),
                reader.GetDouble(4),
                reader.GetInt32(5),
                reader.GetBoolean(6),
                reader.GetString(7));
        }

        return metric;
    }

    private static int SylvanReadRangeGetValues(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(0);
        var values = new object[reader.FieldCount];

        while (reader.Read()) {
            reader.GetValues(values);
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

        return metric;
    }

    private static int SylvanReadRowsOnly(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(reader.FieldCount);
        int rowsRead = 0;
        while (reader.Read()) {
            rowsRead++;
        }

        return AddIntMetric(metric, rowsRead);
    }

    private static int SylvanReadRangeFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(reader.FieldCount);
        int rowsRead = 0;
        while (reader.Read()) {
            rowsRead++;
            metric = AddSalesIdDataMetric(metric, rowsRead, rowCount, reader.GetInt32(0));
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount.ToString(CultureInfo.InvariantCulture)} data rows, got {rowsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int SylvanReadRangeDecimal(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(0);

        while (reader.Read()) {
            metric = AddSalesRangeDecimalMetric(
                metric,
                reader.GetInt32(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetDateTime(3),
                Convert.ToDecimal(reader.GetValue(4), CultureInfo.InvariantCulture),
                reader.GetInt32(5),
                reader.GetBoolean(6),
                reader.GetString(7));
        }

        return metric;
    }

    private static int ExcelDataReaderReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue, int skipDataRows = 0) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(0);
        int rowsSkipped = 0;
        int rowsRead = 0;

        if (!reader.Read()) {
            return metric;
        }

        while (reader.Read()) {
            if (rowsSkipped < skipDataRows) {
                rowsSkipped++;
                continue;
            }

            if (rowsRead >= maxDataRows) {
                break;
            }

            rowsRead++;
            metric = AddSalesRangeMetric(
                metric,
                Convert.ToInt32(GetExcelDataReaderValue(reader, 0), CultureInfo.InvariantCulture),
                Convert.ToString(GetExcelDataReaderValue(reader, 1), CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(GetExcelDataReaderValue(reader, 2), CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(GetExcelDataReaderValue(reader, 3)),
                Convert.ToDouble(GetExcelDataReaderValue(reader, 4), CultureInfo.InvariantCulture),
                Convert.ToInt32(GetExcelDataReaderValue(reader, 5), CultureInfo.InvariantCulture),
                Convert.ToBoolean(GetExcelDataReaderValue(reader, 6), CultureInfo.InvariantCulture),
                Convert.ToString(GetExcelDataReaderValue(reader, 7), CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int ExcelDataReaderReadRangeGetValues(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(0);
        var values = new object[reader.FieldCount];

        if (!reader.Read()) {
            return metric;
        }

        while (reader.Read()) {
            reader.GetValues(values);
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

        return metric;
    }

    private static int ExcelDataReaderReadRowsOnly(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(reader.FieldCount);
        if (!reader.Read()) {
            return metric;
        }

        int rowsRead = 0;
        while (reader.Read()) {
            rowsRead++;
        }

        return AddIntMetric(metric, rowsRead);
    }

    private static int ExcelDataReaderReadRangeFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(reader.FieldCount);

        if (!reader.Read()) {
            return metric;
        }

        int rowsRead = 0;
        while (reader.Read()) {
            rowsRead++;
            int id = Convert.ToInt32(GetExcelDataReaderValue(reader, 0), CultureInfo.InvariantCulture);
            metric = AddSalesIdDataMetric(metric, rowsRead, rowCount, id);
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount.ToString(CultureInfo.InvariantCulture)} data rows, got {rowsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int ExcelDataReaderReadRangeTyped(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(0);

        if (!reader.Read()) {
            return metric;
        }

        while (reader.Read()) {
            metric = AddSalesRangeMetric(
                metric,
                checked((int)reader.GetDouble(0)),
                reader.GetString(1),
                reader.GetString(2),
                ReadDateCell(reader.GetValue(3)),
                reader.GetDouble(4),
                checked((int)reader.GetDouble(5)),
                reader.GetBoolean(6),
                reader.GetString(7));
        }

        return metric;
    }

    private static int ExcelDataReaderReadRangeDecimal(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = AddSalesHeadersMetric(0);

        if (!reader.Read()) {
            return metric;
        }

        while (reader.Read()) {
            metric = AddSalesRangeDecimalMetric(
                metric,
                Convert.ToInt32(GetExcelDataReaderValue(reader, 0), CultureInfo.InvariantCulture),
                Convert.ToString(GetExcelDataReaderValue(reader, 1), CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(GetExcelDataReaderValue(reader, 2), CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(GetExcelDataReaderValue(reader, 3)),
                Convert.ToDecimal(GetExcelDataReaderValue(reader, 4), CultureInfo.InvariantCulture),
                Convert.ToInt32(GetExcelDataReaderValue(reader, 5), CultureInfo.InvariantCulture),
                Convert.ToBoolean(GetExcelDataReaderValue(reader, 6), CultureInfo.InvariantCulture),
                Convert.ToString(GetExcelDataReaderValue(reader, 7), CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

    private static int OfficeImoReadFirstColumn(byte[] workbookBytes, string dataRange, int rowCount) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int expectedRows = rowCount + 1;
        int metric = 0;
        int rowsRead = 0;

        foreach (object? value in reader.GetSheet("Data").ReadColumn(dataRange)) {
            rowsRead++;
            metric = AddSalesIdColumnMetric(metric, rowsRead, rowCount, value);
        }

        if (rowsRead != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows.ToString(CultureInfo.InvariantCulture)} first-column rows, got {rowsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int ClosedXmlReadFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int metric = 0;
        int expectedRows = rowCount + 1;

        for (int row = 1; row <= expectedRows; row++) {
            object value = row == 1
                ? worksheet.Cell(row, 1).GetString()
                : worksheet.Cell(row, 1).GetValue<int>();
            metric = AddSalesIdColumnMetric(metric, row, rowCount, value);
        }

        return metric;
    }

    private static int EpPlusReadFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int metric = 0;
        int expectedRows = rowCount + 1;

        for (int row = 1; row <= expectedRows; row++) {
            metric = AddSalesIdColumnMetric(metric, row, rowCount, worksheet.Cells[row, 1].Value);
        }

        return metric;
    }

    private static int MiniExcelReadFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        int metric = 0;
        int rowsRead = 0;
        string lastCell = "A" + (rowCount + 1).ToString(CultureInfo.InvariantCulture);

        foreach (var item in MiniExcelApi.QueryRange(
            stream,
            useHeaderRow: false,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX,
            startCell: "A1",
            endCell: lastCell)) {
            rowsRead++;
            var row = (IDictionary<string, object?>)item;
            metric = AddSalesIdColumnMetric(metric, rowsRead, rowCount, row.TryGetValue("A", out object? value) ? value : null);
        }

        int expectedRows = rowCount + 1;
        if (rowsRead != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows.ToString(CultureInfo.InvariantCulture)} first-column rows, got {rowsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int SylvanReadFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream, ExcelSchema.NoHeaders);
        OpenSylvanWorksheet(reader, "Data");
        int metric = 0;
        int rowsRead = 0;

        while (reader.Read()) {
            rowsRead++;
            object? value = reader.IsDBNull(0) ? null : reader.GetValue(0);
            metric = AddSalesIdColumnMetric(metric, rowsRead, rowCount, value);
        }

        int expectedRows = rowCount + 1;
        if (rowsRead != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows.ToString(CultureInfo.InvariantCulture)} first-column rows, got {rowsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int ExcelDataReaderReadFirstColumn(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = 0;
        int rowsRead = 0;

        while (reader.Read()) {
            rowsRead++;
            object? value = GetExcelDataReaderValue(reader, 0);
            metric = AddSalesIdColumnMetric(metric, rowsRead, rowCount, value);
        }

        int expectedRows = rowCount + 1;
        if (rowsRead != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows.ToString(CultureInfo.InvariantCulture)} first-column rows, got {rowsRead.ToString(CultureInfo.InvariantCulture)}.");
        }

        return metric;
    }

    private static int OfficeImoReadRangeStream(byte[] workbookBytes, string dataRange, int chunkRows = 512, bool rangeIncludesHeader = true) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = rangeIncludesHeader ? 0 : AddSalesHeadersMetric(0);

        foreach (var chunk in reader.GetSheet("Data").ReadRangeStream(dataRange, chunkRows: chunkRows)) {
            for (int rowOffset = 0; rowOffset < chunk.RowCount; rowOffset++) {
                int absoluteRow = chunk.StartRow + rowOffset;
                object?[] values = chunk.Rows[rowOffset];
                if (rangeIncludesHeader && absoluteRow == 1) {
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

    private static int ClosedXmlReadRangeStream(byte[] workbookBytes, int maxDataRows = int.MaxValue, int skipDataRows = 0) => ClosedXmlReadRange(workbookBytes, maxDataRows, skipDataRows);

    private static int EpPlusReadRangeStream(byte[] workbookBytes, int maxDataRows = int.MaxValue, int skipDataRows = 0) => EpPlusReadRange(workbookBytes, maxDataRows, skipDataRows);

    private static int OfficeImoReadDataTable(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        DataTable table = reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true);
        return AddSalesDataTableMetric(table);
    }

    private static int ClosedXmlReadDataTable(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        DataTable table = CreateSalesDataTable();

        for (int row = 2; row <= lastRow; row++) {
            table.Rows.Add(
                worksheet.Cell(row, 1).GetValue<int>(),
                worksheet.Cell(row, 2).GetValue<string>(),
                worksheet.Cell(row, 3).GetValue<string>(),
                worksheet.Cell(row, 4).GetValue<DateTime>(),
                worksheet.Cell(row, 5).GetValue<double>(),
                worksheet.Cell(row, 6).GetValue<int>(),
                worksheet.Cell(row, 7).GetValue<bool>(),
                worksheet.Cell(row, 8).GetValue<string>());
        }

        return AddSalesDataTableMetric(table);
    }

    private static int EpPlusReadDataTable(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Data"];
        int lastRow = worksheet.Dimension?.End.Row ?? 0;
        DataTable table = CreateSalesDataTable();

        for (int row = 2; row <= lastRow; row++) {
            table.Rows.Add(
                Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
                Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(worksheet.Cells[row, 4].Value),
                Convert.ToDouble(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
                Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
                Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
                Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return AddSalesDataTableMetric(table);
    }

    private static int MiniExcelReadDataTable(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        DataTable table = MiniExcelApi.QueryAsDataTable(
            stream,
            useHeaderRow: true,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX);
        return AddSalesDataTableMetric(table);
    }

    private static int SylvanReadDataTable(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        DataTable table = CreateSalesDataTable();

        while (reader.Read()) {
            table.Rows.Add(
                reader.GetInt32(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetDateTime(3),
                reader.GetDouble(4),
                reader.GetInt32(5),
                reader.GetBoolean(6),
                reader.GetString(7));
        }

        return AddSalesDataTableMetric(table);
    }

    private static int ExcelDataReaderReadDataTable(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        DataTable table = CreateSalesDataTable();

        if (!reader.Read()) {
            return AddSalesDataTableMetric(table);
        }

        while (reader.Read()) {
            table.Rows.Add(
                Convert.ToInt32(GetExcelDataReaderValue(reader, 0), CultureInfo.InvariantCulture),
                Convert.ToString(GetExcelDataReaderValue(reader, 1), CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(GetExcelDataReaderValue(reader, 2), CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(GetExcelDataReaderValue(reader, 3)),
                Convert.ToDouble(GetExcelDataReaderValue(reader, 4), CultureInfo.InvariantCulture),
                Convert.ToInt32(GetExcelDataReaderValue(reader, 5), CultureInfo.InvariantCulture),
                Convert.ToBoolean(GetExcelDataReaderValue(reader, 6), CultureInfo.InvariantCulture),
                Convert.ToString(GetExcelDataReaderValue(reader, 7), CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return AddSalesDataTableMetric(table);
    }

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

    private static int MiniExcelReadSparseColumn(byte[] workbookBytes, int expectedRows) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        int metric = 0;
        int rowIndex = 0;

        foreach (var item in MiniExcelApi.QueryRange(
            stream,
            useHeaderRow: false,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX,
            startCell: "A1",
            endCell: "A" + expectedRows.ToString(CultureInfo.InvariantCulture))) {
            rowIndex++;
            var row = (IDictionary<string, object?>)item;
            metric = AddSparseMetric(metric, rowIndex, expectedRows, row.TryGetValue("A", out object? value) ? value : null);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse rows, got {rowIndex}.");
        }

        return metric;
    }

    private static int SylvanReadSparseColumn(byte[] workbookBytes, int expectedRows) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream, ExcelSchema.NoHeaders);
        OpenSylvanWorksheet(reader, "Data");
        int metric = 0;
        int rowIndex = 0;

        while (reader.Read()) {
            rowIndex++;
            object? value = reader.IsDBNull(0) ? null : reader.GetValue(0);
            metric = AddSparseMetric(metric, rowIndex, expectedRows, value);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse rows, got {rowIndex}.");
        }

        return metric;
    }

    private static int ExcelDataReaderReadSparseColumn(byte[] workbookBytes, int expectedRows) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = 0;
        int rowIndex = 0;

        while (reader.Read()) {
            int currentRow = reader.Depth + 1;
            while (rowIndex < currentRow - 1) {
                rowIndex++;
                metric = AddSparseMetric(metric, rowIndex, expectedRows, null);
            }

            rowIndex++;
            metric = AddSparseMetric(metric, rowIndex, expectedRows, GetExcelDataReaderValue(reader, 0));
        }

        while (rowIndex < expectedRows) {
            rowIndex++;
            metric = AddSparseMetric(metric, rowIndex, expectedRows, null);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse rows, got {rowIndex}.");
        }

        return metric;
    }

    private static int OfficeImoReadObjects(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;
        foreach (var row in reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange)) {
            metric = AddSalesRecordMetric(metric, row);
        }

        return metric;
    }

    private static int OfficeImoReadObjectsStream(byte[] workbookBytes, string dataRange) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int metric = 0;
        foreach (var row in reader.GetSheet("Data").ReadObjectsStream<ReadSalesRecord>(dataRange)) {
            metric = AddSalesRecordMetric(metric, row);
        }

        return metric;
    }

    private static int ClosedXmlReadObjects(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var records = new List<ReadSalesRecord>(Math.Max(0, lastRow - 1));

        for (int row = 2; row <= lastRow; row++) {
            records.Add(new ReadSalesRecord {
                Id = worksheet.Cell(row, 1).GetValue<int>(),
                Region = worksheet.Cell(row, 2).GetValue<string>(),
                Owner = worksheet.Cell(row, 3).GetValue<string>(),
                CreatedOn = worksheet.Cell(row, 4).GetValue<DateTime>(),
                Amount = worksheet.Cell(row, 5).GetValue<double>(),
                Units = worksheet.Cell(row, 6).GetValue<int>(),
                Active = worksheet.Cell(row, 7).GetValue<bool>(),
                Notes = worksheet.Cell(row, 8).GetValue<string>()
            });
        }

        int metric = 0;
        foreach (var record in records) {
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int ClosedXmlReadObjectsStream(byte[] workbookBytes) {
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
        var records = new List<ReadSalesRecord>(Math.Max(0, lastRow - 1));

        for (int row = 2; row <= lastRow; row++) {
            records.Add(new ReadSalesRecord {
                Id = Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
                Region = Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                Owner = Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
                CreatedOn = worksheet.Cells[row, 4].Value is DateTime date ? date : DateTime.FromOADate(Convert.ToDouble(worksheet.Cells[row, 4].Value, CultureInfo.InvariantCulture)),
                Amount = Convert.ToDouble(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
                Units = Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
                Active = Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
                Notes = Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty
            });
        }

        int metric = 0;
        foreach (var record in records) {
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int EpPlusReadObjectsStream(byte[] workbookBytes) {
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

    private static int MiniExcelReadObjects(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        var records = MiniExcelApi.Query<ReadSalesRecord>(
            stream,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX).ToList();

        int metric = 0;
        foreach (var row in records) {
            metric = AddSalesRecordMetric(metric, row);
        }

        return metric;
    }

    private static int MiniExcelReadObjectsStream(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        int metric = 0;
        foreach (var row in MiniExcelApi.Query<ReadSalesRecord>(
            stream,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX)) {
            metric = AddSalesRecordMetric(metric, row);
        }

        return metric;
    }

    private static int SylvanReadObjects(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        var records = new List<ReadSalesRecord>();

        while (reader.Read()) {
            records.Add(new ReadSalesRecord {
                Id = reader.GetInt32(0),
                Region = reader.GetString(1),
                Owner = reader.GetString(2),
                CreatedOn = reader.GetDateTime(3),
                Amount = reader.GetDouble(4),
                Units = reader.GetInt32(5),
                Active = reader.GetBoolean(6),
                Notes = reader.GetString(7)
            });
        }

        int metric = 0;
        foreach (var record in records) {
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int SylvanReadObjectsStream(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream);
        OpenSylvanWorksheet(reader, "Data");
        int metric = 0;

        while (reader.Read()) {
            var record = new ReadSalesRecord {
                Id = reader.GetInt32(0),
                Region = reader.GetString(1),
                Owner = reader.GetString(2),
                CreatedOn = reader.GetDateTime(3),
                Amount = reader.GetDouble(4),
                Units = reader.GetInt32(5),
                Active = reader.GetBoolean(6),
                Notes = reader.GetString(7)
            };
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int ExcelDataReaderReadObjects(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");

        if (!reader.Read()) {
            return 0;
        }

        var records = new List<ReadSalesRecord>();
        while (reader.Read()) {
            records.Add(new ReadSalesRecord {
                Id = Convert.ToInt32(GetExcelDataReaderValue(reader, 0), CultureInfo.InvariantCulture),
                Region = Convert.ToString(GetExcelDataReaderValue(reader, 1), CultureInfo.InvariantCulture) ?? string.Empty,
                Owner = Convert.ToString(GetExcelDataReaderValue(reader, 2), CultureInfo.InvariantCulture) ?? string.Empty,
                CreatedOn = ReadDateCell(GetExcelDataReaderValue(reader, 3)),
                Amount = Convert.ToDouble(GetExcelDataReaderValue(reader, 4), CultureInfo.InvariantCulture),
                Units = Convert.ToInt32(GetExcelDataReaderValue(reader, 5), CultureInfo.InvariantCulture),
                Active = Convert.ToBoolean(GetExcelDataReaderValue(reader, 6), CultureInfo.InvariantCulture),
                Notes = Convert.ToString(GetExcelDataReaderValue(reader, 7), CultureInfo.InvariantCulture) ?? string.Empty
            });
        }

        int metric = 0;
        foreach (var record in records) {
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int ExcelDataReaderReadObjectsStream(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");

        if (!reader.Read()) {
            return 0;
        }

        int metric = 0;
        while (reader.Read()) {
            var record = new ReadSalesRecord {
                Id = Convert.ToInt32(GetExcelDataReaderValue(reader, 0), CultureInfo.InvariantCulture),
                Region = Convert.ToString(GetExcelDataReaderValue(reader, 1), CultureInfo.InvariantCulture) ?? string.Empty,
                Owner = Convert.ToString(GetExcelDataReaderValue(reader, 2), CultureInfo.InvariantCulture) ?? string.Empty,
                CreatedOn = ReadDateCell(GetExcelDataReaderValue(reader, 3)),
                Amount = Convert.ToDouble(GetExcelDataReaderValue(reader, 4), CultureInfo.InvariantCulture),
                Units = Convert.ToInt32(GetExcelDataReaderValue(reader, 5), CultureInfo.InvariantCulture),
                Active = Convert.ToBoolean(GetExcelDataReaderValue(reader, 6), CultureInfo.InvariantCulture),
                Notes = Convert.ToString(GetExcelDataReaderValue(reader, 7), CultureInfo.InvariantCulture) ?? string.Empty
            };
            metric = AddSalesRecordMetric(metric, record);
        }

        return metric;
    }

    private static int OfficeImoAutoFitExisting(byte[] workbookBytes)
        => ByteCount(OfficeImoAutoFitExistingBytes(workbookBytes));

    private static byte[] OfficeImoAutoFitExistingBytes(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Load(input)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            document.GetSheet("Data").AutoFitColumns();
            document.Save(output);
        }

        return output.ToArray();
    }

    private static int ClosedXmlAutoFitExisting(byte[] workbookBytes)
        => ByteCount(ClosedXmlAutoFitExistingBytes(workbookBytes));

    private static byte[] ClosedXmlAutoFitExistingBytes(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using (var workbook = new XLWorkbook(input)) {
            workbook.Worksheet("Data").ColumnsUsed().AdjustToContents();
            workbook.SaveAs(output);
        }

        return output.ToArray();
    }

    private static int EpPlusAutoFitExisting(byte[] workbookBytes)
        => ByteCount(EpPlusAutoFitExistingBytes(workbookBytes));

    private static byte[] EpPlusAutoFitExistingBytes(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using (var package = new ExcelPackage(input)) {
            var worksheet = package.Workbook.Worksheets["Data"];
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            package.SaveAs(output);
        }

        return output.ToArray();
    }

    private static int OfficeImoWriteSharedStrings(int rowCount)
        => ByteCount(OfficeImoWriteSharedStringsBytes(rowCount));

    private static byte[] OfficeImoWriteSharedStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Strings");
            sheet.CellValues(BuildSharedStringCells(rowCount), ExecutionMode.Parallel);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "shared string comparison");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueStrings(int rowCount)
        => ByteCount(OfficeImoWriteCellValueStringsBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Strings");
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, "Repeated value " + (row % 12));
                sheet.CellValue(row, 2, "Distinct value " + row.ToString(CultureInfo.InvariantCulture));
                sheet.CellValue(row, 3, "Long segment " + new string((char)('A' + (row % 26)), 48));
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueRepeatedStrings(int rowCount)
        => ByteCount(OfficeImoWriteCellValueRepeatedStringsBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueRepeatedStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Repeated");
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, RepeatedText(row, 0));
                sheet.CellValue(row, 2, RepeatedText(row, 1));
                sheet.CellValue(row, 3, RepeatedText(row, 2));
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueDistinctStrings(int rowCount)
        => ByteCount(OfficeImoWriteCellValueDistinctStringsBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueDistinctStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Distinct");
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, DistinctText(row, 0));
                sheet.CellValue(row, 2, DistinctText(row, 1));
                sheet.CellValue(row, 3, DistinctText(row, 2));
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueEmptyStrings(int rowCount)
        => ByteCount(OfficeImoWriteCellValueEmptyStringsBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueEmptyStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("EmptyText");
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, string.Empty);
                sheet.CellValue(row, 2, row % 3 == 0 ? string.Empty : "Status " + (row % 8).ToString(CultureInfo.InvariantCulture));
                sheet.CellValue(row, 3, row % 5 == 0 ? string.Empty : "Note " + row.ToString(CultureInfo.InvariantCulture));
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueNumbers(int rowCount)
        => ByteCount(OfficeImoWriteCellValueNumbersBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueNumbersBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Numbers");
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, (double)row * 1.25d);
                sheet.CellValue(row, 2, (double)row + 0.5d);
                sheet.CellValue(row, 3, (double)(row % 17));
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueScalars(int rowCount)
        => ByteCount(OfficeImoWriteCellValueScalarsBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueScalarsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Scalars");
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, row * 10.75m);
                sheet.CellValue(row, 2, row % 2 == 0);
                sheet.CellValue(row, 3, row % 3 == 0);
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueTemporal(int rowCount)
        => ByteCount(OfficeImoWriteCellValueTemporalBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueTemporalBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Temporal");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, start.AddDays(row));
                sheet.CellValue(row, 2, TimeSpan.FromMinutes(row * 7));
                sheet.CellValue(row, 3, start.AddHours(row % 24));
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueObjectMixed(int rowCount)
        => ByteCount(OfficeImoWriteCellValueObjectMixedBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueObjectMixedBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Objects");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                object? name = "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                object? amount = (double)row * 1.25d;
                object? active = row % 2 == 0;
                object? created = start.AddDays(row);
                sheet.CellValue(row, 1, name);
                sheet.CellValue(row, 2, amount);
                sheet.CellValue(row, 3, active);
                sheet.CellValue(row, 4, created);
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueObjectSparse(int rowCount)
        => ByteCount(OfficeImoWriteCellValueObjectSparseBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueObjectSparseBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("SparseObjects");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                object? name = row % 3 == 0 ? null : "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                object? amount = row % 4 == 0 ? null : (double)row * 1.25d;
                object? active = row % 5 == 0 ? null : row % 2 == 0;
                object? created = row % 7 == 0 ? null : start.AddDays(row);
                sheet.CellValue(row, 1, name);
                sheet.CellValue(row, 2, amount);
                sheet.CellValue(row, 3, active);
                sheet.CellValue(row, 4, created);
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellValueObjectSparseBatch(int rowCount)
        => ByteCount(OfficeImoWriteCellValueObjectSparseBatchBytes(rowCount));

    private static byte[] OfficeImoWriteCellValueObjectSparseBatchBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("SparseObjects");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            sheet.Batch(s => {
                for (int row = 1; row <= rowCount; row++) {
                    object? name = row % 3 == 0 ? null : "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                    object? amount = row % 4 == 0 ? null : (double)row * 1.25d;
                    object? active = row % 5 == 0 ? null : row % 2 == 0;
                    object? created = row % 7 == 0 ? null : start.AddDays(row);
                    s.CellValue(row, 1, name);
                    s.CellValue(row, 2, amount);
                    s.CellValue(row, 3, active);
                    s.CellValue(row, 4, created);
                }
            });

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteCellFormula(int rowCount)
        => ByteCount(OfficeImoWriteCellFormulaBytes(rowCount));

    private static byte[] OfficeImoWriteCellFormulaBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Formulas");
            for (int row = 1; row <= rowCount; row++) {
                sheet.CellValue(row, 1, (double)row);
                sheet.CellValue(row, 2, (double)(row % 17));
                sheet.CellValue(row, 3, (double)(row % 29));
                sheet.CellFormula(row, 4, $"SUM(A{row}:C{row})");
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteSharedStrings(int rowCount)
        => ByteCount(ClosedXmlWriteSharedStringsBytes(rowCount));

    private static byte[] ClosedXmlWriteSharedStringsBytes(int rowCount) {
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

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueRepeatedStrings(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueRepeatedStringsBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueRepeatedStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Repeated");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = RepeatedText(row, 0);
                worksheet.Cell(row, 2).Value = RepeatedText(row, 1);
                worksheet.Cell(row, 3).Value = RepeatedText(row, 2);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueDistinctStrings(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueDistinctStringsBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueDistinctStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Distinct");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = DistinctText(row, 0);
                worksheet.Cell(row, 2).Value = DistinctText(row, 1);
                worksheet.Cell(row, 3).Value = DistinctText(row, 2);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueEmptyStrings(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueEmptyStringsBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueEmptyStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("EmptyText");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = string.Empty;
                worksheet.Cell(row, 2).Value = row % 3 == 0 ? string.Empty : "Status " + (row % 8).ToString(CultureInfo.InvariantCulture);
                worksheet.Cell(row, 3).Value = row % 5 == 0 ? string.Empty : "Note " + row.ToString(CultureInfo.InvariantCulture);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueNumbers(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueNumbersBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueNumbersBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Numbers");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = row * 1.25d;
                worksheet.Cell(row, 2).Value = row + 0.5d;
                worksheet.Cell(row, 3).Value = row % 17;
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueScalars(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueScalarsBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueScalarsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Scalars");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = row * 10.75m;
                worksheet.Cell(row, 2).Value = row % 2 == 0;
                worksheet.Cell(row, 3).Value = row % 3 == 0;
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueTemporal(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueTemporalBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueTemporalBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Temporal");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = start.AddDays(row);
                worksheet.Cell(row, 2).Value = TimeSpan.FromMinutes(row * 7);
                worksheet.Cell(row, 3).Value = start.AddHours(row % 24);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueObjectMixed(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueObjectMixedBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueObjectMixedBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Objects");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                object? name = "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                object? amount = (double)row * 1.25d;
                object? active = row % 2 == 0;
                object? created = start.AddDays(row);
                worksheet.Cell(row, 1).Value = XLCellValue.FromObject(name, CultureInfo.InvariantCulture);
                worksheet.Cell(row, 2).Value = XLCellValue.FromObject(amount, CultureInfo.InvariantCulture);
                worksheet.Cell(row, 3).Value = XLCellValue.FromObject(active, CultureInfo.InvariantCulture);
                worksheet.Cell(row, 4).Value = XLCellValue.FromObject(created, CultureInfo.InvariantCulture);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteHeaderlessMixedRows(int rowCount)
        => ByteCount(ClosedXmlWriteHeaderlessMixedRowsBytes(rowCount));

    private static byte[] ClosedXmlWriteHeaderlessMixedRowsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Headerless");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = row * 1.25d;
                worksheet.Cell(row, 2).Value = row % 2 == 0;
                worksheet.Cell(row, 3).Value = "Item " + row.ToString(CultureInfo.InvariantCulture);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteBlogStringRows(IReadOnlyList<BlogStringRow> rows)
        => ByteCount(ClosedXmlWriteBlogStringRowsBytes(rows));

    private static byte[] ClosedXmlWriteBlogStringRowsBytes(IReadOnlyList<BlogStringRow> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteBlogStringHeaders(worksheet);
            for (int i = 0; i < rows.Count; i++) {
                WriteBlogStringRow(worksheet, i + 2, rows[i]);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellValueObjectSparse(int rowCount)
        => ByteCount(ClosedXmlWriteCellValueObjectSparseBytes(rowCount));

    private static byte[] ClosedXmlWriteCellValueObjectSparseBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("SparseObjects");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                object? name = row % 3 == 0 ? null : "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                object? amount = row % 4 == 0 ? null : (double)row * 1.25d;
                object? active = row % 5 == 0 ? null : row % 2 == 0;
                object? created = row % 7 == 0 ? null : start.AddDays(row);
                worksheet.Cell(row, 1).Value = name == null ? Blank.Value : XLCellValue.FromObject(name, CultureInfo.InvariantCulture);
                worksheet.Cell(row, 2).Value = amount == null ? Blank.Value : XLCellValue.FromObject(amount, CultureInfo.InvariantCulture);
                worksheet.Cell(row, 3).Value = active == null ? Blank.Value : XLCellValue.FromObject(active, CultureInfo.InvariantCulture);
                worksheet.Cell(row, 4).Value = created == null ? Blank.Value : XLCellValue.FromObject(created, CultureInfo.InvariantCulture);
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteCellFormula(int rowCount)
        => ByteCount(ClosedXmlWriteCellFormulaBytes(rowCount));

    private static byte[] ClosedXmlWriteCellFormulaBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Formulas");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cell(row, 1).Value = (double)row;
                worksheet.Cell(row, 2).Value = (double)(row % 17);
                worksheet.Cell(row, 3).Value = (double)(row % 29);
                worksheet.Cell(row, 4).FormulaA1 = $"SUM(A{row}:C{row})";
            }

            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteSharedStrings(int rowCount)
        => ByteCount(EpPlusWriteSharedStringsBytes(rowCount));

    private static byte[] EpPlusWriteSharedStringsBytes(int rowCount) {
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

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueRepeatedStrings(int rowCount)
        => ByteCount(EpPlusWriteCellValueRepeatedStringsBytes(rowCount));

    private static byte[] EpPlusWriteCellValueRepeatedStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Repeated");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = RepeatedText(row, 0);
                worksheet.Cells[row, 2].Value = RepeatedText(row, 1);
                worksheet.Cells[row, 3].Value = RepeatedText(row, 2);
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueDistinctStrings(int rowCount)
        => ByteCount(EpPlusWriteCellValueDistinctStringsBytes(rowCount));

    private static byte[] EpPlusWriteCellValueDistinctStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Distinct");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = DistinctText(row, 0);
                worksheet.Cells[row, 2].Value = DistinctText(row, 1);
                worksheet.Cells[row, 3].Value = DistinctText(row, 2);
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueEmptyStrings(int rowCount)
        => ByteCount(EpPlusWriteCellValueEmptyStringsBytes(rowCount));

    private static byte[] EpPlusWriteCellValueEmptyStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("EmptyText");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = string.Empty;
                worksheet.Cells[row, 2].Value = row % 3 == 0 ? string.Empty : "Status " + (row % 8).ToString(CultureInfo.InvariantCulture);
                worksheet.Cells[row, 3].Value = row % 5 == 0 ? string.Empty : "Note " + row.ToString(CultureInfo.InvariantCulture);
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueNumbers(int rowCount)
        => ByteCount(EpPlusWriteCellValueNumbersBytes(rowCount));

    private static byte[] EpPlusWriteCellValueNumbersBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Numbers");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = row * 1.25d;
                worksheet.Cells[row, 2].Value = row + 0.5d;
                worksheet.Cells[row, 3].Value = row % 17;
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueScalars(int rowCount)
        => ByteCount(EpPlusWriteCellValueScalarsBytes(rowCount));

    private static byte[] EpPlusWriteCellValueScalarsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Scalars");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = row * 10.75m;
                worksheet.Cells[row, 2].Value = row % 2 == 0;
                worksheet.Cells[row, 3].Value = row % 3 == 0;
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueTemporal(int rowCount)
        => ByteCount(EpPlusWriteCellValueTemporalBytes(rowCount));

    private static byte[] EpPlusWriteCellValueTemporalBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Temporal");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = start.AddDays(row);
                worksheet.Cells[row, 2].Value = TimeSpan.FromMinutes(row * 7);
                worksheet.Cells[row, 3].Value = start.AddHours(row % 24);
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueObjectMixed(int rowCount)
        => ByteCount(EpPlusWriteCellValueObjectMixedBytes(rowCount));

    private static byte[] EpPlusWriteCellValueObjectMixedBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Objects");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                object? name = "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                object? amount = (double)row * 1.25d;
                object? active = row % 2 == 0;
                object? created = start.AddDays(row);
                worksheet.Cells[row, 1].Value = name;
                worksheet.Cells[row, 2].Value = amount;
                worksheet.Cells[row, 3].Value = active;
                worksheet.Cells[row, 4].Value = created;
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteHeaderlessMixedRows(int rowCount)
        => ByteCount(EpPlusWriteHeaderlessMixedRowsBytes(rowCount));

    private static byte[] EpPlusWriteHeaderlessMixedRowsBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Headerless");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = row * 1.25d;
                worksheet.Cells[row, 2].Value = row % 2 == 0;
                worksheet.Cells[row, 3].Value = "Item " + row.ToString(CultureInfo.InvariantCulture);
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellValueObjectSparse(int rowCount)
        => ByteCount(EpPlusWriteCellValueObjectSparseBytes(rowCount));

    private static byte[] EpPlusWriteCellValueObjectSparseBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("SparseObjects");
            var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
            for (int row = 1; row <= rowCount; row++) {
                object? name = row % 3 == 0 ? null : "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
                object? amount = row % 4 == 0 ? null : (double)row * 1.25d;
                object? active = row % 5 == 0 ? null : row % 2 == 0;
                object? created = row % 7 == 0 ? null : start.AddDays(row);
                worksheet.Cells[row, 1].Value = name;
                worksheet.Cells[row, 2].Value = amount;
                worksheet.Cells[row, 3].Value = active;
                worksheet.Cells[row, 4].Value = created;
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteCellFormula(int rowCount)
        => ByteCount(EpPlusWriteCellFormulaBytes(rowCount));

    private static byte[] EpPlusWriteCellFormulaBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Formulas");
            for (int row = 1; row <= rowCount; row++) {
                worksheet.Cells[row, 1].Value = (double)row;
                worksheet.Cells[row, 2].Value = (double)(row % 17);
                worksheet.Cells[row, 3].Value = (double)(row % 29);
                worksheet.Cells[row, 4].Formula = $"SUM(A{row}:C{row})";
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static int MiniExcelWriteSharedStrings(int rowCount)
        => ByteCount(MiniExcelWriteSharedStringsBytes(rowCount));

    private static byte[] MiniExcelWriteSharedStringsBytes(int rowCount) {
        using var stream = new MemoryStream();
        var rows = Enumerable.Range(1, rowCount).Select(row => new MiniExcelStringRecord {
            Repeated = "Repeated value " + (row % 12),
            Distinct = "Distinct value " + row.ToString(CultureInfo.InvariantCulture),
            LongSegment = "Long segment " + new string((char)('A' + (row % 26)), 48)
        });
        MiniExcelApi.SaveAs(stream, rows, sheetName: "Strings", excelType: MiniExcelLibs.ExcelType.XLSX);
        return stream.ToArray();
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

    private static int MiniExcelReadSharedStrings(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        int metric = 0;
        int rowsRead = 0;
        foreach (var item in MiniExcelApi.QueryRange(
            stream,
            useHeaderRow: false,
            sheetName: "Strings",
            excelType: MiniExcelLibs.ExcelType.XLSX,
            startCell: "A1",
            endCell: "C" + rowCount.ToString(CultureInfo.InvariantCulture))) {
            rowsRead++;
            var row = (IDictionary<string, object?>)item;
            metric = AddStringMetric(metric, Convert.ToString(row["A"], CultureInfo.InvariantCulture));
            metric = AddStringMetric(metric, Convert.ToString(row["B"], CultureInfo.InvariantCulture));
            metric = AddStringMetric(metric, Convert.ToString(row["C"], CultureInfo.InvariantCulture));
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} shared string rows, got {rowsRead}.");
        }

        return metric;
    }

    private static int SylvanReadSharedStrings(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream, ExcelSchema.NoHeaders);
        OpenSylvanWorksheet(reader, "Strings");
        int metric = 0;
        int rowsRead = 0;

        while (reader.Read()) {
            rowsRead++;
            for (int col = 0; col < 3; col++) {
                metric = AddStringMetric(metric, reader.GetString(col));
            }
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} shared string rows, got {rowsRead}.");
        }

        return metric;
    }

    private static int ExcelDataReaderReadSharedStrings(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Strings");
        int metric = 0;
        int rowsRead = 0;

        while (reader.Read()) {
            rowsRead++;
            for (int col = 0; col < 3; col++) {
                metric = AddStringMetric(metric, Convert.ToString(GetExcelDataReaderValue(reader, col), CultureInfo.InvariantCulture));
            }
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} shared string rows, got {rowsRead}.");
        }

        return metric;
    }

    private static int OfficeImoReadHelloWorldRange(byte[] workbookBytes, int rowCount) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        string dataRange = "A1:J" + rowCount.ToString(CultureInfo.InvariantCulture);
        object?[,] values = reader.GetSheet("Data").ReadRange(dataRange, OfficeIMO.Excel.ExecutionMode.Sequential);
        int metric = 0;
        for (int row = 0; row < values.GetLength(0); row++) {
            int rowIndex = row + 1;
            for (int column = 0; column < values.GetLength(1); column++) {
                metric = AddHelloWorldMetric(metric, values[row, column], rowIndex, HelloWorldColumnNames[column]);
            }
        }

        return metric;
    }

    private static int OfficeImoReadHelloWorldStream(byte[] workbookBytes, int rowCount) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        string dataRange = "A1:J" + rowCount.ToString(CultureInfo.InvariantCulture);
        int metric = 0;
        int rowsRead = 0;

        foreach (var chunk in reader.GetSheet("Data").ReadRangeStream(dataRange, chunkRows: 4096)) {
            for (int rowOffset = 0; rowOffset < chunk.RowCount; rowOffset++) {
                rowsRead++;
                object?[] values = chunk.Rows[rowOffset];
                for (int column = 0; column < HelloWorldColumnCount; column++) {
                    metric = AddHelloWorldMetric(metric, values[column], rowsRead, HelloWorldColumnNames[column]);
                }
            }
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} HelloWorld rows, got {rowsRead}.");
        }

        return metric;
    }

    private static int MiniExcelReadHelloWorldRange(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        int metric = 0;
        int rowsRead = 0;

        foreach (var item in MiniExcelApi.QueryRange(
            stream,
            useHeaderRow: false,
            sheetName: "Data",
            excelType: MiniExcelLibs.ExcelType.XLSX,
            startCell: "A1",
            endCell: "J" + rowCount.ToString(CultureInfo.InvariantCulture))) {
            rowsRead++;
            var row = (IDictionary<string, object?>)item;
            for (int column = 0; column < HelloWorldColumnNames.Length; column++) {
                string columnName = HelloWorldColumnNames[column];
                metric = AddHelloWorldMetric(metric, row[columnName], rowsRead, columnName);
            }
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} HelloWorld rows, got {rowsRead}.");
        }

        return metric;
    }

    private static int SylvanReadHelloWorldRange(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateSylvanReader(stream, ExcelSchema.NoHeaders);
        OpenSylvanWorksheet(reader, "Data");
        int metric = 0;
        int rowsRead = 0;

        while (reader.Read()) {
            rowsRead++;
            for (int column = 0; column < HelloWorldColumnCount; column++) {
                metric = AddHelloWorldMetric(metric, reader.GetString(column), rowsRead, HelloWorldColumnNames[column]);
            }
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} HelloWorld rows, got {rowsRead}.");
        }

        return metric;
    }

    private static int ExcelDataReaderReadHelloWorldRange(byte[] workbookBytes, int rowCount) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = CreateExcelDataReader(stream);
        OpenExcelDataReaderWorksheet(reader, "Data");
        int metric = 0;
        int rowsRead = 0;

        while (reader.Read()) {
            rowsRead++;
            for (int column = 0; column < HelloWorldColumnCount; column++) {
                metric = AddHelloWorldMetric(metric, GetExcelDataReaderValue(reader, column), rowsRead, HelloWorldColumnNames[column]);
            }
        }

        if (rowsRead != rowCount) {
            throw new InvalidOperationException($"Expected {rowCount} HelloWorld rows, got {rowsRead}.");
        }

        return metric;
    }

    private static global::ExcelDataReader.IExcelDataReader CreateExcelDataReader(Stream stream)
        => global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream, new global::ExcelDataReader.ExcelReaderConfiguration {
            LeaveOpen = false
        });

    private static void OpenExcelDataReaderWorksheet(global::ExcelDataReader.IExcelDataReader reader, string sheetName) {
        do {
            if (string.Equals(reader.Name, sheetName, StringComparison.Ordinal)) {
                return;
            }
        } while (reader.NextResult());

        throw new InvalidOperationException($"ExcelDataReader could not open worksheet '{sheetName}'.");
    }

    private static object? GetExcelDataReaderValue(global::ExcelDataReader.IExcelDataReader reader, int ordinal) {
        if (ordinal < 0 || ordinal >= reader.FieldCount || reader.IsDBNull(ordinal)) {
            return null;
        }

        return reader.GetValue(ordinal);
    }

    private static SylvanExcelDataReader CreateSylvanReader(Stream stream, IExcelSchemaProvider? schema = null)
        => SylvanExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions {
            Schema = schema ?? ExcelSchema.Default
        });

    private static void OpenSylvanWorksheet(SylvanExcelDataReader reader, string sheetName) {
        if (string.Equals(reader.WorksheetName, sheetName, StringComparison.Ordinal)) {
            return;
        }

        if (!reader.TryOpenWorksheet(sheetName)) {
            throw new InvalidOperationException($"Sylvan.Data.Excel could not open worksheet '{sheetName}'.");
        }
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

    private static byte[] CreateMiniExcelWorkbookBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        MiniExcelApi.SaveAs(stream, rows, sheetName: "Data", excelType: MiniExcelLibs.ExcelType.XLSX);
        return stream.ToArray();
    }

    private static byte[] CreateHelloWorldWorkbookBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValues(BuildHelloWorldCells(rowCount), ExecutionMode.Parallel);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, DenseHelloWorldReadStreamScenario + " fixture generation");
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
            document.Save(stream);
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

    private static (int Row, int Column, object Value)[] BuildHelloWorldCells(int rowCount) {
        var cells = new (int Row, int Column, object Value)[checked(rowCount * HelloWorldColumnCount)];
        int offset = 0;
        for (int row = 1; row <= rowCount; row++) {
            for (int column = 1; column <= HelloWorldColumnCount; column++) {
                cells[offset++] = (row, column, HelloWorldValue);
            }
        }

        return cells;
    }

    private static (int Row, int Column, object Value)[] BuildSalesCells(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        var cells = new (int Row, int Column, object Value)[(rows.Count + 1) * 8];
        int offset = 0;
        AddSalesHeaderCells(cells, ref offset);
        for (int i = 0; i < rows.Count; i++) {
            var row = rows[i];
            int rowNumber = i + 2;
            cells[offset++] = (rowNumber, 1, row.Id);
            cells[offset++] = (rowNumber, 2, row.Region);
            cells[offset++] = (rowNumber, 3, row.Owner);
            cells[offset++] = (rowNumber, 4, row.CreatedOn);
            cells[offset++] = (rowNumber, 5, row.Amount);
            cells[offset++] = (rowNumber, 6, row.Units);
            cells[offset++] = (rowNumber, 7, row.Active);
            cells[offset++] = (rowNumber, 8, row.Notes);
        }

        return cells;
    }

    private static (int Row, int Column, object Value)[] BuildHeaderlessMixedCells(int rowCount) {
        var cells = new (int Row, int Column, object Value)[checked(rowCount * 3)];
        int offset = 0;
        for (int row = 1; row <= rowCount; row++) {
            cells[offset++] = (row, 1, row * 1.25d);
            cells[offset++] = (row, 2, row % 2 == 0);
            cells[offset++] = (row, 3, "Item " + row.ToString(CultureInfo.InvariantCulture));
        }

        return cells;
    }

    private static (int Row, int Column, object Value)[] BuildSparseObjectCells(int rowCount) {
        var cells = new (int Row, int Column, object Value)[(rowCount + 1) * 4];
        int offset = 0;
        cells[offset++] = (1, 1, "Name");
        cells[offset++] = (1, 2, "Amount");
        cells[offset++] = (1, 3, "Active");
        cells[offset++] = (1, 4, "Created");

        var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
        for (int row = 1; row <= rowCount; row++) {
            int rowNumber = row + 1;
            object? name = row % 3 == 0 ? null : "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
            object? amount = row % 4 == 0 ? null : (double)row * 1.25d;
            object? active = row % 5 == 0 ? null : row % 2 == 0;
            object? created = row % 7 == 0 ? null : start.AddDays(row);
            cells[offset++] = (rowNumber, 1, name!);
            cells[offset++] = (rowNumber, 2, amount!);
            cells[offset++] = (rowNumber, 3, active!);
            cells[offset++] = (rowNumber, 4, created!);
        }

        return cells;
    }

    private static void AddSalesHeaderCells((int Row, int Column, object Value)[] cells, ref int offset) {
        cells[offset++] = (1, 1, "Id");
        cells[offset++] = (1, 2, "Region");
        cells[offset++] = (1, 3, "Owner");
        cells[offset++] = (1, 4, "CreatedOn");
        cells[offset++] = (1, 5, "Amount");
        cells[offset++] = (1, 6, "Units");
        cells[offset++] = (1, 7, "Active");
        cells[offset++] = (1, 8, "Notes");
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
            string tableName = string.Equals(worksheet.Name, "Data", StringComparison.OrdinalIgnoreCase) ? "SalesData" : worksheet.Name;
            var table = worksheet.Tables.Add(worksheet.Cells[1, 1, rows.Count + 1, 8], tableName);
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

    private static void WriteBlogStringHeaders(IXLWorksheet worksheet) {
        for (int i = 0; i < BlogStringColumnNames.Length; i++) {
            worksheet.Cell(1, i + 1).Value = BlogStringColumnNames[i];
        }
    }

    private static void WriteBlogStringRow(IXLWorksheet worksheet, int rowNumber, BlogStringRow row) {
        worksheet.Cell(rowNumber, 1).Value = row.C1;
        worksheet.Cell(rowNumber, 2).Value = row.C2;
        worksheet.Cell(rowNumber, 3).Value = row.C3;
        worksheet.Cell(rowNumber, 4).Value = row.C4;
        worksheet.Cell(rowNumber, 5).Value = row.C5;
        worksheet.Cell(rowNumber, 6).Value = row.C6;
        worksheet.Cell(rowNumber, 7).Value = row.C7;
        worksheet.Cell(rowNumber, 8).Value = row.C8;
        worksheet.Cell(rowNumber, 9).Value = row.C9;
        worksheet.Cell(rowNumber, 10).Value = row.C10;
        worksheet.Cell(rowNumber, 11).Value = row.C11;
        worksheet.Cell(rowNumber, 12).Value = row.C12;
        worksheet.Cell(rowNumber, 13).Value = row.C13;
        worksheet.Cell(rowNumber, 14).Value = row.C14;
        worksheet.Cell(rowNumber, 15).Value = row.C15;
        worksheet.Cell(rowNumber, 16).Value = row.C16;
        worksheet.Cell(rowNumber, 17).Value = row.C17;
        worksheet.Cell(rowNumber, 18).Value = row.C18;
        worksheet.Cell(rowNumber, 19).Value = row.C19;
        worksheet.Cell(rowNumber, 20).Value = row.C20;
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

    private static void WriteSalesRows(IXLWorksheet worksheet, IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns) {
        WriteFullHeaders(worksheet);
        for (int i = 0; i < rows.Count; i++) {
            var row = rows[i];
            int r = i + 2;
            worksheet.Cell(r, 1).Value = row.Id;
            worksheet.Cell(r, 2).Value = row.Region;
            worksheet.Cell(r, 3).Value = row.Owner;
            worksheet.Cell(r, 4).Value = includeAllColumns ? row.CreatedOn : row.Amount;
            if (!includeAllColumns) {
                continue;
            }

            worksheet.Cell(r, 5).Value = row.Amount;
            worksheet.Cell(r, 6).Value = row.Units;
            worksheet.Cell(r, 7).Value = row.Active;
            worksheet.Cell(r, 8).Value = row.Notes;
        }
    }

    private static void WriteSalesRows(ExcelWorksheet worksheet, IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns) {
        WriteHeaders(worksheet);
        for (int i = 0; i < rows.Count; i++) {
            var row = rows[i];
            int r = i + 2;
            worksheet.Cells[r, 1].Value = row.Id;
            worksheet.Cells[r, 2].Value = row.Region;
            worksheet.Cells[r, 3].Value = row.Owner;
            worksheet.Cells[r, 4].Value = includeAllColumns ? row.CreatedOn : row.Amount;
            if (!includeAllColumns) {
                continue;
            }

            worksheet.Cells[r, 5].Value = row.Amount;
            worksheet.Cells[r, 6].Value = row.Units;
            worksheet.Cells[r, 7].Value = row.Active;
            worksheet.Cells[r, 8].Value = row.Notes;
        }
    }

    private static void WriteLargeXlsxDataTable(XlsxWriter writer, string sheetName, DataTable dataTable) {
        writer.BeginWorksheet(sheetName);
        writer.BeginRow();
        foreach (DataColumn column in dataTable.Columns) {
            writer.Write(column.ColumnName);
        }

        foreach (DataRow row in dataTable.Rows) {
            writer.BeginRow();
            foreach (DataColumn column in dataTable.Columns) {
                WriteLargeXlsxValue(writer, row[column]);
            }
        }
    }

    private static void WriteLargeXlsxSalesRows(XlsxWriter writer, IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool includeAllColumns) {
        writer.BeginWorksheet("Data");
        writer.BeginRow()
            .Write("Id")
            .Write("Region")
            .Write("Owner")
            .Write(includeAllColumns ? "CreatedOn" : "Amount");
        if (includeAllColumns) {
            writer.Write("Amount")
                .Write("Units")
                .Write("Active")
                .Write("Notes");
        }

        for (int i = 0; i < rows.Count; i++) {
            var row = rows[i];
            writer.BeginRow()
                .Write(row.Id)
                .Write(row.Region)
                .Write(row.Owner);
            if (includeAllColumns) {
                writer.Write(row.CreatedOn, LargeXlsxDateTimeStyle)
                    .Write(row.Amount)
                    .Write(row.Units)
                    .Write(row.Active)
                    .Write(row.Notes);
            } else {
                writer.Write(row.Amount);
            }
        }
    }

    private static void WriteLargeXlsxPowerShellMixedRows(XlsxWriter writer, IReadOnlyList<Dictionary<string, object?>> rows) {
        writer.BeginWorksheet("Data");
        writer.BeginRow();
        for (int i = 0; i < PowerShellMixedColumnNames.Length; i++) {
            writer.Write(PowerShellMixedColumnNames[i]);
        }

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            Dictionary<string, object?> row = rows[rowIndex];
            writer.BeginRow();
            for (int columnIndex = 0; columnIndex < PowerShellMixedColumnNames.Length; columnIndex++) {
                row.TryGetValue(PowerShellMixedColumnNames[columnIndex], out object? value);
                WriteLargeXlsxValue(writer, value);
            }
        }
    }

    private static void WriteLargeXlsxPowerShellWideRows(XlsxWriter writer, IReadOnlyList<Dictionary<string, object?>> rows) {
        writer.BeginWorksheet("Data");
        writer.BeginRow();
        for (int i = 0; i < PowerShellWideColumnNames.Length; i++) {
            writer.Write(PowerShellWideColumnNames[i]);
        }

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            Dictionary<string, object?> row = rows[rowIndex];
            writer.BeginRow();
            for (int columnIndex = 0; columnIndex < PowerShellWideColumnNames.Length; columnIndex++) {
                row.TryGetValue(PowerShellWideColumnNames[columnIndex], out object? value);
                WriteLargeXlsxValue(writer, value);
            }
        }
    }

    private static void WriteLargeXlsxBlogStringRows(XlsxWriter writer, IReadOnlyList<BlogStringRow> rows) {
        writer.BeginWorksheet("Data");
        writer.BeginRow();
        for (int i = 0; i < BlogStringColumnNames.Length; i++) {
            writer.Write(BlogStringColumnNames[i]);
        }

        for (int i = 0; i < rows.Count; i++) {
            BlogStringRow row = rows[i];
            writer.BeginRow()
                .Write(row.C1)
                .Write(row.C2)
                .Write(row.C3)
                .Write(row.C4)
                .Write(row.C5)
                .Write(row.C6)
                .Write(row.C7)
                .Write(row.C8)
                .Write(row.C9)
                .Write(row.C10)
                .Write(row.C11)
                .Write(row.C12)
                .Write(row.C13)
                .Write(row.C14)
                .Write(row.C15)
                .Write(row.C16)
                .Write(row.C17)
                .Write(row.C18)
                .Write(row.C19)
                .Write(row.C20);
        }
    }

    private static void WriteLargeXlsxValue(XlsxWriter writer, object? value) {
        if (value == null || value == DBNull.Value) {
            writer.Write();
            return;
        }

        switch (value) {
            case string text:
                writer.Write(text);
                break;
            case int number:
                writer.Write(number);
                break;
            case double number:
                writer.Write(number);
                break;
            case decimal number:
                writer.Write(number);
                break;
            case DateTime dateTime:
                writer.Write(dateTime, LargeXlsxDateTimeStyle);
                break;
            case bool flag:
                writer.Write(flag);
                break;
            case byte number:
                writer.Write((int)number);
                break;
            case short number:
                writer.Write((int)number);
                break;
            case long number:
                writer.Write((double)number);
                break;
            case float number:
                writer.Write((double)number);
                break;
            default:
                writer.Write(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }
    }

    private static void WriteFullHeaders(IXLWorksheet worksheet) {
        worksheet.Cell(1, 1).Value = "Id";
        worksheet.Cell(1, 2).Value = "Region";
        worksheet.Cell(1, 3).Value = "Owner";
        worksheet.Cell(1, 4).Value = "CreatedOn";
        worksheet.Cell(1, 5).Value = "Amount";
        worksheet.Cell(1, 6).Value = "Units";
        worksheet.Cell(1, 7).Value = "Active";
        worksheet.Cell(1, 8).Value = "Notes";
    }

    private static void WriteAppendHeaders(ExcelWorksheet worksheet) {
        worksheet.Cells[1, 1].Value = "Id";
        worksheet.Cells[1, 2].Value = "Region";
        worksheet.Cells[1, 3].Value = "Owner";
        worksheet.Cells[1, 4].Value = "Amount";
    }

    private static MiniExcelConfiguration? CreateMiniExcelConfiguration(bool includeTable = false, bool autoFit = false) {
        if (!includeTable && !autoFit) {
            return null;
        }

        return new MiniExcelConfiguration {
            TableStyles = includeTable ? MiniExcelTableStyles.Default : MiniExcelTableStyles.None,
            AutoFilter = includeTable,
            EnableAutoWidth = autoFit,
            FastMode = autoFit
        };
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
           || string.Equals(scenario, "read-used-range", StringComparison.Ordinal)
           || string.Equals(scenario, "read-datareader", StringComparison.Ordinal)
           || string.Equals(scenario, "read-datareader-readonly", StringComparison.Ordinal)
           || string.Equals(scenario, "read-datareader-first-column", StringComparison.Ordinal)
           || string.Equals(scenario, "read-datareader-getvalues", StringComparison.Ordinal)
           || string.Equals(scenario, "read-datareader-typed", StringComparison.Ordinal)
           || string.Equals(scenario, "read-range-decimal", StringComparison.Ordinal)
           || string.Equals(scenario, "enumerate-range", StringComparison.Ordinal)
           || string.Equals(scenario, "enumerate-cells", StringComparison.Ordinal)
           || string.Equals(scenario, "read-range-stream", StringComparison.Ordinal)
           || string.Equals(scenario, "read-first-column-from-wide-sheet", StringComparison.Ordinal)
           || string.Equals(scenario, "read-top-range", StringComparison.Ordinal)
           || string.Equals(scenario, "read-bottom-range", StringComparison.Ordinal)
           || string.Equals(scenario, "read-top-range-stream", StringComparison.Ordinal)
           || string.Equals(scenario, "read-bottom-range-stream", StringComparison.Ordinal)
           || string.Equals(scenario, "read-top-range-stream-small-chunks", StringComparison.Ordinal)
           || string.Equals(scenario, "read-datatable", StringComparison.Ordinal)
           || string.Equals(scenario, "large-sparse-column-read", StringComparison.Ordinal)
           || string.Equals(scenario, "large-sparse-row-read", StringComparison.Ordinal)
           || string.Equals(scenario, "read-objects", StringComparison.Ordinal)
           || string.Equals(scenario, "read-objects-stream", StringComparison.Ordinal)
           || string.Equals(scenario, "formula-heavy-read", StringComparison.Ordinal)
           || string.Equals(scenario, "shared-string-read", StringComparison.Ordinal)
           || string.Equals(scenario, DenseHelloWorldReadRangeScenario, StringComparison.Ordinal)
           || string.Equals(scenario, DenseHelloWorldReadStreamScenario, StringComparison.Ordinal);

    private static DataTable CreateSalesDataTable() {
        var table = new DataTable("Data") { Locale = CultureInfo.InvariantCulture };
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Owner", typeof(string));
        table.Columns.Add("CreatedOn", typeof(DateTime));
        table.Columns.Add("Amount", typeof(double));
        table.Columns.Add("Units", typeof(int));
        table.Columns.Add("Active", typeof(bool));
        table.Columns.Add("Notes", typeof(string));
        return table;
    }

    private static DataSet CreateSalesDataSet(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> first,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> second) {
        var dataSet = new DataSet("Sales") { Locale = CultureInfo.InvariantCulture };
        dataSet.Tables.Add(CreateSalesDataTable(first, "SalesA"));
        dataSet.Tables.Add(CreateSalesDataTable(second, "SalesB"));
        return dataSet;
    }

    private static DataTable CreateSalesDataTable(IEnumerable<ExcelBenchmarkScenarioFactory.SalesRecord> rows, string tableName) {
        var table = CreateSalesDataTable();
        table.TableName = tableName;
        foreach (var row in rows) {
            table.Rows.Add(row.Id, row.Region, row.Owner, row.CreatedOn, row.Amount, row.Units, row.Active, row.Notes);
        }

        return table;
    }

    private static IReadOnlyList<object?> CreateTypedObjectRows(IEnumerable<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        var result = new List<object?>();
        foreach (var row in rows) {
            result.Add(row);
        }

        return result;
    }

    private static IReadOnlyList<object?> CreateDictionaryRows(IEnumerable<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        var result = new List<object?>();
        foreach (var row in rows) {
            result.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                ["Id"] = row.Id,
                ["Region"] = row.Region,
                ["Owner"] = row.Owner,
                ["CreatedOn"] = row.CreatedOn,
                ["Amount"] = row.Amount,
                ["Units"] = row.Units,
                ["Active"] = row.Active,
                ["Notes"] = row.Notes
            });
        }

        return result;
    }

    private static IReadOnlyList<object?> CreateLegacyDictionaryRows(IEnumerable<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        var result = new List<object?>();
        foreach (var row in rows) {
            var dictionary = new System.Collections.Specialized.OrderedDictionary(StringComparer.OrdinalIgnoreCase) {
                ["Id"] = row.Id,
                ["Region"] = row.Region,
                ["Owner"] = row.Owner,
                ["CreatedOn"] = row.CreatedOn,
                ["Amount"] = row.Amount,
                ["Units"] = row.Units,
                ["Active"] = row.Active,
                ["Notes"] = row.Notes
            };
            result.Add(dictionary);
        }

        return result;
    }

    private static IReadOnlyList<Dictionary<string, object?>> CreatePowerShellMixedRows(int count) {
        var result = new List<Dictionary<string, object?>>(count);
        var start = new DateTime(2024, 1, 1, 8, 0, 0, DateTimeKind.Unspecified);
        for (int i = 0; i < count; i++) {
            int id = i + 1;
            result.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                ["Id"] = id,
                ["Name"] = "Server-" + id.ToString("D6", CultureInfo.InvariantCulture),
                ["Department"] = "Department-" + (id % 12 + 1).ToString(CultureInfo.InvariantCulture),
                ["Region"] = PowerShellMixedRegions[i % PowerShellMixedRegions.Length],
                ["IsEnabled"] = id % 4 != 0,
                ["Created"] = start.AddDays(i % 365).AddMinutes(i % 240),
                ["Score"] = Math.Round(100D + ((id * 17.456D) % 900D), 3),
                ["Owner"] = "owner" + (id % 128).ToString(CultureInfo.InvariantCulture) + "@example.test",
                ["TicketCount"] = id % 17,
                ["Notes"] = "Benchmark row " + id.ToString(CultureInfo.InvariantCulture)
            });
        }

        return result;
    }

    private static IReadOnlyList<Dictionary<string, object?>> CreatePowerShellWideRows(int count) {
        var result = new List<Dictionary<string, object?>>(count);
        var start = new DateTime(2024, 1, 1, 8, 0, 0, DateTimeKind.Unspecified);
        for (int i = 0; i < count; i++) {
            int id = i + 1;
            var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                ["Id"] = id,
                ["Name"] = "Server-" + id.ToString("D6", CultureInfo.InvariantCulture),
                ["Created"] = start.AddDays(i % 365).AddMinutes(i % 240),
                ["Enabled"] = id % 4 != 0
            };

            for (int metric = 1; metric <= 36; metric++) {
                row["Metric" + metric.ToString(CultureInfo.InvariantCulture)] = Math.Round(((id * (metric + 7)) % 10000) / 10D, 3);
            }

            result.Add(row);
        }

        return result;
    }

    private static IReadOnlyList<System.Management.Automation.PSObject> CreatePowerShellObjectMixedRows(IEnumerable<Dictionary<string, object?>> rows) {
        var result = new List<System.Management.Automation.PSObject>();
        foreach (var row in rows) {
            var properties = new System.Management.Automation.PSPropertyInfo[row.Count];
            int index = 0;
            foreach (var entry in row) {
                properties[index++] = new System.Management.Automation.PSPropertyInfo(entry.Key, entry.Value);
            }

            result.Add(new System.Management.Automation.PSObject(properties));
        }

        return result;
    }

    private static IReadOnlyList<BlogStringRow> CreateBlogStringRows(int count) {
        var result = new List<BlogStringRow>(count);
        for (int i = 0; i < count; i++) {
            result.Add(BlogStringRow.Create(i));
        }

        return result;
    }

    private static DataTable CreatePowerShellMixedDataTable(IEnumerable<IReadOnlyDictionary<string, object?>> rows, string tableName) {
        var table = new DataTable(tableName) { Locale = CultureInfo.InvariantCulture };
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Department", typeof(string));
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("IsEnabled", typeof(bool));
        table.Columns.Add("Created", typeof(DateTime));
        table.Columns.Add("Score", typeof(double));
        table.Columns.Add("Owner", typeof(string));
        table.Columns.Add("TicketCount", typeof(int));
        table.Columns.Add("Notes", typeof(string));

        foreach (IReadOnlyDictionary<string, object?> sourceRow in rows) {
            object?[] values = new object?[PowerShellMixedColumnNames.Length];
            for (int i = 0; i < values.Length; i++) {
                values[i] = sourceRow.TryGetValue(PowerShellMixedColumnNames[i], out object? value)
                    ? value
                    : DBNull.Value;
            }

            table.Rows.Add(values);
        }

        return table;
    }

    private static DataTable CreatePowerShellWideDataTable(IEnumerable<IReadOnlyDictionary<string, object?>> rows, string tableName) {
        var table = new DataTable(tableName) { Locale = CultureInfo.InvariantCulture };
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Created", typeof(DateTime));
        table.Columns.Add("Enabled", typeof(bool));
        for (int metric = 1; metric <= 36; metric++) {
            table.Columns.Add("Metric" + metric.ToString(CultureInfo.InvariantCulture), typeof(double));
        }

        foreach (IReadOnlyDictionary<string, object?> sourceRow in rows) {
            object?[] values = new object?[PowerShellWideColumnNames.Length];
            for (int i = 0; i < values.Length; i++) {
                values[i] = sourceRow.TryGetValue(PowerShellWideColumnNames[i], out object? value)
                    ? value
                    : DBNull.Value;
            }

            table.Rows.Add(values);
        }

        return table;
    }

    private static DataTable CreateObjectColumnSalesDataTable(IEnumerable<ExcelBenchmarkScenarioFactory.SalesRecord> rows, string tableName) {
        var table = new DataTable(tableName) { Locale = CultureInfo.InvariantCulture };
        table.Columns.Add("Id", typeof(object));
        table.Columns.Add("Region", typeof(object));
        table.Columns.Add("Owner", typeof(object));
        table.Columns.Add("CreatedOn", typeof(object));
        table.Columns.Add("Amount", typeof(object));
        table.Columns.Add("Units", typeof(object));
        table.Columns.Add("Active", typeof(object));
        table.Columns.Add("Notes", typeof(object));

        foreach (var row in rows) {
            table.Rows.Add(row.Id, row.Region, row.Owner, row.CreatedOn, row.Amount, row.Units, row.Active, row.Notes);
        }

        return table;
    }

    private static DataSet CreateSparseDataSet(int rowCount) {
        int firstCount = rowCount / 2;
        int secondCount = rowCount - firstCount;
        var dataSet = new DataSet("Sparse") { Locale = CultureInfo.InvariantCulture };
        dataSet.Tables.Add(CreateSparseDataTable(firstCount, "SparseA", 0));
        dataSet.Tables.Add(CreateSparseDataTable(secondCount, "SparseB", firstCount));
        return dataSet;
    }

    private static DataTable CreateSparseDataTable(int rowCount, string tableName, int offset) {
        var table = new DataTable(tableName) { Locale = CultureInfo.InvariantCulture };
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Owner", typeof(string));
        table.Columns.Add("CreatedOn", typeof(DateTime));
        table.Columns.Add("Amount", typeof(double));
        table.Columns.Add("Units", typeof(int));
        table.Columns.Add("Active", typeof(bool));
        table.Columns.Add("Notes", typeof(string));
        table.Columns.Add("OptionalCode", typeof(string));
        table.Columns.Add("ReviewDate", typeof(DateTime));
        table.Columns.Add("Score", typeof(double));

        for (int i = 0; i < rowCount; i++) {
            int id = offset + i + 1;
            table.Rows.Add(
                id,
                "Item " + id.ToString(CultureInfo.InvariantCulture),
                id % 3 == 0 ? DBNull.Value : "Region " + (id % 5).ToString(CultureInfo.InvariantCulture),
                id % 4 == 0 ? DBNull.Value : "Owner " + (id % 7).ToString(CultureInfo.InvariantCulture),
                id % 5 == 0 ? DBNull.Value : new DateTime(2026, 1, 1).AddDays(id % 365),
                id % 2 == 0 ? DBNull.Value : Math.Round(id * 12.345, 2),
                id % 6 == 0 ? DBNull.Value : id % 17,
                id % 7 == 0 ? DBNull.Value : id % 2 == 0,
                id % 3 == 0 ? DBNull.Value : "Sparse note " + id.ToString(CultureInfo.InvariantCulture),
                id % 8 == 0 ? "C" + id.ToString(CultureInfo.InvariantCulture) : DBNull.Value,
                id % 9 == 0 ? new DateTime(2026, 6, 1).AddDays(id % 30) : DBNull.Value,
                id % 10 == 0 ? Math.Round(id / 10D, 2) : DBNull.Value);
        }

        return table;
    }

    private static int AddSalesDataTableMetric(DataTable table) {
        int metric = 0;
        foreach (DataColumn column in table.Columns) {
            metric = AddStringMetric(metric, column.ColumnName);
        }

        foreach (DataRow row in table.Rows) {
            metric = AddSalesRangeMetric(
                metric,
                Convert.ToInt32(row[0], CultureInfo.InvariantCulture),
                Convert.ToString(row[1], CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(row[2], CultureInfo.InvariantCulture) ?? string.Empty,
                ReadDateCell(row[3]),
                Convert.ToDouble(row[4], CultureInfo.InvariantCulture),
                Convert.ToInt32(row[5], CultureInfo.InvariantCulture),
                Convert.ToBoolean(row[6], CultureInfo.InvariantCulture),
                Convert.ToString(row[7], CultureInfo.InvariantCulture) ?? string.Empty);
        }

        return metric;
    }

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

    private static int AddSalesIdColumnMetric(int metric, int rowIndex, int rowCount, object? value) {
        if (rowIndex == 1) {
            string? header = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (!string.Equals(header, "Id", StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Expected first-column header 'Id', got '{header}'.");
            }

            return AddStringMetric(metric, header);
        }

        int expectedId = rowIndex - 1;
        if (expectedId > rowCount) {
            throw new InvalidOperationException($"Unexpected first-column row {rowIndex.ToString(CultureInfo.InvariantCulture)}.");
        }

        int id = Convert.ToInt32(value, CultureInfo.InvariantCulture);
        if (id != expectedId) {
            throw new InvalidOperationException($"Expected Id {expectedId.ToString(CultureInfo.InvariantCulture)} at row {rowIndex.ToString(CultureInfo.InvariantCulture)}, got {id.ToString(CultureInfo.InvariantCulture)}.");
        }

        return AddIntMetric(metric, id);
    }

    private static int AddSalesIdDataMetric(int metric, int rowIndex, int rowCount, int id) {
        if (rowIndex <= 0 || rowIndex > rowCount) {
            throw new InvalidOperationException($"Unexpected data row {rowIndex.ToString(CultureInfo.InvariantCulture)}.");
        }

        if (id != rowIndex) {
            throw new InvalidOperationException($"Expected Id {rowIndex.ToString(CultureInfo.InvariantCulture)}, got {id.ToString(CultureInfo.InvariantCulture)}.");
        }

        return AddIntMetric(metric, id);
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

    private static int AddSalesRangeDecimalMetric(
        int metric,
        int id,
        string region,
        string owner,
        DateTime createdOn,
        decimal amount,
        int units,
        bool active,
        string notes) {
        metric = AddIntMetric(metric, id);
        metric = AddStringMetric(metric, region);
        metric = AddStringMetric(metric, owner);
        metric = AddIntMetric(metric, createdOn.DayOfYear);
        metric = AddDecimalMetric(metric, amount);
        metric = AddIntMetric(metric, units);
        metric = AddIntMetric(metric, active ? 1 : 0);
        return AddStringMetric(metric, notes);
    }

    private static int AddSalesEnumeratedCellMetric(int metric, int row, int column, object? value) {
        metric = AddIntMetric(metric, row);
        metric = AddIntMetric(metric, column);
        if (row == 1) {
            return AddStringMetric(metric, Convert.ToString(value, CultureInfo.InvariantCulture));
        }

        return column switch {
            1 or 6 => AddIntMetric(metric, Convert.ToInt32(value, CultureInfo.InvariantCulture)),
            4 => AddIntMetric(metric, ReadDateCell(value).DayOfYear),
            5 => AddDoubleMetric(metric, Convert.ToDouble(value, CultureInfo.InvariantCulture)),
            7 => AddIntMetric(metric, Convert.ToBoolean(value, CultureInfo.InvariantCulture) ? 1 : 0),
            _ => AddStringMetric(metric, Convert.ToString(value, CultureInfo.InvariantCulture))
        };
    }

    private static DateTime ReadDateCell(object? value) {
        if (value is DateTime dateTime) {
            return dateTime;
        }

        if (value is string text && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)) {
            return parsed;
        }

        return DateTime.FromOADate(Convert.ToDouble(value, CultureInfo.InvariantCulture));
    }

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

    private static int AddDecimalMetric(int metric, decimal value) {
        unchecked {
            return AddIntMetric(metric, (int)Math.Round(value * 100m, 0, MidpointRounding.AwayFromZero));
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

    private static int AddHelloWorldMetric(int metric, object? value, int rowIndex, string columnName) {
        string? text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (!string.Equals(text, HelloWorldValue, StringComparison.Ordinal)) {
            throw new InvalidOperationException($"Expected {HelloWorldValue} at {columnName}{rowIndex.ToString(CultureInfo.InvariantCulture)}, got '{text}'.");
        }

        return AddStringMetric(metric, text);
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
        public double StandardDeviationMilliseconds { get; init; }
        public double StandardErrorMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
        public double AverageAllocatedBytes { get; init; }
        public double MedianAllocatedBytes { get; init; }
        public List<long> SamplesAllocatedBytes { get; init; } = [];
    }

    private sealed record LibraryComparisonCase(string Library, string Notes, Func<int> Action);

    private sealed record PackageProfileCase(string Library, string Notes, Func<byte[]> CreatePackage);

    private sealed class ExcelPackageProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public string BuildConfiguration { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public int WarmupIterations { get; init; }
        public int MeasuredIterations { get; init; }
        public string Notes { get; init; } = string.Empty;
        public List<ExcelPackageProfileScenario> Scenarios { get; init; } = [];
    }

    private sealed class ExcelPackageProfileScenario {
        public string Scenario { get; init; } = string.Empty;
        public string Library { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
        public double AverageMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public double StandardDeviationMilliseconds { get; init; }
        public double StandardErrorMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
        public double AverageAllocatedBytes { get; init; }
        public double MedianAllocatedBytes { get; init; }
        public List<long> SamplesAllocatedBytes { get; init; } = [];
        public ExcelPackageProfileSummary Package { get; init; } = new();
    }

    private sealed class ExcelPackageProfileSummary {
        public long FileSizeBytes { get; init; }
        public int PartCount { get; set; }
        public int WorksheetRowCount { get; set; }
        public int WorksheetCellCount { get; set; }
        public int SharedStringCount { get; set; }
        public int UniqueSharedStringCount { get; set; }
        public int CellStyleCount { get; set; }
        public long WorksheetCompressedBytes { get; set; }
        public long WorksheetUncompressedBytes { get; set; }
        public long SharedStringsCompressedBytes { get; set; }
        public long SharedStringsUncompressedBytes { get; set; }
        public long StylesCompressedBytes { get; set; }
        public long StylesUncompressedBytes { get; set; }
        public long TablesCompressedBytes { get; set; }
        public long TablesUncompressedBytes { get; set; }
        public long WorkbookCompressedBytes { get; set; }
        public long WorkbookUncompressedBytes { get; set; }
        public long RelationshipsCompressedBytes { get; set; }
        public long RelationshipsUncompressedBytes { get; set; }
        public long DocPropsCompressedBytes { get; set; }
        public long DocPropsUncompressedBytes { get; set; }
        public long OtherCompressedBytes { get; set; }
        public long OtherUncompressedBytes { get; set; }
        public List<ExcelPackagePartProfile> Parts { get; set; } = [];
    }

    private sealed class ExcelPackagePartProfile {
        public string Name { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public long CompressedBytes { get; init; }
        public long UncompressedBytes { get; init; }
    }

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

    private sealed class MiniExcelAppendRecord {
        public int Id { get; set; }
        public string Region { get; set; } = string.Empty;
        public string Owner { get; set; } = string.Empty;
        public double Amount { get; set; }
    }

    private sealed class MiniExcelStringRecord {
        public string Repeated { get; set; } = string.Empty;
        public string Distinct { get; set; } = string.Empty;
        public string LongSegment { get; set; } = string.Empty;
    }

    private sealed class BlogStringRow {
        public string C1 { get; init; } = string.Empty;
        public string C2 { get; init; } = string.Empty;
        public string C3 { get; init; } = string.Empty;
        public string C4 { get; init; } = string.Empty;
        public string C5 { get; init; } = string.Empty;
        public string C6 { get; init; } = string.Empty;
        public string C7 { get; init; } = string.Empty;
        public string C8 { get; init; } = string.Empty;
        public string C9 { get; init; } = string.Empty;
        public string C10 { get; init; } = string.Empty;
        public string C11 { get; init; } = string.Empty;
        public string C12 { get; init; } = string.Empty;
        public string C13 { get; init; } = string.Empty;
        public string C14 { get; init; } = string.Empty;
        public string C15 { get; init; } = string.Empty;
        public string C16 { get; init; } = string.Empty;
        public string C17 { get; init; } = string.Empty;
        public string C18 { get; init; } = string.Empty;
        public string C19 { get; init; } = string.Empty;
        public string C20 { get; init; } = string.Empty;

        internal static BlogStringRow Create(int rowIndex) {
            int start = checked(rowIndex * 1000);
            return new BlogStringRow {
                C1 = (start + 0).ToString(CultureInfo.InvariantCulture),
                C2 = (start + 1).ToString(CultureInfo.InvariantCulture),
                C3 = (start + 2).ToString(CultureInfo.InvariantCulture),
                C4 = (start + 3).ToString(CultureInfo.InvariantCulture),
                C5 = (start + 4).ToString(CultureInfo.InvariantCulture),
                C6 = (start + 5).ToString(CultureInfo.InvariantCulture),
                C7 = (start + 6).ToString(CultureInfo.InvariantCulture),
                C8 = (start + 7).ToString(CultureInfo.InvariantCulture),
                C9 = (start + 8).ToString(CultureInfo.InvariantCulture),
                C10 = (start + 9).ToString(CultureInfo.InvariantCulture),
                C11 = (start + 10).ToString(CultureInfo.InvariantCulture),
                C12 = (start + 11).ToString(CultureInfo.InvariantCulture),
                C13 = (start + 12).ToString(CultureInfo.InvariantCulture),
                C14 = (start + 13).ToString(CultureInfo.InvariantCulture),
                C15 = (start + 14).ToString(CultureInfo.InvariantCulture),
                C16 = (start + 15).ToString(CultureInfo.InvariantCulture),
                C17 = (start + 16).ToString(CultureInfo.InvariantCulture),
                C18 = (start + 17).ToString(CultureInfo.InvariantCulture),
                C19 = (start + 18).ToString(CultureInfo.InvariantCulture),
                C20 = (start + 19).ToString(CultureInfo.InvariantCulture)
            };
        }
    }
}
