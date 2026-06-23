using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private const string LegacyXlsComValidationEnv = "OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION";
        private const int XlExcel8FileFormat = 56;
        private const int XlCellValueCondition = 1;
        private const int XlGreaterCondition = 5;
        private const int XlValidateList = 3;
        private const int XlValidAlertStop = 1;
        private const int XlBetween = 1;

        [Fact]
        public void LegacyXls_ComGeneratedWorkbook_ImportsAndOpensInDesktopExcelWhenRequested() {
            if (!IsLegacyXlsComValidationRequested()) {
                return;
            }

            Assert.True(IsWindowsPlatform(), "Legacy XLS COM validation requires Windows.");
            Assert.True(IsExcelComAvailable(), "Legacy XLS COM validation requires Microsoft Excel COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "LegacyXlsComValidation");
            Directory.CreateDirectory(directory);
            string sourceXlsPath = Path.Combine(directory, "ExcelGenerated.LegacyFeatures.xls");
            string importedXlsxPath = Path.Combine(directory, "ExcelGenerated.LegacyFeatures.imported.xlsx");

            CreateLegacyXlsWorkbookViaExcelCom(sourceXlsPath);
            AssertWorkbooksOpenViaExcelComWhenAvailable(new[] { sourceXlsPath }, "The generated legacy XLS workbook did not open through desktop Excel.");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(sourceXlsPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.True(result.ImportReport.CellCount >= 16);
            Assert.True(result.ImportReport.FormulaCellCount >= 1);

            result.Document.Save(importedXlsxPath, openExcel: false);
            AssertWorkbooksOpenViaExcelComWhenAvailable(new[] { importedXlsxPath }, "The imported XLSX workbook did not open through desktop Excel.");
        }

        [Fact]
        public void LegacyXls_CorpusFixtures_OpenBeforeAndAfterImportInDesktopExcelWhenRequested() {
            if (!IsLegacyXlsComValidationRequested()) {
                return;
            }

            Assert.True(IsWindowsPlatform(), "Legacy XLS COM validation requires Windows.");
            Assert.True(IsExcelComAvailable(), "Legacy XLS COM validation requires Microsoft Excel COM automation.");

            string corpusDirectory = Path.Combine(GetTestsProjectRoot(), "Documents", "LegacyXlsCorpus");
            if (!Directory.Exists(corpusDirectory)) {
                return;
            }

            string[] workbookPaths = Directory.GetFiles(corpusDirectory, "*.xls", SearchOption.AllDirectories)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (workbookPaths.Length == 0) {
                return;
            }

            AssertWorkbooksOpenViaExcelComWhenAvailable(workbookPaths, "One or more legacy XLS corpus source workbooks failed to open in desktop Excel.");

            string outputDirectory = Path.Combine(_directoryWithFiles, "LegacyXlsCorpusComValidation");
            Directory.CreateDirectory(outputDirectory);
            var importedPaths = new List<string>(workbookPaths.Length);
            foreach (string workbookPath in workbookPaths) {
                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                    ReportUnsupportedRecords = true
                });
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);

                string outputName = GetRelativePath(corpusDirectory, workbookPath)
                    .Replace(Path.DirectorySeparatorChar, '_')
                    .Replace(Path.AltDirectorySeparatorChar, '_');
                string importedPath = Path.Combine(outputDirectory, Path.ChangeExtension(outputName, ".imported.xlsx"));
                result.Document.Save(importedPath, openExcel: false);
                importedPaths.Add(importedPath);
            }

            AssertWorkbooksOpenViaExcelComWhenAvailable(importedPaths, "One or more imported legacy XLS corpus workbooks failed to open in desktop Excel.");
        }

        private static bool IsLegacyXlsComValidationRequested() {
            string? value = Environment.GetEnvironmentVariable(LegacyXlsComValidationEnv);
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateLegacyXlsWorkbookViaExcelCom(string path) {
            var failures = new List<string>();
            var thread = new Thread(() => {
                try {
                    CreateLegacyXlsWorkbookViaExcelComOnStaThread(path);
                } catch (Exception ex) when (ex is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException) {
                    failures.Add(DescribeExcelComFailure(ex));
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(ExcelComOpenTimeout)) {
                failures.Add($"Excel COM legacy XLS generation timed out after {ExcelComOpenTimeout.TotalSeconds:0} seconds.");
            }

            Assert.True(failures.Count == 0, "Failed to generate the legacy XLS workbook through desktop Excel." + Environment.NewLine + string.Join(Environment.NewLine, failures));
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateLegacyXlsWorkbookViaExcelComOnStaThread(string path) {
            object? excel = null;
            object? workbooks = null;
            object? workbook = null;
            object? worksheet = null;

            try {
                var excelType = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException("Excel COM automation is not available.");
                excel = Activator.CreateInstance(excelType)
                    ?? throw new InvalidOperationException("Failed to create Excel COM automation instance.");

                SetComProperty(excel, "DisplayAlerts", false);
                SetComProperty(excel, "Visible", false);
                workbooks = GetComProperty(excel, "Workbooks");
                workbook = InvokeCom(workbooks!, "Add");
                worksheet = GetComProperty(workbook!, "Worksheets", 1);
                SetComProperty(worksheet!, "Name", "Data");

                SetCellValue(worksheet!, 1, 1, "Status");
                SetCellValue(worksheet!, 1, 2, "Amount");
                SetCellValue(worksheet!, 1, 3, "Adjusted");
                SetCellValue(worksheet!, 1, 4, "When");
                SetCellValue(worksheet!, 2, 1, "Open");
                SetCellValue(worksheet!, 2, 2, 125.5d);
                SetCellFormula(worksheet!, 2, 3, "=B2*1.1");
                SetCellValue(worksheet!, 2, 4, new DateTime(2026, 1, 15));
                SetCellValue(worksheet!, 3, 1, "Closed");
                SetCellValue(worksheet!, 3, 2, 80d);
                SetCellFormula(worksheet!, 3, 3, "=B3*1.1");
                SetCellValue(worksheet!, 3, 4, new DateTime(2026, 2, 20));
                SetCellValue(worksheet!, 4, 1, "Open");
                SetCellValue(worksheet!, 4, 2, 210.25d);
                SetCellFormula(worksheet!, 4, 3, "=B4*1.1");
                SetCellValue(worksheet!, 4, 4, new DateTime(2026, 3, 10));
                SetCellValue(worksheet!, 5, 1, "Pending");
                SetCellValue(worksheet!, 5, 2, 42d);
                SetCellFormula(worksheet!, 5, 3, "=B5*1.1");
                SetCellValue(worksheet!, 5, 4, new DateTime(2026, 4, 5));

                object headerRange = GetComProperty(worksheet!, "Range", "A1:D1")!;
                SetComProperty(GetComProperty(headerRange, "Font")!, "Bold", true);
                SetComProperty(GetComProperty(headerRange, "Interior")!, "Color", 0xD9EAD3);
                SetComProperty(GetComProperty(worksheet!, "Range", "B2:C5")!, "NumberFormat", "$#,##0.00");
                SetComProperty(GetComProperty(worksheet!, "Range", "D2:D5")!, "NumberFormat", "m/d/yyyy");
                SetComProperty(GetComProperty(worksheet!, "Columns", "A:D")!, "ColumnWidth", 14d);

                object dataRange = GetComProperty(worksheet!, "Range", "A1:D5")!;
                InvokeCom(dataRange, "AutoFilter", 1, "Open");

                object statusValidation = GetComProperty(GetComProperty(worksheet!, "Range", "A2:A5")!, "Validation")!;
                InvokeCom(statusValidation, "Delete");
                InvokeCom(statusValidation, "Add", XlValidateList, XlValidAlertStop, XlBetween, "Open,Closed,Pending");

                object formatConditions = GetComProperty(GetComProperty(worksheet!, "Range", "B2:B5")!, "FormatConditions")!;
                object condition = InvokeCom(formatConditions, "Add", XlCellValueCondition, XlGreaterCondition, "100")!;
                SetComProperty(GetComProperty(condition, "Interior")!, "Color", 0xC6EFCE);

                InvokeCom(GetComProperty(worksheet!, "Range", "A2")!, "AddComment", "Generated by Excel COM for legacy XLS import validation.");
                object hyperlinks = GetComProperty(worksheet!, "Hyperlinks")!;
                InvokeCom(hyperlinks, "Add", GetComProperty(worksheet!, "Range", "A6"), "https://officeimo.net/", Type.Missing, "OfficeIMO", "OfficeIMO");

                SetComProperty(GetComProperty(worksheet!, "PageSetup")!, "PrintTitleRows", "$1:$1");
                SetComProperty(GetComProperty(worksheet!, "PageSetup")!, "PrintArea", "$A$1:$D$6");
                InvokeCom(GetComProperty(workbook!, "Names")!, "Add", "Amounts", "=Data!$B$2:$B$5");

                if (File.Exists(path)) {
                    File.Delete(path);
                }

                InvokeCom(workbook!, "SaveAs", path, XlExcel8FileFormat);
            } finally {
                try {
                    workbook?.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                } catch (Exception ex) when (ex is COMException or MissingMethodException or TargetInvocationException) {
                    throw new InvalidOperationException("Failed to close generated legacy XLS workbook.", ex);
                } finally {
                    try {
                        excel?.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excel, null);
                    } catch (Exception ex) when (ex is COMException or MissingMethodException or TargetInvocationException) {
                        throw new InvalidOperationException("Failed to quit Excel after legacy XLS generation.", ex);
                    }

                    ReleaseComObject(worksheet);
                    ReleaseComObject(workbook);
                    ReleaseComObject(workbooks);
                    ReleaseComObject(excel);
                }
            }
        }

        private static object? GetComProperty(object target, string name, params object[] args) =>
            target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, args.Length == 0 ? null : args);

        private static void SetComProperty(object target, string name, object value) =>
            target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new[] { value });

        private static object? InvokeCom(object target, string name, params object[] args) =>
            target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args.Length == 0 ? null : args);

        private static void SetCellValue(object worksheet, int row, int column, object value) {
            object cell = GetComProperty(worksheet, "Cells", row, column)!;
            SetComProperty(cell, "Value2", value);
            ReleaseComObject(cell);
        }

        private static void SetCellFormula(object worksheet, int row, int column, string formula) {
            object cell = GetComProperty(worksheet, "Cells", row, column)!;
            SetComProperty(cell, "Formula", formula);
            ReleaseComObject(cell);
        }

        private static void ReleaseComObject(object? value) {
            if (value != null && Marshal.IsComObject(value)) {
                Marshal.FinalReleaseComObject(value);
            }
        }
    }
}
