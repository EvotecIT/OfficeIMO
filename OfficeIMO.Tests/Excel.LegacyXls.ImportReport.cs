using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_ImportReport_SummarizesCorpusSignals() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5UnsupportedSheetTypesWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            LegacyXlsImportReport report = result.ImportReport;
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(3, report.UnsupportedSheetCount);
            Assert.Equal(1, report.CellCount);
            Assert.Equal(0, report.FormulaCellCount);
            Assert.Equal(0, report.CommentCount);
            Assert.Equal(0, report.HyperlinkCount);
            Assert.Equal(0, report.DataValidationCount);
            Assert.Equal(0, report.ConditionalFormattingCount);
            Assert.Equal(0, report.AutoFilterCriteriaCount);
            Assert.Equal(3, report.UnsupportedFeatureCount);
            Assert.False(report.HasImportErrors);
            Assert.True(report.HasUnsupportedFeatures);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.MacroSheet]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.ChartSheet]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.VbaModuleSheet]);
            Assert.Equal(1, report.UnsupportedFeaturesByCode["XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED"]);
            Assert.Equal(1, report.UnsupportedFeaturesByRecordType["MacroSheet|XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED|0x0085"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["MacroSheet|XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED|Sheet:MacroSheet"]);
            Assert.Equal(1, report.UnsupportedFeaturesByLocation["XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED|Macro1"]);
            Assert.Equal(1, report.DiagnosticsByCode["XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Worksheets: 1", markdown);
            Assert.Contains("Unsupported sheets: 3", markdown);
            Assert.Contains("XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED", markdown);
            Assert.Contains("Unsupported Feature Record Types", markdown);
            Assert.Contains("Unsupported Feature Details", markdown);
            Assert.Contains("Sheet:ChartSheet", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_NamesPreserveOnlyFeatureDetails() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5PreserveOnlyFeatureDetailsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(5, report.UnsupportedFeatureCount);
            Assert.Equal(3, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DrawingObject]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawingGroup"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawing"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView"]);
            Assert.Contains(workbook.Diagnostics, d => d.DetailCode == "Chart:Chart");
            Assert.Contains(workbook.Diagnostics, d => d.DetailCode == "PivotTable:SxView");
            Assert.Contains("Drawing:MsoDrawingGroup", report.ToMarkdown());
        }

        [Fact]
        public void LegacyXls_ImportReport_CountsImportedWorkbookFeatures() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4DefinedNamesWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(7, report.CellCount);
            Assert.Equal(5, report.DefinedNameCount);
            Assert.Equal(0, report.DataValidationCount);
            Assert.Equal(0, report.ConditionalFormattingCount);
            Assert.Equal(0, report.AutoFilterCriteriaCount);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.False(report.HasUnsupportedFeatures);
        }

        [Fact]
        public void LegacyXls_ImportReport_CountsImportedDataValidations() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4TypedDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(3, report.DataValidationCount);
            Assert.Equal(0, report.ConditionalFormattingCount);
            Assert.Equal(0, report.AutoFilterCriteriaCount);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.Contains("Data validations: 3", report.ToMarkdown());
        }

        [Fact]
        public void LegacyXls_ImportReport_CountsImportedConditionalFormatting() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4ConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(1, report.ConditionalFormattingCount);
            Assert.Equal(0, report.AutoFilterCriteriaCount);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.Contains("Conditional formatting rules: 1", report.ToMarkdown());
        }

        [Fact]
        public void LegacyXls_ImportReport_CountsImportedAutoFilterCriteria() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AutoFilterCriteriaWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(2, report.AutoFilterCriteriaCount);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.Contains("AutoFilter criteria columns: 2", report.ToMarkdown());
        }

        [Fact]
        public void LegacyXls_Load_ReportsVbaProjectStorageAsPreserveOnly() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithVbaProjectStorage(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Single(result.Document.Sheets);
            LegacyXlsUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.VbaProject);
            Assert.Equal("XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED", feature.Code);
            Assert.Contains("_VBA_PROJECT_CUR", feature.Description);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED");
            Assert.True(result.ImportReport.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["VbaProject|XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED|Compound:VbaProjectStorage"]);
            Assert.Contains("VbaProject", result.ImportReport.ToMarkdown());
        }

        [Fact]
        public void LegacyXls_Load_ReportsOleObjectStorageAsPreserveOnly() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithOleObjectStorage(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Single(result.Document.Sheets);
            LegacyXlsUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.OleObject);
            Assert.Equal("XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED", feature.Code);
            Assert.Contains("ObjectPool", feature.Description);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED");
            Assert.True(result.ImportReport.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.OleObject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["OleObject|XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED|Compound:OleObjectStorage"]);
            Assert.Contains("OleObject", result.ImportReport.ToMarkdown());
        }
    }
}
