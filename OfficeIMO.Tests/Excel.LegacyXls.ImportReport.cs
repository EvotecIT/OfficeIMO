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
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.MacroSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.ChartSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.VbaModuleSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x01|MacroSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x02|ChartSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x06|VbaModuleSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["Macro1"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["Chart1"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["Module1"]);
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
            Assert.Contains("Unsupported Sheets By Kind", markdown);
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
            Assert.Equal(5, report.PreservedFeatureRecordCount);
            Assert.Equal(3, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DrawingObject]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(3, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.DrawingObject]);
            Assert.Equal(1, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(1, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(1, report.PivotTableRecordCount);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.PreserveOnly]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxView"]);
            Assert.Equal(1, report.ChartRecordCount);
            Assert.Equal(1, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Container]);
            Assert.Equal(1, report.ChartRecordsByName["Chart"]);
            Assert.Equal(1, report.ChartRecordsByLocation["FeatureMap"]);
            Assert.Equal(3, report.DrawingRecordCount);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.DrawingGroup]);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.Drawing]);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.Object]);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawingGroup"]);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawing"]);
            Assert.Equal(1, report.DrawingRecordsByName["Obj"]);
            Assert.Equal(1, report.DrawingRecordsByObjectType["ObjectType:0x0008"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordType["EscherRecordType:0xF000"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordType["EscherRecordType:0xF002"]);
            Assert.Equal(1, report.DrawingRecordsByLocation["(workbook)"]);
            Assert.Equal(2, report.DrawingRecordsByLocation["FeatureMap"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawingGroup"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawing"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawingGroup"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawing"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView"]);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:MsoDrawingGroup" && record.SheetName == null);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:Obj" && record.SheetName == "FeatureMap");
            Assert.Contains(workbook.DrawingRecords, record => record.SheetName == "FeatureMap" && record.ObjectType == 0x0008 && record.ObjectId == 1);
            Assert.Contains(workbook.DrawingRecords, record => record.RecordName == "MsoDrawingGroup" && record.EscherRecordType == 0xf000 && record.EscherRecordInstance == 2 && record.EscherRecordVersion == 0x0f && record.EscherPayloadLength == 8);
            Assert.Contains(workbook.DrawingRecords, record => record.RecordName == "MsoDrawing" && record.EscherRecordType == 0xf002 && record.EscherRecordInstance == 1 && record.EscherRecordVersion == 0x0f && record.EscherPayloadLength == 0);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.DetailCode == "Chart:Chart" && record.RecordType == 0x1002);
            Assert.Contains(workbook.Diagnostics, d => d.DetailCode == "Chart:Chart");
            Assert.Contains(workbook.Diagnostics, d => d.DetailCode == "PivotTable:SxView");
            string markdown = report.ToMarkdown();
            Assert.Contains("Preserved feature records: 5", markdown);
            Assert.Contains("Drawing:MsoDrawingGroup", markdown);
            Assert.Contains("Pivot Table Records By Name", markdown);
            Assert.Contains("Chart Records By Name", markdown);
            Assert.Contains("Drawing Records By Name", markdown);
            Assert.Contains("Drawing Records By Object Type", markdown);
            Assert.Contains("Drawing Records By Escher Record Type", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_DecodesPivotTableMetadataRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5PivotTableMetadataWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(3, workbook.PivotTableRecords.Count);
            Assert.Equal(3, report.PivotTableRecordCount);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.DataItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.GroupingRange]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.ExtendedPivotField]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxdi"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxRng"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxVdEx"]);
            Assert.Equal(1, report.PivotTableDataItemAggregations["AggregationFunction:0"]);
            Assert.Equal(1, report.PivotTableGroupingKinds["Months"]);

            LegacyXlsPivotTableRecord dataItem = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.DataItem);
            Assert.Null(dataItem.SheetName);
            Assert.Equal("Sxdi", dataItem.RecordName);
            Assert.Equal((short)2, dataItem.DataItemFieldIndex);
            Assert.Equal((short)0, dataItem.AggregationFunction);
            Assert.Equal((short)7, dataItem.DisplayCalculation);
            Assert.Equal((ushort)14, dataItem.NumberFormatId);
            Assert.Equal("Sales", dataItem.Name);

            LegacyXlsPivotTableRecord grouping = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.GroupingRange);
            Assert.Equal("PivotMeta", grouping.SheetName);
            Assert.Equal("SxRng", grouping.RecordName);
            Assert.True(grouping.AutoStart);
            Assert.True(grouping.AutoEnd);
            Assert.Equal(LegacyXlsPivotGroupingKind.Months, grouping.GroupingKind);

            LegacyXlsPivotTableRecord extended = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.ExtendedPivotField);
            Assert.Equal("PivotMeta", extended.SheetName);
            Assert.Equal("SxVdEx", extended.RecordName);
            Assert.True(extended.ShowAllItems);
            Assert.True(extended.CanDragToRow);
            Assert.True(extended.CanDragToColumn);
            Assert.True(extended.CanDragToPage);
            Assert.True(extended.CanDragToHide);
            Assert.False(extended.PreventDragToData);
            Assert.True(extended.ServerBased);
            Assert.Equal(3, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(3, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Pivot table records: 3", markdown);
            Assert.Contains("Pivot Table Records By Kind", markdown);
            Assert.Contains("Pivot Table Data Item Aggregations", markdown);
            Assert.Contains("Pivot Table Grouping Kinds", markdown);
            Assert.Contains("SxVdEx", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_CountsCalculationSettings() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateCalculationSettingsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(7, workbook.CalculationSettings.Records.Count);
            Assert.Equal(LegacyXlsCalculationMode.Automatic, workbook.CalculationSettings.Mode);
            Assert.Equal((short)42, workbook.CalculationSettings.IterationCount);
            Assert.True(workbook.CalculationSettings.FullPrecision);
            Assert.True(workbook.CalculationSettings.A1ReferenceMode);
            Assert.Equal(0.001d, workbook.CalculationSettings.Delta!.Value);
            Assert.True(workbook.CalculationSettings.IterationEnabled);
            Assert.True(workbook.CalculationSettings.RecalculateBeforeSave);
            Assert.DoesNotContain(workbook.CalculationSettings.Records, record => record.SheetName != null);
            Assert.DoesNotContain(workbook.UnsupportedFeatures, feature => feature.DetailCode == "BiffRecord:CalcMode");
            Assert.DoesNotContain(workbook.UnsupportedFeatures, feature => feature.DetailCode == "BiffRecord:CalcCount");
            Assert.Equal(7, report.CalculationSettingRecordCount);
            Assert.Equal(1, report.CalculationSettingsByKind[LegacyXlsCalculationSettingKind.Mode]);
            Assert.Equal(1, report.CalculationSettingsByKind[LegacyXlsCalculationSettingKind.IterationCount]);
            Assert.Equal(1, report.CalculationSettingsByKind[LegacyXlsCalculationSettingKind.FullPrecision]);
            Assert.Equal(1, report.CalculationSettingsByKind[LegacyXlsCalculationSettingKind.A1ReferenceMode]);
            Assert.Equal(1, report.CalculationSettingsByKind[LegacyXlsCalculationSettingKind.Delta]);
            Assert.Equal(1, report.CalculationSettingsByKind[LegacyXlsCalculationSettingKind.IterationEnabled]);
            Assert.Equal(1, report.CalculationSettingsByKind[LegacyXlsCalculationSettingKind.RecalculateBeforeSave]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Calculation setting records: 7", markdown);
            Assert.Contains("Calculation Settings By Kind", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_ScansUnsupportedChartSheetSubstreams() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ChartSheetSubstreamWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(workbook.Worksheets);
            Assert.Equal("Data", sheet.Name);
            LegacyXlsUnsupportedSheet unsupportedSheet = Assert.Single(workbook.UnsupportedSheets);
            Assert.Equal("ChartOnly", unsupportedSheet.Name);
            Assert.Equal(LegacyXlsUnsupportedSheetKind.ChartSheet, unsupportedSheet.Kind);
            Assert.Equal(1, unsupportedSheet.ChartTextObjectCount);
            Assert.Equal(5, report.UnsupportedFeatureCount);
            Assert.Equal(4, report.PreservedFeatureRecordCount);
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.ChartSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x02|ChartSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["ChartOnly"]);
            Assert.Equal(1, report.UnsupportedSheetMetadataRecordCount);
            Assert.Equal(1, report.UnsupportedSheetMetadataRecordsByKind[LegacyXlsUnsupportedSheetMetadataKind.ChartTextObject]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.ChartSheet]);
            Assert.Equal(4, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(4, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(4, report.ChartRecordCount);
            Assert.Equal(2, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Container]);
            Assert.Equal(1, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Formatting]);
            Assert.Equal(1, report.ChartRecordsByKind[LegacyXlsChartRecordKind.ChartType]);
            Assert.Equal(1, report.ChartRecordsByName["Units"]);
            Assert.Equal(1, report.ChartRecordsByName["Chart"]);
            Assert.Equal(1, report.ChartRecordsByName["ChartFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["Scatter"]);
            Assert.Equal(1, report.ChartRecordsByChartType["Scatter"]);
            Assert.Equal(4, report.ChartRecordsByLocation["ChartOnly"]);
            Assert.Equal(1, report.DrawingRecordCount);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.TextObject]);
            Assert.Equal(1, report.DrawingRecordsByName["TxO"]);
            Assert.Equal(1, report.DrawingRecordsByLocation["ChartOnly"]);
            Assert.Equal(4, report.UnsupportedFeaturesByLocation["XLS-BIFF-FEATURE-CHART-UNSUPPORTED|ChartOnly"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Units"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Scatter"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Units"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Scatter"]);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.SheetName == "ChartOnly" && record.DetailCode == "Chart:Chart");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Chart");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.ChartTypeName == "Scatter");
            Assert.Contains(workbook.Diagnostics, d => d.SheetName == "ChartOnly" && d.DetailCode == "Chart:Chart");
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
            Assert.Equal(1, report.DataValidationsByType["Date"]);
            Assert.Equal(1, report.DataValidationsByType["Time"]);
            Assert.Equal(1, report.DataValidationsByType["TextLength"]);
            Assert.Equal(1, report.DataValidationsByOperator["Between"]);
            Assert.Equal(1, report.DataValidationsByOperator["GreaterThanOrEqual"]);
            Assert.Equal(1, report.DataValidationsByOperator["LessThanOrEqual"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("Data validations: 3", markdown);
            Assert.Contains("Data Validations By Type", markdown);
            Assert.Contains("Data Validations By Operator", markdown);
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
            Assert.Equal(1, report.ConditionalFormattingsByType["CellIs"]);
            Assert.Equal(1, report.ConditionalFormattingsByOperator["GreaterThan"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("Conditional formatting rules: 1", markdown);
            Assert.Contains("Conditional Formatting By Type", markdown);
            Assert.Contains("Conditional Formatting By Operator", markdown);
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
            Assert.Equal(1, report.AutoFilterCriteriaByOperator["Equal"]);
            Assert.Equal(1, report.AutoFilterCriteriaByOperator["GreaterThanOrEqual"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("AutoFilter criteria columns: 2", markdown);
            Assert.Contains("AutoFilter Criteria By Operator", markdown);
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
            LegacyXlsCompoundFeatureRecord compoundRecord = Assert.Single(result.Workbook.CompoundFeatureRecords);
            Assert.Equal(LegacyXlsCompoundFeatureRecordKind.VbaProject, compoundRecord.Kind);
            Assert.Contains("_VBA_PROJECT_CUR", compoundRecord.Entries);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntryCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordsByKind[LegacyXlsCompoundFeatureRecordKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKind[LegacyXlsCompoundFeatureRecordKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByName["_VBA_PROJECT_CUR"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["VbaProject|XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED|Compound:VbaProjectStorage"]);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("VbaProject", markdown);
            Assert.Contains("Compound Feature Entries By Name", markdown);
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
            LegacyXlsCompoundFeatureRecord compoundRecord = Assert.Single(result.Workbook.CompoundFeatureRecords);
            Assert.Equal(LegacyXlsCompoundFeatureRecordKind.OleObject, compoundRecord.Kind);
            Assert.Contains("ObjectPool", compoundRecord.Entries);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntryCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordsByKind[LegacyXlsCompoundFeatureRecordKind.OleObject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKind[LegacyXlsCompoundFeatureRecordKind.OleObject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByName["ObjectPool"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.OleObject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["OleObject|XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED|Compound:OleObjectStorage"]);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("OleObject", markdown);
            Assert.Contains("Compound Feature Entries By Name", markdown);
        }
    }
}
