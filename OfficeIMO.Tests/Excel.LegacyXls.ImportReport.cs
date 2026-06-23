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
            LegacyXlsPivotTableRecord pivotRecord = Assert.Single(workbook.PivotTableRecords);
            Assert.Equal(LegacyXlsPivotTableRecordKind.View, pivotRecord.Kind);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.View]);
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
            Assert.Equal(1, report.DrawingRecordsByObjectTypeName["Picture"]);
            Assert.Equal(1, report.DrawingRecordsByObjectFlags["ObjectFlags:0x4011"]);
            Assert.Equal(1, report.DrawingRecordsByObjectFlagName["Locked"]);
            Assert.Equal(1, report.DrawingRecordsByObjectFlagName["Printable"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordType["EscherRecordType:0xF000"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordType["EscherRecordType:0xF002"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtDggContainer"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtDgContainer"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByType["Png"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByEmbeddedRecordType["OfficeArtBlipPNG"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesBySize["SizeBytes:12"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByReferenceCount["References:1"]);
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
            Assert.Contains(workbook.DrawingRecords, record => record.SheetName == "FeatureMap" && record.ObjectType == 0x0008 && record.ObjectTypeKind == LegacyXlsDrawingObjectType.Picture && record.ObjectTypeName == "Picture" && record.ObjectId == 1 && record.ObjectFlags == 0x4011 && record.IsObjectLocked && record.IsObjectPrintable);
            LegacyXlsDrawingRecord drawingGroup = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "MsoDrawingGroup");
            Assert.Equal((ushort)0xf000, drawingGroup.EscherRecordType);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtDggContainer, drawingGroup.EscherRecordTypeKind);
            Assert.Equal("OfficeArtDggContainer", drawingGroup.EscherRecordTypeName);
            Assert.Equal((ushort)2, drawingGroup.EscherRecordInstance);
            Assert.Equal((byte)0x0f, drawingGroup.EscherRecordVersion);
            Assert.Equal((uint)80, drawingGroup.EscherPayloadLength);
            LegacyXlsDrawingBlipStoreEntry blipEntry = Assert.Single(drawingGroup.BlipStoreEntries);
            Assert.Equal((ushort)0x0006, blipEntry.RecordInstance);
            Assert.Equal(LegacyXlsDrawingBlipType.Png, blipEntry.RecordInstanceBlipTypeKind);
            Assert.Equal("Png", blipEntry.RecordInstanceBlipTypeName);
            Assert.Equal((byte)0x06, blipEntry.Win32BlipType);
            Assert.Equal(LegacyXlsDrawingBlipType.Png, blipEntry.Win32BlipTypeKind);
            Assert.Equal("Png", blipEntry.Win32BlipTypeName);
            Assert.Equal((byte)0x06, blipEntry.MacOsBlipType);
            Assert.Equal(LegacyXlsDrawingBlipType.Png, blipEntry.MacOsBlipTypeKind);
            Assert.Equal("Png", blipEntry.MacOsBlipTypeName);
            Assert.Equal((uint)12, blipEntry.SizeBytes);
            Assert.Equal((uint)1, blipEntry.ReferenceCount);
            Assert.Equal((ushort)0xf01e, blipEntry.EmbeddedBlipRecordType);
            Assert.Equal("OfficeArtBlipPNG", blipEntry.EmbeddedBlipRecordTypeName);
            Assert.Equal((uint)4, blipEntry.EmbeddedBlipPayloadLength);
            Assert.Contains(workbook.DrawingRecords, record => record.RecordName == "MsoDrawing" && record.EscherRecordType == 0xf002 && record.EscherRecordTypeKind == LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer && record.EscherRecordTypeName == "OfficeArtDgContainer" && record.EscherRecordInstance == 1 && record.EscherRecordVersion == 0x0f && record.EscherPayloadLength == 0);
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
            Assert.Contains("Drawing Records By Object Type Name", markdown);
            Assert.Contains("Picture", markdown);
            Assert.Contains("Drawing Records By Object Flags", markdown);
            Assert.Contains("Drawing Records By Object Flag Name", markdown);
            Assert.Contains("Drawing Records By Escher Record Type", markdown);
            Assert.Contains("Drawing Records By Escher Record Type Name", markdown);
            Assert.Contains("Drawing BLIP Store Entries By Type", markdown);
            Assert.Contains("OfficeArtBlipPNG", markdown);
            Assert.Contains("OfficeArtDggContainer", markdown);
        }

        [Theory]
        [InlineData(0x0005, LegacyXlsDrawingObjectType.Chart, "Chart")]
        [InlineData(0x0008, LegacyXlsDrawingObjectType.Picture, "Picture")]
        [InlineData(0x0019, LegacyXlsDrawingObjectType.Note, "Note")]
        [InlineData(0x001E, LegacyXlsDrawingObjectType.OfficeArtObject, "OfficeArtObject")]
        public void LegacyXlsDrawingRecord_DecodesKnownObjectTypeNames(int objectType, LegacyXlsDrawingObjectType expectedKind, string expectedName) {
            var record = new LegacyXlsDrawingRecord(
                LegacyXlsDrawingRecordKind.Object,
                "Obj",
                "Sheet1",
                0,
                0x005d,
                22,
                checked((ushort)objectType),
                1);

            Assert.Equal(expectedKind, record.ObjectTypeKind);
            Assert.Equal(expectedName, record.ObjectTypeName);
        }

        [Fact]
        public void LegacyXlsDrawingRecord_UsesHexObjectTypeNameForUnknownObjectTypes() {
            var record = new LegacyXlsDrawingRecord(
                LegacyXlsDrawingRecordKind.Object,
                "Obj",
                "Sheet1",
                0,
                0x005d,
                22,
                0x0fff,
                1);

            Assert.Null(record.ObjectTypeKind);
            Assert.Equal("ObjectType:0x0FFF", record.ObjectTypeName);
        }

        [Fact]
        public void LegacyXlsDrawingRecord_DecodesObjectFlagNames() {
            var record = new LegacyXlsDrawingRecord(
                LegacyXlsDrawingRecordKind.Object,
                "Obj",
                "Sheet1",
                0,
                0x005d,
                22,
                0x0008,
                1,
                objectFlags: 0x1395);

            Assert.True(record.IsObjectLocked);
            Assert.True(record.UsesDefaultObjectSize);
            Assert.True(record.IsObjectPrintable);
            Assert.True(record.IsObjectDisabled);
            Assert.True(record.IsUiObject);
            Assert.True(record.RecalculatesObjectOnLoad);
            Assert.True(record.AlwaysRecalculatesObject);
            Assert.Equal(new[] { "Locked", "DefaultSize", "Printable", "Disabled", "UiObject", "RecalculateOnLoad", "AlwaysRecalculate" }, record.ObjectFlagNames);
        }

        [Theory]
        [InlineData(0xF000, LegacyXlsDrawingEscherRecordType.OfficeArtDggContainer, "OfficeArtDggContainer")]
        [InlineData(0xF002, LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer, "OfficeArtDgContainer")]
        [InlineData(0xF004, LegacyXlsDrawingEscherRecordType.OfficeArtSpContainer, "OfficeArtSpContainer")]
        [InlineData(0xF011, LegacyXlsDrawingEscherRecordType.OfficeArtFClientData, "OfficeArtFClientData")]
        public void LegacyXlsDrawingRecord_DecodesKnownEscherRecordTypeNames(int recordType, LegacyXlsDrawingEscherRecordType expectedKind, string expectedName) {
            var record = new LegacyXlsDrawingRecord(
                LegacyXlsDrawingRecordKind.Drawing,
                "MsoDrawing",
                "Sheet1",
                0,
                0x00ec,
                8,
                escherRecordType: checked((ushort)recordType));

            Assert.Equal(expectedKind, record.EscherRecordTypeKind);
            Assert.Equal(expectedName, record.EscherRecordTypeName);
        }

        [Fact]
        public void LegacyXlsDrawingRecord_UsesHexEscherRecordTypeNameForUnknownRecordTypes() {
            var record = new LegacyXlsDrawingRecord(
                LegacyXlsDrawingRecordKind.Drawing,
                "MsoDrawing",
                "Sheet1",
                0,
                0x00ec,
                8,
                escherRecordType: 0xffff);

            Assert.Null(record.EscherRecordTypeKind);
            Assert.Equal("EscherRecordType:0xFFFF", record.EscherRecordTypeName);
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
            Assert.Equal(12, workbook.PivotTableRecords.Count);
            Assert.Equal(12, report.PivotTableRecordCount);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.View]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Field]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Item]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.DataItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Cache]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.CacheItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Table]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.GroupingRange]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Filter]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Format]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.ExtendedPivotField]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.PivotChart]);
            Assert.DoesNotContain(report.PivotTableRecordsByKind, entry => entry.Key == LegacyXlsPivotTableRecordKind.PreserveOnly);
            Assert.Equal(1, report.PivotTableRecordsByName["SxView"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvd"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvi"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxdi"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxdb"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxnum"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxtbl"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxRng"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxFilt"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxFormat"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxVdEx"]);
            Assert.Equal(1, report.PivotTableRecordsByName["PivotChartBits"]);
            Assert.Equal(1, report.PivotTableDataItemAggregations["AggregationFunction:0"]);
            Assert.Equal(1, report.PivotTableDataItemAggregationKinds["Sum"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculations["PercentOfGrandTotal"]);
            Assert.Equal(1, report.PivotTableGroupingKinds["Months"]);

            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.View && record.RecordName == "SxView" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Field && record.RecordName == "Sxvd" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Item && record.RecordName == "Sxvi" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Cache && record.RecordName == "Sxdb" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.CacheItem && record.RecordName == "Sxnum" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Table && record.RecordName == "Sxtbl" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Filter && record.RecordName == "SxFilt" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Format && record.RecordName == "SxFormat" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.PivotChart && record.RecordName == "PivotChartBits" && record.SheetName == "PivotMeta");

            LegacyXlsPivotTableRecord dataItem = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.DataItem);
            Assert.Null(dataItem.SheetName);
            Assert.Equal("Sxdi", dataItem.RecordName);
            Assert.Equal((short)2, dataItem.DataItemFieldIndex);
            Assert.Equal((short)0, dataItem.AggregationFunction);
            Assert.Equal(LegacyXlsPivotAggregationFunction.Sum, dataItem.AggregationFunctionKind);
            Assert.Equal("Sum", dataItem.AggregationFunctionName);
            Assert.Equal((short)7, dataItem.DisplayCalculation);
            Assert.Equal(LegacyXlsPivotDisplayCalculation.PercentOfGrandTotal, dataItem.DisplayCalculationKind);
            Assert.Equal("PercentOfGrandTotal", dataItem.DisplayCalculationName);
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
            Assert.Equal(12, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(12, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Pivot table records: 12", markdown);
            Assert.Contains("Pivot Table Records By Kind", markdown);
            Assert.Contains("Pivot Table Data Item Aggregations", markdown);
            Assert.Contains("Pivot Table Data Item Aggregation Kinds", markdown);
            Assert.Contains("PercentOfGrandTotal", markdown);
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
            Assert.Equal(17, report.UnsupportedFeatureCount);
            Assert.Equal(16, report.PreservedFeatureRecordCount);
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.ChartSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x02|ChartSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["ChartOnly"]);
            Assert.Equal(1, report.UnsupportedSheetMetadataRecordCount);
            Assert.Equal(1, report.UnsupportedSheetMetadataRecordsByKind[LegacyXlsUnsupportedSheetMetadataKind.ChartTextObject]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.ChartSheet]);
            Assert.Equal(16, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(16, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(16, report.ChartRecordCount);
            Assert.Equal(2, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Container]);
            Assert.Equal(3, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Axis]);
            Assert.Equal(1, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Series]);
            Assert.Equal(5, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Formatting]);
            Assert.Equal(2, report.ChartRecordsByKind[LegacyXlsChartRecordKind.ChartType]);
            Assert.Equal(3, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Text]);
            Assert.Equal(1, report.ChartRecordsByName["Units"]);
            Assert.Equal(1, report.ChartRecordsByName["Chart"]);
            Assert.Equal(1, report.ChartRecordsByName["DataFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["ChartFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["Axis"]);
            Assert.Equal(1, report.ChartRecordsByName["Tick"]);
            Assert.Equal(1, report.ChartRecordsByName["AxesUsed"]);
            Assert.Equal(1, report.ChartRecordsByName["Series"]);
            Assert.Equal(1, report.ChartRecordsByName["Scatter"]);
            Assert.Equal(1, report.ChartRecordsByName["LineFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["AreaFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["MarkerFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["DefaultText"]);
            Assert.Equal(1, report.ChartRecordsByName["Text"]);
            Assert.Equal(1, report.ChartRecordsByName["ObjectLink"]);
            Assert.Equal(1, report.ChartRecordsByName["Legend"]);
            Assert.Equal(1, report.ChartRecordsByChartType["Scatter"]);
            Assert.Equal(1, report.ChartRecordsByRectangle["X:100;Y:200;Width:3000;Height:2200"]);
            Assert.Equal(1, report.ChartRecordsByAxisType["ValueOrVerticalValue"]);
            Assert.Equal(1, report.ChartRecordsByAxesUsedCount["AxesUsed:1"]);
            Assert.Equal(1, report.ChartSeriesCategoryDataTypes["Text"]);
            Assert.Equal(1, report.ChartSeriesValueCounts["Categories:4;Values:4;BubbleSizes:0"]);
            Assert.Equal(1, report.ChartDataFormatTargets["Series"]);
            Assert.Equal(1, report.ChartDataFormatSeriesIndexes["SeriesIndex:2"]);
            Assert.Equal(1, report.ChartLineFormatStyles["Dash"]);
            Assert.Equal(1, report.ChartLineFormatWeights["Medium"]);
            Assert.Equal(1, report.ChartAreaFormatPatterns["Solid"]);
            Assert.Equal(1, report.ChartMarkerFormatTypes["Circle"]);
            Assert.Equal(1, report.ChartMarkerFormatSizes["SizeTwips:240"]);
            Assert.Equal(1, report.ChartDefaultTextTargets["ChartUnscaledText"]);
            Assert.Equal(1, report.ChartTextHorizontalAlignments["Center"]);
            Assert.Equal(1, report.ChartTextVerticalAlignments["Bottom"]);
            Assert.Equal(1, report.ChartTextDataLabelPositions["Center"]);
            Assert.Equal(1, report.ChartTextFlags["AutoColor"]);
            Assert.Equal(1, report.ChartTextFlags["ShowValue"]);
            Assert.Equal(1, report.ChartTextFlags["AutoText"]);
            Assert.Equal(1, report.ChartTextFlags["AutoMode"]);
            Assert.Equal(1, report.ChartTextFlags["ShowLabel"]);
            Assert.Equal(1, report.ChartObjectLinkTargets["SeriesOrDataPoint"]);
            Assert.Equal(1, report.ChartLegendLayouts["Vertical"]);
            Assert.Equal(1, report.ChartTickMajorLocations["Outside"]);
            Assert.Equal(1, report.ChartTickLabelLocations["NextToAxis"]);
            Assert.Equal(16, report.ChartRecordsByLocation["ChartOnly"]);
            Assert.Equal(1, report.DrawingRecordCount);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.TextObject]);
            Assert.Equal(1, report.DrawingRecordsByName["TxO"]);
            Assert.Equal(1, report.DrawingRecordsByLocation["ChartOnly"]);
            Assert.Equal(16, report.UnsupportedFeaturesByLocation["XLS-BIFF-FEATURE-CHART-UNSUPPORTED|ChartOnly"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Units"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DataFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:LineFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AreaFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:MarkerFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Axis"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Tick"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxesUsed"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Series"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Scatter"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DefaultText"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Text"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ObjectLink"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Legend"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Units"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DataFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:LineFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AreaFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:MarkerFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Axis"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Tick"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxesUsed"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Series"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Scatter"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DefaultText"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Text"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ObjectLink"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Legend"]);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.SheetName == "ChartOnly" && record.DetailCode == "Chart:Chart");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Chart" && record.ChartX == 100 && record.ChartY == 200 && record.ChartWidth == 3000 && record.ChartHeight == 2200);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Axis" && record.AxisType == 0x0001 && record.AxisTypeName == "ValueOrVerticalValue");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AxesUsed" && record.AxesUsedCount == 1);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Series" && record.SeriesCategoryDataType == 0x0003 && record.SeriesCategoryDataTypeName == "Text" && record.SeriesValueDataType == 0x0001 && record.SeriesCategoryCount == 4 && record.SeriesValueCount == 4 && record.SeriesBubbleSizeDataType == 0x0001 && record.SeriesBubbleSizeCount == 0);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "DataFormat" && record.DataFormatPointIndex == 0xffff && record.DataFormatSeriesIndex == 2 && record.DataFormatOrder == 1 && record.DataFormatTarget == "Series");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "LineFormat" && record.LineFormat != null && record.LineFormat.RgbHex == "#112233" && record.LineFormat.Style == 0x0001 && record.LineFormat.StyleName == "Dash" && record.LineFormat.Weight == 1 && record.LineFormat.WeightName == "Medium" && !record.LineFormat.Automatic && record.LineFormat.AxisVisible && !record.LineFormat.AutomaticColor && record.LineFormat.ColorIndex == 0x004d);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AreaFormat" && record.AreaFormat != null && record.AreaFormat.ForegroundRgbHex == "#AABBCC" && record.AreaFormat.BackgroundRgbHex == "#102030" && record.AreaFormat.Pattern == 0x0001 && record.AreaFormat.PatternName == "Solid" && record.AreaFormat.Automatic && record.AreaFormat.InvertNegative && record.AreaFormat.ForegroundColorIndex == 0x004e && record.AreaFormat.BackgroundColorIndex == 0x004d);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "MarkerFormat" && record.MarkerFormat != null && record.MarkerFormat.ForegroundRgbHex == "#DEADBE" && record.MarkerFormat.BackgroundRgbHex == "#445566" && record.MarkerFormat.MarkerType == 0x0008 && record.MarkerFormat.MarkerTypeName == "Circle" && record.MarkerFormat.Automatic && !record.MarkerFormat.InteriorHidden && record.MarkerFormat.BorderHidden && record.MarkerFormat.ForegroundColorIndex == 0x004e && record.MarkerFormat.BackgroundColorIndex == 0x004d && record.MarkerFormat.SizeTwips == 240);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "DefaultText" && record.DefaultTextId == 0x0002 && record.DefaultTextTargetName == "ChartUnscaledText");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Text" && record.Text != null && record.Text.HorizontalAlignmentName == "Center" && record.Text.VerticalAlignmentName == "Bottom" && record.Text.BackgroundModeName == "Transparent" && record.Text.RgbHex == "#224466" && record.Text.X == 120 && record.Text.Y == 240 && record.Text.Width == 800 && record.Text.Height == 160 && record.Text.Flags == 0x4095 && record.Text.FlagNames.SequenceEqual(new[] { "AutoColor", "ShowValue", "AutoText", "AutoMode", "ShowLabel" }) && record.Text.ColorIndex == 0x004d && record.Text.DataLabelPositionName == "Center" && record.Text.ReadingOrderName == "LeftToRight" && record.Text.Rotation == 45);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "ObjectLink" && record.ObjectLink != null && record.ObjectLink.LinkedObject == 0x0004 && record.ObjectLink.LinkedObjectName == "SeriesOrDataPoint" && record.ObjectLink.SeriesIndex == 2 && record.ObjectLink.DataPointIndex == 0xffff);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Legend" && record.Legend != null && record.Legend.X == 10 && record.Legend.Y == 20 && record.Legend.Width == 300 && record.Legend.Height == 400 && record.Legend.Spacing == 1 && record.Legend.Flags == 0x001d && record.Legend.AutoPosition && record.Legend.AutoPositionX && record.Legend.AutoPositionY && record.Legend.Vertical && !record.Legend.WasDataTable);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Tick" && record.Tick != null && record.Tick.MajorTickLocationName == "Outside" && record.Tick.MinorTickLocationName == "Inside" && record.Tick.LabelLocationName == "NextToAxis" && record.Tick.BackgroundModeName == "Transparent" && record.Tick.RgbHex == "#998877" && record.Tick.Flags == 0x402d && record.Tick.RotationModeName == "RotatedClockwise" && record.Tick.AutoColor && !record.Tick.AutoBackground && record.Tick.AutoRotation && record.Tick.ReadingOrderName == "LeftToRight" && record.Tick.ColorIndex == 0x004d && record.Tick.Rotation == 30);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.ChartTypeName == "Scatter");
            Assert.Contains(workbook.Diagnostics, d => d.SheetName == "ChartOnly" && d.DetailCode == "Chart:Chart");
            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Records By Rectangle", markdown);
            Assert.Contains("Chart Records By Axis Type", markdown);
            Assert.Contains("Chart Records By Axes Used Count", markdown);
            Assert.Contains("Chart Series Category Data Types", markdown);
            Assert.Contains("Chart Series Value Counts", markdown);
            Assert.Contains("Chart DataFormat Targets", markdown);
            Assert.Contains("Chart DataFormat Series Indexes", markdown);
            Assert.Contains("Chart LineFormat Styles", markdown);
            Assert.Contains("Chart LineFormat Weights", markdown);
            Assert.Contains("Chart AreaFormat Patterns", markdown);
            Assert.Contains("Chart MarkerFormat Types", markdown);
            Assert.Contains("Chart MarkerFormat Sizes", markdown);
            Assert.Contains("Chart DefaultText Targets", markdown);
            Assert.Contains("Chart Text Horizontal Alignments", markdown);
            Assert.Contains("Chart ObjectLink Targets", markdown);
            Assert.Contains("Chart Tick Label Locations", markdown);
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
            Assert.Equal(LegacyXlsCompoundFeatureEntryRole.VbaProjectStorage, compoundRecord.EntryRoles["_VBA_PROJECT_CUR"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntryCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordsByKind[LegacyXlsCompoundFeatureRecordKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKind[LegacyXlsCompoundFeatureRecordKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByName["_VBA_PROJECT_CUR"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRole["VbaProjectStorage"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKindAndRole["VbaProject|VbaProjectStorage"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["VbaProject|XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED|Compound:VbaProjectStorage"]);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("VbaProject", markdown);
            Assert.Contains("Compound Feature Entries By Name", markdown);
            Assert.Contains("Compound Feature Entries By Role", markdown);
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
            Assert.Equal(LegacyXlsCompoundFeatureEntryRole.OleObjectPoolStorage, compoundRecord.EntryRoles["ObjectPool"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntryCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordsByKind[LegacyXlsCompoundFeatureRecordKind.OleObject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKind[LegacyXlsCompoundFeatureRecordKind.OleObject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByName["ObjectPool"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRole["OleObjectPoolStorage"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKindAndRole["OleObject|OleObjectPoolStorage"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.OleObject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["OleObject|XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED|Compound:OleObjectStorage"]);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("OleObject", markdown);
            Assert.Contains("Compound Feature Entries By Name", markdown);
            Assert.Contains("Compound Feature Entries By Role", markdown);
        }
    }
}
