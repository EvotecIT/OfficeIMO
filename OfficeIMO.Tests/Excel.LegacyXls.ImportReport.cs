using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
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
            Assert.Equal(1, report.WorksheetsByVisibility["Visible"]);
            Assert.Equal(3, report.UnsupportedSheetsByVisibility["Visible"]);
            Assert.Equal(1, report.UnsupportedSheetsByKindAndVisibility["MacroSheet|Visible"]);
            Assert.Equal(1, report.UnsupportedSheetsByKindAndVisibility["ChartSheet|Visible"]);
            Assert.Equal(1, report.UnsupportedSheetsByKindAndVisibility["VbaModuleSheet|Visible"]);
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
            Assert.Contains("Unsupported Sheets By Visibility", markdown);
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
            Assert.Equal(8, report.UnsupportedFeatureCount);
            Assert.Equal(8, report.PreservedFeatureRecordCount);
            Assert.Equal(6, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DrawingObject]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(6, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.DrawingObject]);
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
            Assert.Equal(6, report.DrawingRecordCount);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.DrawingGroup]);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.Drawing]);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.Object]);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.ShapePropertiesStream]);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.TextPropertiesStream]);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.RichTextStream]);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawingGroup"]);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawing"]);
            Assert.Equal(1, report.DrawingRecordsByName["Obj"]);
            Assert.Equal(1, report.DrawingRecordsByName["ShapePropsStream"]);
            Assert.Equal(1, report.DrawingRecordsByName["TextPropsStream"]);
            Assert.Equal(1, report.DrawingRecordsByName["RichTextStream"]);
            Assert.Equal(1, report.DrawingRecordsByObjectType["ObjectType:0x0008"]);
            Assert.Equal(1, report.DrawingRecordsByObjectTypeName["Picture"]);
            Assert.Equal(1, report.DrawingRecordsByObjectFlags["ObjectFlags:0x4011"]);
            Assert.Equal(1, report.DrawingRecordsByObjectFlagName["Locked"]);
            Assert.Equal(1, report.DrawingRecordsByObjectFlagName["Printable"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByType["SubRecordType:0x0015"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByType["SubRecordType:0x000D"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByType["SubRecordType:0x0000"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByName["FtCmo"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByName["FtNts"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByName["FtEnd"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByDeclaredLength["DeclaredBytes:18"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByDeclaredLength["DeclaredBytes:22"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByDeclaredLength["DeclaredBytes:0"]);
            Assert.Equal(3, report.DrawingObjectSubRecordsByCompleteness["Complete"]);
            Assert.Equal(1, report.DrawingFutureRecordWrappedTypes["ShapePropsStream|0x08A3"]);
            Assert.Equal(1, report.DrawingFutureRecordWrappedTypes["TextPropsStream|0x08A4"]);
            Assert.Equal(1, report.DrawingFutureRecordWrappedTypes["RichTextStream|0x08A5"]);
            Assert.Equal(1, report.DrawingFutureRecordFlags["ShapePropsStream|Flags:0x0000"]);
            Assert.Equal(1, report.DrawingFutureRecordFlags["TextPropsStream|Flags:0x0001"]);
            Assert.Equal(1, report.DrawingFutureRecordFlags["RichTextStream|Flags:0x0000"]);
            Assert.Equal(1, report.DrawingFutureRecordReferenceStates["ShapePropsStream|NoRange"]);
            Assert.Equal(1, report.DrawingFutureRecordReferenceStates["TextPropsStream|HasRange"]);
            Assert.Equal(1, report.DrawingFutureRecordReferenceStates["RichTextStream|NoRange"]);
            Assert.Equal(1, report.DrawingFutureRecordRanges["TextPropsStream|A2:B3"]);
            Assert.Equal(1, report.DrawingFutureRecordStreamByteCounts["ShapePropsStream|StreamBytes:5"]);
            Assert.Equal(1, report.DrawingFutureRecordStreamByteCounts["TextPropsStream|StreamBytes:4"]);
            Assert.Equal(1, report.DrawingFutureRecordStreamByteCounts["RichTextStream|StreamBytes:4"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordType["EscherRecordType:0xF000"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordType["EscherRecordType:0xF002"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtDggContainer"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtDgContainer"]);
            Assert.Equal(11, report.DrawingOfficeArtRecordCount);
            Assert.Equal(1, report.DrawingGroupBlockCount);
            Assert.Equal(1, report.DrawingGroupInfoCount);
            Assert.Equal(1, report.DrawingIdentifierClusterCount);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByType["EscherRecordType:0xF000"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByType["EscherRecordType:0xF006"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByType["EscherRecordType:0xF008"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByType["EscherRecordType:0xF00B"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByType["EscherRecordType:0xF00F"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByTypeName["OfficeArtDggContainer"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByTypeName["OfficeArtFDGGBlock"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByTypeName["OfficeArtFDG"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByTypeName["OfficeArtFOPT"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByTypeName["OfficeArtChildAnchor"]);
            Assert.Equal(2, report.DrawingOfficeArtRecordsByDepth["Depth:0"]);
            Assert.Equal(4, report.DrawingOfficeArtRecordsByDepth["Depth:1"]);
            Assert.Equal(5, report.DrawingOfficeArtRecordsByDepth["Depth:2"]);
            Assert.Equal(4, report.DrawingOfficeArtRecordsByContainerState["Container"]);
            Assert.Equal(7, report.DrawingOfficeArtRecordsByContainerState["Leaf"]);
            Assert.Equal(2, report.DrawingOfficeArtRecordsByPayloadLength["PayloadLength:8"]);
            Assert.Equal(2, report.DrawingOfficeArtRecordsByPayloadLength["PayloadLength:16"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByPayloadLength["PayloadLength:24"]);
            Assert.Equal(1, report.DrawingGroupBlocksByMaxShapeId["MaxShapeId:2048"]);
            Assert.Equal(1, report.DrawingGroupBlocksByDeclaredIdentifierClusterCount["DeclaredIdentifierClusters:2"]);
            Assert.Equal(1, report.DrawingGroupBlocksByDecodedIdentifierClusterCount["DecodedIdentifierClusters:1"]);
            Assert.Equal(1, report.DrawingGroupBlocksBySavedShapeCount["SavedShapes:2"]);
            Assert.Equal(1, report.DrawingGroupBlocksBySavedDrawingCount["SavedDrawings:1"]);
            Assert.Equal(1, report.DrawingIdentifierClustersByDrawingId["DrawingId:1"]);
            Assert.Equal(1, report.DrawingIdentifierClustersByCurrentShapeId["CurrentShapeId:1024"]);
            Assert.Equal(1, report.DrawingGroupInfosByDrawingId["DrawingId:1"]);
            Assert.Equal(1, report.DrawingGroupInfosByShapeCount["Shapes:1"]);
            Assert.Equal(1, report.DrawingGroupInfosByLastShapeId["LastShapeId:1024"]);
            Assert.Equal(2, report.DrawingShapePropertyCount);
            Assert.Equal(1, report.DrawingShapePropertiesById["PropertyId:0x00BF"]);
            Assert.Equal(1, report.DrawingShapePropertiesById["PropertyId:0x0005"]);
            Assert.Equal(1, report.DrawingShapePropertiesByName["TextBooleanProperties"]);
            Assert.Equal(1, report.DrawingShapePropertiesByName["PropertyId:0x0005"]);
            Assert.Equal(1, report.DrawingShapePropertiesByGroup["Text"]);
            Assert.Equal(1, report.DrawingShapePropertiesByGroup["Protection"]);
            Assert.Equal(1, report.DrawingShapePropertiesByFlagState["Simple"]);
            Assert.Equal(1, report.DrawingShapePropertiesByFlagState["Complex"]);
            Assert.Equal(1, report.DrawingShapePropertiesByValue["PropertyId:0x00BF;Value:0x00000001"]);
            Assert.Equal(1, report.DrawingShapeComplexPropertiesByDeclaredLength["PropertyId:0x0005;DeclaredBytes:4"]);
            Assert.Equal(1, report.DrawingShapeComplexPropertiesByAvailableLength["PropertyId:0x0005;AvailableBytes:4"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByType["Png"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByEmbeddedRecordType["OfficeArtBlipPNG"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesBySize["SizeBytes:12"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByReferenceCount["References:1"]);
            Assert.Equal(1, report.DrawingPictureStates["PictureObjects:Present|BlipStore:Present|PictureBlipReferences:Missing|ReferencedBlips:None"]);
            Assert.Equal(1, report.DrawingShapeEntriesByType["PictureFrame"]);
            Assert.Equal(1, report.DrawingShapeEntriesById["ShapeId:1024"]);
            Assert.Equal(1, report.DrawingShapeEntriesByFlags["Flags:0x00000A02"]);
            Assert.Equal(1, report.DrawingShapeEntriesByFlagName["Child"]);
            Assert.Equal(1, report.DrawingShapeEntriesByFlagName["HaveAnchor"]);
            Assert.Equal(1, report.DrawingShapeEntriesByFlagName["HaveShapeType"]);
            Assert.Equal(1, report.DrawingAnchorEntriesByRange["R2C1:R4C3"]);
            Assert.Equal(1, report.DrawingAnchorEntriesByOffset["StartDx:10;StartDy:20;EndDx:30;EndDy:40"]);
            Assert.Equal(1, report.DrawingAnchorEntriesByFlags["Flags:0x0000"]);
            Assert.Equal(1, report.DrawingChildAnchorEntriesByRectangle["Left:100;Top:200;Right:700;Bottom:900"]);
            Assert.Equal(1, report.DrawingChildAnchorEntriesBySize["Width:600;Height:700"]);
            Assert.Equal(1, report.DrawingRecordsByLocation["(workbook)"]);
            Assert.Equal(5, report.DrawingRecordsByLocation["FeatureMap"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawingGroup"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawing"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:ShapePropsStream"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:TextPropsStream"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:RichTextStream"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawingGroup"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawing"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:ShapePropsStream"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:TextPropsStream"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:RichTextStream"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView"]);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:MsoDrawingGroup" && record.SheetName == null);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:Obj" && record.SheetName == "FeatureMap");
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:ShapePropsStream" && record.SheetName == "FeatureMap");
            Assert.Contains(workbook.DrawingRecords, record => record.SheetName == "FeatureMap" && record.ObjectType == 0x0008 && record.ObjectTypeKind == LegacyXlsDrawingObjectType.Picture && record.ObjectTypeName == "Picture" && record.ObjectId == 1 && record.ObjectFlags == 0x4011 && record.IsObjectLocked && record.IsObjectPrintable);
            LegacyXlsDrawingRecord objectRecord = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "Obj");
            Assert.Equal(new[] { "FtCmo", "FtNts", "FtEnd" }, objectRecord.ObjectSubRecords.Select(subRecord => subRecord.SubRecordName).ToArray());
            LegacyXlsDrawingObjectSubRecord commonObjectSubRecord = objectRecord.ObjectSubRecords[0];
            Assert.Equal((ushort)0x0015, commonObjectSubRecord.SubRecordType);
            Assert.Equal("SubRecordType:0x0015", commonObjectSubRecord.SubRecordTypeKey);
            Assert.Equal(0, commonObjectSubRecord.Offset);
            Assert.Equal((ushort)18, commonObjectSubRecord.DeclaredLength);
            Assert.Equal(18, commonObjectSubRecord.AvailableLength);
            Assert.True(commonObjectSubRecord.IsComplete);
            LegacyXlsDrawingRecord drawingGroup = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "MsoDrawingGroup");
            Assert.Equal((ushort)0xf000, drawingGroup.EscherRecordType);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtDggContainer, drawingGroup.EscherRecordTypeKind);
            Assert.Equal("OfficeArtDggContainer", drawingGroup.EscherRecordTypeName);
            Assert.Equal((ushort)2, drawingGroup.EscherRecordInstance);
            Assert.Equal((byte)0x0f, drawingGroup.EscherRecordVersion);
            Assert.Equal((uint)96, drawingGroup.EscherPayloadLength);
            Assert.Equal(new[] {
                "OfficeArtDggContainer",
                "OfficeArtFDGGBlock",
                "OfficeArtBStoreContainer",
                "OfficeArtFBSE"
            }, drawingGroup.OfficeArtRecords.Select(record => record.RecordTypeName).ToArray());
            LegacyXlsDrawingGroupBlock drawingGroupBlock = Assert.Single(drawingGroup.DrawingGroupBlocks);
            Assert.Equal((uint)2048, drawingGroupBlock.MaxShapeId);
            Assert.Equal((uint)2, drawingGroupBlock.DeclaredIdentifierClusterCount);
            Assert.Equal((uint)2, drawingGroupBlock.SavedShapeCount);
            Assert.Equal((uint)1, drawingGroupBlock.SavedDrawingCount);
            LegacyXlsDrawingIdentifierCluster identifierCluster = Assert.Single(drawingGroupBlock.IdentifierClusters);
            Assert.Equal((uint)1, identifierCluster.DrawingId);
            Assert.Equal((uint)1024, identifierCluster.CurrentShapeId);
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
            LegacyXlsDrawingRecord drawing = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "MsoDrawing");
            Assert.Equal((ushort)0xf002, drawing.EscherRecordType);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer, drawing.EscherRecordTypeKind);
            Assert.Equal("OfficeArtDgContainer", drawing.EscherRecordTypeName);
            Assert.Equal((ushort)1, drawing.EscherRecordInstance);
            Assert.Equal((byte)0x0f, drawing.EscherRecordVersion);
            Assert.Equal((uint)114, drawing.EscherPayloadLength);
            Assert.Equal(new[] {
                "OfficeArtDgContainer",
                "OfficeArtFDG",
                "OfficeArtSpContainer",
                "OfficeArtFSP",
                "OfficeArtFOPT",
                "OfficeArtFClientAnchor",
                "OfficeArtChildAnchor"
            }, drawing.OfficeArtRecords.Select(record => record.RecordTypeName).ToArray());
            LegacyXlsDrawingGroupInfo drawingInfo = Assert.Single(drawing.DrawingGroupInfos);
            Assert.Equal((ushort)1, drawingInfo.DrawingId);
            Assert.Equal((uint)1, drawingInfo.ShapeCount);
            Assert.Equal((uint)1024, drawingInfo.LastShapeId);
            Assert.Equal(2, drawing.ShapeProperties.Count);
            LegacyXlsDrawingShapeProperty simpleProperty = drawing.ShapeProperties[0];
            Assert.Equal(0, simpleProperty.Index);
            Assert.Equal((ushort)0x00bf, simpleProperty.RawOperationId);
            Assert.Equal((ushort)0x00bf, simpleProperty.PropertyId);
            Assert.Equal("PropertyId:0x00BF", simpleProperty.PropertyIdKey);
            Assert.Equal("TextBooleanProperties", simpleProperty.PropertyName);
            Assert.Equal("Text", simpleProperty.PropertyGroupName);
            Assert.False(simpleProperty.IsBlipId);
            Assert.False(simpleProperty.IsComplex);
            Assert.Equal((uint)1, simpleProperty.Value);
            Assert.Null(simpleProperty.DeclaredComplexDataLength);
            Assert.Null(simpleProperty.AvailableComplexDataLength);
            LegacyXlsDrawingShapeProperty complexProperty = drawing.ShapeProperties[1];
            Assert.Equal(1, complexProperty.Index);
            Assert.Equal((ushort)0x8005, complexProperty.RawOperationId);
            Assert.Equal((ushort)0x0005, complexProperty.PropertyId);
            Assert.Equal("PropertyId:0x0005", complexProperty.PropertyName);
            Assert.Equal("Protection", complexProperty.PropertyGroupName);
            Assert.False(complexProperty.IsBlipId);
            Assert.True(complexProperty.IsComplex);
            Assert.Equal((uint)4, complexProperty.Value);
            Assert.Equal((uint)4, complexProperty.DeclaredComplexDataLength);
            Assert.Equal(4, complexProperty.AvailableComplexDataLength);
            Assert.Null(complexProperty.ComplexText);
            LegacyXlsDrawingOfficeArtRecord childAnchorRecord = Assert.Single(drawing.OfficeArtRecords, record => record.RecordTypeKind == LegacyXlsDrawingEscherRecordType.OfficeArtChildAnchor);
            Assert.False(childAnchorRecord.IsContainer);
            Assert.Equal(2, childAnchorRecord.Depth);
            Assert.Equal((uint)16, childAnchorRecord.PayloadLength);
            LegacyXlsDrawingShape shape = Assert.Single(drawing.ShapeEntries);
            Assert.Equal((ushort)0x004b, shape.ShapeType);
            Assert.Equal("PictureFrame", shape.ShapeTypeName);
            Assert.Equal((uint)1024, shape.ShapeId);
            Assert.Equal((uint)0x00000a02, shape.Flags);
            Assert.Equal(new[] { "Child", "HaveAnchor", "HaveShapeType" }, shape.FlagNames);
            LegacyXlsDrawingAnchor anchor = Assert.Single(drawing.AnchorEntries);
            Assert.Equal((ushort)0x0000, anchor.Flags);
            Assert.Equal((ushort)1, anchor.StartColumn);
            Assert.Equal((ushort)10, anchor.StartDx);
            Assert.Equal((ushort)2, anchor.StartRow);
            Assert.Equal((ushort)20, anchor.StartDy);
            Assert.Equal((ushort)3, anchor.EndColumn);
            Assert.Equal((ushort)30, anchor.EndDx);
            Assert.Equal((ushort)4, anchor.EndRow);
            Assert.Equal((ushort)40, anchor.EndDy);
            LegacyXlsDrawingChildAnchor childAnchor = Assert.Single(drawing.ChildAnchorEntries);
            Assert.Equal(100, childAnchor.Left);
            Assert.Equal(200, childAnchor.Top);
            Assert.Equal(700, childAnchor.Right);
            Assert.Equal(900, childAnchor.Bottom);
            Assert.Equal(600, childAnchor.Width);
            Assert.Equal(700, childAnchor.Height);
            LegacyXlsDrawingRecord shapeStream = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "ShapePropsStream");
            Assert.Equal(LegacyXlsDrawingRecordKind.ShapePropertiesStream, shapeStream.Kind);
            Assert.True(shapeStream.HasFutureRecordHeader);
            Assert.Equal((ushort)0x08a3, shapeStream.FutureRecordHeader?.WrappedRecordType);
            Assert.False(shapeStream.FutureRecordHeader?.HasRange);
            Assert.Equal(5, shapeStream.FutureRecordHeader?.StreamByteCount);
            LegacyXlsDrawingRecord textStream = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "TextPropsStream");
            Assert.Equal(LegacyXlsDrawingRecordKind.TextPropertiesStream, textStream.Kind);
            Assert.True(textStream.HasFutureRecordHeader);
            Assert.Equal((ushort)0x08a4, textStream.FutureRecordHeader?.WrappedRecordType);
            Assert.True(textStream.FutureRecordHeader?.HasRange);
            Assert.Equal((ushort)1, textStream.FutureRecordHeader?.FirstRow);
            Assert.Equal((ushort)2, textStream.FutureRecordHeader?.LastRow);
            Assert.Equal((ushort)0, textStream.FutureRecordHeader?.FirstColumn);
            Assert.Equal((ushort)1, textStream.FutureRecordHeader?.LastColumn);
            Assert.Equal(4, textStream.FutureRecordHeader?.StreamByteCount);
            LegacyXlsDrawingRecord richStream = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "RichTextStream");
            Assert.Equal(LegacyXlsDrawingRecordKind.RichTextStream, richStream.Kind);
            Assert.True(richStream.HasFutureRecordHeader);
            Assert.Equal((ushort)0x08a5, richStream.FutureRecordHeader?.WrappedRecordType);
            Assert.False(richStream.FutureRecordHeader?.HasRange);
            Assert.Equal(4, richStream.FutureRecordHeader?.StreamByteCount);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.DetailCode == "Chart:Chart" && record.RecordType == 0x1002);
            Assert.Contains(workbook.Diagnostics, d => d.DetailCode == "Chart:Chart");
            Assert.Contains(workbook.Diagnostics, d => d.DetailCode == "PivotTable:SxView");
            string markdown = report.ToMarkdown();
            Assert.Contains("Preserved feature records: 8", markdown);
            Assert.Contains("Drawing:MsoDrawingGroup", markdown);
            Assert.Contains("Pivot Table Records By Name", markdown);
            Assert.Contains("Chart Records By Name", markdown);
            Assert.Contains("Drawing Records By Name", markdown);
            Assert.Contains("Drawing Records By Object Type", markdown);
            Assert.Contains("Drawing Records By Object Type Name", markdown);
            Assert.Contains("Picture", markdown);
            Assert.Contains("Drawing Records By Object Flags", markdown);
            Assert.Contains("Drawing Records By Object Flag Name", markdown);
            Assert.Contains("Drawing Object Subrecords By Name", markdown);
            Assert.Contains("FtCmo", markdown);
            Assert.Contains("Drawing Future Record Wrapped Types", markdown);
            Assert.Contains("ShapePropsStream\\|0x08A3", markdown);
            Assert.Contains("Drawing Future Record Ranges", markdown);
            Assert.Contains("TextPropsStream\\|A2:B3", markdown);
            Assert.Contains("Drawing Future Record Stream Byte Counts", markdown);
            Assert.Contains("Drawing Records By Escher Record Type", markdown);
            Assert.Contains("Drawing Records By Escher Record Type Name", markdown);
            Assert.Contains("Drawing OfficeArt records: 11", markdown);
            Assert.Contains("Drawing group blocks: 1", markdown);
            Assert.Contains("Drawing group infos: 1", markdown);
            Assert.Contains("Drawing identifier clusters: 1", markdown);
            Assert.Contains("Drawing shape properties: 2", markdown);
            Assert.Contains("Drawing OfficeArt Records By Type Name", markdown);
            Assert.Contains("Drawing Group Blocks By Max Shape Id", markdown);
            Assert.Contains("Drawing Group Infos By Shape Count", markdown);
            Assert.Contains("Drawing Identifier Clusters By Current Shape Id", markdown);
            Assert.Contains("Drawing Shape Properties By Id", markdown);
            Assert.Contains("Drawing BLIP Store Entries By Type", markdown);
            Assert.Contains("Drawing Shape Entries By Type", markdown);
            Assert.Contains("PictureFrame", markdown);
            Assert.Contains("Drawing Anchor Entries By Range", markdown);
            Assert.Contains("Drawing Child Anchor Entries By Rectangle", markdown);
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
        [InlineData(0xF11E, LegacyXlsDrawingEscherRecordType.OfficeArtSplitMenuColorContainer, "OfficeArtSplitMenuColorContainer")]
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
            Assert.Equal(24, workbook.PivotTableRecords.Count);
            Assert.Equal(24, report.PivotTableRecordCount);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.View]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Field]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Item]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.DataItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Cache]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.CacheStream]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.CacheSource]);
            Assert.Equal(8, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.CacheItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Formula]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Table]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.GroupingRange]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Filter]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Format]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.ExtendedPivotField]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.PivotChart]);
            Assert.Equal(2, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Additional]);
            Assert.DoesNotContain(report.PivotTableRecordsByKind, entry => entry.Key == LegacyXlsPivotTableRecordKind.PreserveOnly);
            Assert.Equal(1, report.PivotTableRecordsByName["SxView"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvd"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvi"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxdi"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxdb"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxStreamId"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvs"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxFormula"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxnum"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxbool"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxerr"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxstring"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxint"]);
            Assert.Equal(2, report.PivotTableRecordsByName["Sxdtr"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxnil"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxtbl"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxRng"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxFilt"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxFormat"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxVdEx"]);
            Assert.Equal(1, report.PivotTableRecordsByName["PivotChartBits"]);
            Assert.Equal(2, report.PivotTableRecordsByName["SxAddl"]);
            Assert.Equal(13, report.PivotTableRecordsByLocation["(workbook)"]);
            Assert.Equal(11, report.PivotTableRecordsByLocation["PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByKindAndLocation["View|(workbook)"]);
            Assert.Equal(1, report.PivotTableRecordsByKindAndLocation["Formula|(workbook)"]);
            Assert.Equal(5, report.PivotTableRecordsByKindAndLocation["CacheItem|(workbook)"]);
            Assert.Equal(3, report.PivotTableRecordsByKindAndLocation["CacheItem|PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByKindAndLocation["Field|PivotMeta"]);
            Assert.Equal(2, report.PivotTableRecordsByKindAndLocation["Additional|PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByNameAndLocation["SxView|(workbook)"]);
            Assert.Equal(1, report.PivotTableRecordsByNameAndLocation["SxFormula|(workbook)"]);
            Assert.Equal(2, report.PivotTableRecordsByNameAndLocation["SxAddl|PivotMeta"]);
            Assert.Equal(1, report.PivotTableWorkbookStates["View:Present|Cache:Present|CacheSource:Present|CacheItems:Present|Fields:Present|Items:Present|DataItems:Present|Grouping:Present|Formulas:Present|Additional:Present|Locations:WorkbookAndSheets"]);
            Assert.Equal(1, report.PivotTableFormulaPayloadLengths["SxFormula|Bytes:4"]);
            Assert.Equal(1, report.PivotTableCacheItemKinds["Number"]);
            Assert.Equal(1, report.PivotTableCacheItemKinds["Integer"]);
            Assert.Equal(1, report.PivotTableCacheItemKinds["Boolean"]);
            Assert.Equal(1, report.PivotTableCacheItemKinds["Error"]);
            Assert.Equal(1, report.PivotTableCacheItemKinds["String"]);
            Assert.Equal(2, report.PivotTableCacheItemKinds["DateTime"]);
            Assert.Equal(1, report.PivotTableCacheItemKinds["Empty"]);
            Assert.Equal(7, report.PivotTableCacheItemValueStates["HasValue"]);
            Assert.Equal(1, report.PivotTableCacheItemValueStates["Empty"]);
            Assert.Equal(1, report.PivotTableCacheItemStringLengths["Characters:4"]);
            Assert.Equal(1, report.PivotTableCacheItemErrorCodes["ErrorCode:0x07"]);
            Assert.Equal(1, report.PivotTableCacheItemBooleanValues["True"]);
            Assert.Equal(1, report.PivotTableCacheStreamNames["Sxdb|0001"]);
            Assert.Equal(1, report.PivotTableCacheStreamNames["SxStreamId|0001"]);
            Assert.Equal(1, report.PivotTableCacheSourceTypes["Sxdb|Sheet"]);
            Assert.Equal(1, report.PivotTableCacheSourceTypes["Sxvs|Sheet"]);
            Assert.Equal(1, report.PivotTableCacheRecordCounts["Records:12"]);
            Assert.Equal(1, report.PivotTableCacheFieldCounts["SourceFields:3;TotalFields:4"]);
            Assert.Equal(1, report.PivotTableCacheUsedRecordCounts["UsedRecords:10"]);
            Assert.Equal(1, report.PivotTableCachePropertyFlags["HasRecords:True"]);
            Assert.Equal(1, report.PivotTableCachePropertyFlags["Invalid:False"]);
            Assert.Equal(1, report.PivotTableCachePropertyFlags["RefreshOnLoad:True"]);
            Assert.Equal(1, report.PivotTableCachePropertyFlags["OptimizeMemory:False"]);
            Assert.Equal(1, report.PivotTableCachePropertyFlags["BackgroundQuery:False"]);
            Assert.Equal(1, report.PivotTableCachePropertyFlags["EnableRefresh:True"]);
            Assert.Equal(1, report.PivotTableCacheRefreshUserStates["HasRefreshUser"]);
            Assert.Equal(1, report.PivotTableDataItemAggregations["AggregationFunction:0"]);
            Assert.Equal(1, report.PivotTableDataItemAggregationKinds["Sum"]);
            Assert.Equal(1, report.PivotTableDataItemFieldIndexes["FieldIndex:2"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculations["PercentOfGrandTotal"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculationFieldIndexes["FieldIndex:-1"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculationItemIndexes["ItemIndex:-1"]);
            Assert.Equal(1, report.PivotTableDataItemNumberFormats["NumberFormatId:14"]);
            Assert.Equal(1, report.PivotTableDataItemNames["Sales"]);
            Assert.Equal(1, report.PivotTableGroupingKinds["Months"]);
            Assert.Equal(1, report.PivotTableGroupingBoundaryStates["AutoStart:True;AutoEnd:True"]);
            Assert.Equal(1, report.PivotTableGroupingDateRanges["Start:2024-01-01 00:00:00;End:2024-12-31 00:00:00;Interval:1"]);
            Assert.Equal(1, report.PivotTableFormulaScopes["AllCacheFields"]);
            Assert.Equal(1, report.PivotTableFormulaCacheFieldIndexes["CacheField:-1"]);
            Assert.Equal(1, report.PivotTableFormulaReservedValues["Reserved:0x0000"]);
            Assert.Equal(1, report.PivotTableExtendedFieldStates["ShowAllItems:True"]);
            Assert.Equal(1, report.PivotTableExtendedFieldStates["CanDragToRow:True"]);
            Assert.Equal(1, report.PivotTableExtendedFieldStates["CanDragToColumn:True"]);
            Assert.Equal(1, report.PivotTableExtendedFieldStates["CanDragToPage:True"]);
            Assert.Equal(1, report.PivotTableExtendedFieldStates["CanDragToHide:True"]);
            Assert.Equal(1, report.PivotTableExtendedFieldStates["PreventDragToData:False"]);
            Assert.Equal(1, report.PivotTableExtendedFieldStates["ServerBased:True"]);
            Assert.Equal(2, report.PivotTableAdditionalClasses["SxcCache"]);
            Assert.Equal(1, report.PivotTableAdditionalTypes["SXDId"]);
            Assert.Equal(1, report.PivotTableAdditionalTypes["SXDEnd"]);
            Assert.Equal(1, report.PivotTableAdditionalClassTypes["SxcCache|SXDId"]);
            Assert.Equal(1, report.PivotTableAdditionalClassTypes["SxcCache|SXDEnd"]);
            Assert.Equal(1, report.PivotTableAdditionalCacheIds["CacheId:1"]);
            Assert.Equal(1, report.PivotTableAdditionalClassDepthsBefore["Depth:0"]);
            Assert.Equal(1, report.PivotTableAdditionalClassDepthsBefore["Depth:1"]);
            Assert.Equal(1, report.PivotTableAdditionalClassDepthsAfter["Depth:1"]);
            Assert.Equal(1, report.PivotTableAdditionalClassDepthsAfter["Depth:0"]);
            Assert.Equal(1, report.PivotTableAdditionalClassTransitions["BeginClass"]);
            Assert.Equal(1, report.PivotTableAdditionalClassTransitions["EndClass"]);
            Assert.Equal(1, report.PivotTableAdditionalClassTransitionsByClassType["SxcCache|SXDId|BeginClass"]);
            Assert.Equal(1, report.PivotTableAdditionalClassTransitionsByClassType["SxcCache|SXDEnd|EndClass"]);

            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.View && record.RecordName == "SxView" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Field && record.RecordName == "Sxvd" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Item && record.RecordName == "Sxvi" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Cache && record.RecordName == "Sxdb" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.CacheStream && record.RecordName == "SxStreamId" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.CacheSource && record.RecordName == "Sxvs" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Formula && record.RecordName == "SxFormula" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.CacheItem && record.RecordName == "Sxnum" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Table && record.RecordName == "Sxtbl" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Filter && record.RecordName == "SxFilt" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Format && record.RecordName == "SxFormat" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.PivotChart && record.RecordName == "PivotChartBits" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Additional && record.RecordName == "SxAddl" && record.SheetName == "PivotMeta");

            LegacyXlsPivotTableRecord numberItem = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxnum");
            Assert.Equal(LegacyXlsPivotCacheItemKind.Number, numberItem.CacheItemKind);
            Assert.Equal("Number", numberItem.CacheItemKindName);
            Assert.Equal(42.5d, numberItem.CacheItemNumericValue);

            LegacyXlsPivotTableRecord booleanItem = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxbool");
            Assert.Equal(LegacyXlsPivotCacheItemKind.Boolean, booleanItem.CacheItemKind);
            Assert.True(booleanItem.CacheItemBooleanValue);

            LegacyXlsPivotTableRecord errorItem = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxerr");
            Assert.Equal(LegacyXlsPivotCacheItemKind.Error, errorItem.CacheItemKind);
            Assert.Equal((ushort)0x0007, errorItem.CacheItemErrorCode);
            Assert.Equal("#DIV/0!", errorItem.CacheItemErrorText);

            LegacyXlsPivotTableRecord stringItem = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxstring");
            Assert.Equal(LegacyXlsPivotCacheItemKind.String, stringItem.CacheItemKind);
            Assert.Equal("East", stringItem.CacheItemStringValue);

            LegacyXlsPivotTableRecord integerItem = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxint");
            Assert.Equal(LegacyXlsPivotCacheItemKind.Integer, integerItem.CacheItemKind);
            Assert.Equal((short)1, integerItem.CacheItemIntegerValue);

            LegacyXlsPivotTableRecord[] dateItems = workbook.PivotTableRecords.Where(record => record.RecordName == "Sxdtr").ToArray();
            Assert.All(dateItems, record => Assert.Equal(LegacyXlsPivotCacheItemKind.DateTime, record.CacheItemKind));
            Assert.Equal("2024-01-01 00:00:00", dateItems[0].CacheItemDateTimeValue?.ToString());
            Assert.Equal("2024-12-31 00:00:00", dateItems[1].CacheItemDateTimeValue?.ToString());

            LegacyXlsPivotTableRecord emptyItem = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxnil");
            Assert.Equal(LegacyXlsPivotCacheItemKind.Empty, emptyItem.CacheItemKind);
            Assert.True(emptyItem.IsEmptyCacheItem);

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
            Assert.Equal((short)-1, dataItem.DisplayCalculationFieldIndex);
            Assert.Equal((short)-1, dataItem.DisplayCalculationItemIndex);
            Assert.Equal((ushort)14, dataItem.NumberFormatId);
            Assert.Equal("Sales", dataItem.Name);

            LegacyXlsPivotTableRecord cache = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxdb");
            Assert.Equal((ushort)1, cache.CacheStreamId);
            Assert.Equal("0001", cache.CacheStreamName);
            Assert.Equal(12, cache.CacheRecordCount);
            Assert.True(cache.CacheHasRecords);
            Assert.False(cache.CacheInvalid);
            Assert.True(cache.CacheRefreshOnLoad);
            Assert.False(cache.CacheOptimizeMemory);
            Assert.False(cache.CacheBackgroundQuery);
            Assert.True(cache.CacheEnableRefresh);
            Assert.Equal((short)3, cache.CacheSourceFieldCount);
            Assert.Equal((short)4, cache.CacheTotalFieldCount);
            Assert.Equal((ushort)10, cache.CacheUsedRecordCount);
            Assert.Equal((ushort)1, cache.CacheSourceType);
            Assert.Equal(LegacyXlsPivotCacheSourceType.Sheet, cache.CacheSourceTypeKind);
            Assert.Equal("Sheet", cache.CacheSourceTypeName);
            Assert.Equal("Excel", cache.CacheRefreshedBy);

            LegacyXlsPivotTableRecord cacheStream = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "SxStreamId");
            Assert.Equal((ushort)1, cacheStream.CacheStreamId);
            Assert.Equal("0001", cacheStream.CacheStreamName);

            LegacyXlsPivotTableRecord cacheSource = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "Sxvs");
            Assert.Equal((ushort)1, cacheSource.CacheSourceType);
            Assert.Equal(LegacyXlsPivotCacheSourceType.Sheet, cacheSource.CacheSourceTypeKind);
            Assert.Equal("Sheet", cacheSource.CacheSourceTypeName);

            LegacyXlsPivotTableRecord formulaScope = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "SxFormula");
            Assert.Equal(LegacyXlsPivotTableRecordKind.Formula, formulaScope.Kind);
            Assert.Equal((ushort)0, formulaScope.CalculatedItemFormulaReserved);
            Assert.Equal((short)-1, formulaScope.CalculatedItemFormulaCacheFieldIndex);
            Assert.True(formulaScope.HasCalculatedItemFormulaScope);
            Assert.True(formulaScope.CalculatedItemFormulaAppliesToAllCacheFields);
            Assert.Equal("AllCacheFields", formulaScope.CalculatedItemFormulaScopeName);

            LegacyXlsPivotTableRecord grouping = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.GroupingRange);
            Assert.Equal("PivotMeta", grouping.SheetName);
            Assert.Equal("SxRng", grouping.RecordName);
            Assert.True(grouping.AutoStart);
            Assert.True(grouping.AutoEnd);
            Assert.Equal(LegacyXlsPivotGroupingKind.Months, grouping.GroupingKind);
            Assert.Equal("2024-01-01 00:00:00", grouping.GroupingDateStart?.ToString());
            Assert.Equal("2024-12-31 00:00:00", grouping.GroupingDateEnd?.ToString());
            Assert.Equal((short)1, grouping.GroupingDateInterval);

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

            LegacyXlsPivotTableRecord additional = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Additional && record.AdditionalType == 0x00);
            Assert.Equal("PivotMeta", additional.SheetName);
            Assert.Equal("SxAddl", additional.RecordName);
            Assert.Equal((ushort)0x0864, additional.AdditionalFutureRecordType);
            Assert.Equal((ushort)0, additional.AdditionalFutureFlags);
            Assert.Equal((byte)0x03, additional.AdditionalClass);
            Assert.Equal("SxcCache", additional.AdditionalClassName);
            Assert.Equal((byte)0x00, additional.AdditionalType);
            Assert.Equal("SXDId", additional.AdditionalTypeName);
            Assert.Equal(1U, additional.AdditionalCacheId);
            Assert.Equal(1, additional.AdditionalSequenceIndex);
            Assert.Equal(0, additional.AdditionalClassDepthBefore);
            Assert.Equal(1, additional.AdditionalClassDepthAfter);
            Assert.Equal("BeginClass", additional.AdditionalClassTransition);

            LegacyXlsPivotTableRecord additionalEnd = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Additional && record.AdditionalType == 0xff);
            Assert.Equal("PivotMeta", additionalEnd.SheetName);
            Assert.Equal("SxAddl", additionalEnd.RecordName);
            Assert.Equal((ushort)0x0864, additionalEnd.AdditionalFutureRecordType);
            Assert.Equal((ushort)0, additionalEnd.AdditionalFutureFlags);
            Assert.Equal((byte)0x03, additionalEnd.AdditionalClass);
            Assert.Equal("SxcCache", additionalEnd.AdditionalClassName);
            Assert.Equal("SXDEnd", additionalEnd.AdditionalTypeName);
            Assert.Null(additionalEnd.AdditionalCacheId);
            Assert.Equal(2, additionalEnd.AdditionalSequenceIndex);
            Assert.Equal(1, additionalEnd.AdditionalClassDepthBefore);
            Assert.Equal(0, additionalEnd.AdditionalClassDepthAfter);
            Assert.Equal("EndClass", additionalEnd.AdditionalClassTransition);
            Assert.Equal(24, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(24, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Pivot table records: 24", markdown);
            Assert.Contains("Pivot Table Records By Kind", markdown);
            Assert.Contains("Pivot Table Records By Location", markdown);
            Assert.Contains("Pivot Table Records By Kind And Location", markdown);
            Assert.Contains("CacheItem\\|(workbook)", markdown);
            Assert.Contains("Pivot Table Records By Name And Location", markdown);
            Assert.Contains("SxAddl\\|PivotMeta", markdown);
            Assert.Contains("Pivot Table Workbook States", markdown);
            Assert.Contains("View:Present\\|Cache:Present\\|CacheSource:Present", markdown);
            Assert.Contains("Pivot Table Formula Payload Lengths", markdown);
            Assert.Contains("Pivot Table Cache Item Kinds", markdown);
            Assert.Contains("Pivot Table Cache Stream Names", markdown);
            Assert.Contains("SxStreamId\\|0001", markdown);
            Assert.Contains("Pivot Table Cache Source Types", markdown);
            Assert.Contains("Sxvs\\|Sheet", markdown);
            Assert.Contains("Pivot Table Cache Record Counts", markdown);
            Assert.Contains("Records:12", markdown);
            Assert.Contains("Pivot Table Cache Field Counts", markdown);
            Assert.Contains("SourceFields:3;TotalFields:4", markdown);
            Assert.Contains("Pivot Table Cache Used Record Counts", markdown);
            Assert.Contains("UsedRecords:10", markdown);
            Assert.Contains("Pivot Table Cache Property Flags", markdown);
            Assert.Contains("RefreshOnLoad:True", markdown);
            Assert.Contains("Pivot Table Cache Refresh User States", markdown);
            Assert.Contains("HasRefreshUser", markdown);
            Assert.Contains("Pivot Table Data Item Aggregations", markdown);
            Assert.Contains("Pivot Table Data Item Aggregation Kinds", markdown);
            Assert.Contains("Pivot Table Data Item Field Indexes", markdown);
            Assert.Contains("PercentOfGrandTotal", markdown);
            Assert.Contains("Pivot Table Data Item Display Calculation Field Indexes", markdown);
            Assert.Contains("Pivot Table Data Item Display Calculation Item Indexes", markdown);
            Assert.Contains("Pivot Table Data Item Number Formats", markdown);
            Assert.Contains("Pivot Table Data Item Names", markdown);
            Assert.Contains("Pivot Table Grouping Kinds", markdown);
            Assert.Contains("Pivot Table Grouping Boundary States", markdown);
            Assert.Contains("Pivot Table Grouping Date Ranges", markdown);
            Assert.Contains("Pivot Table Formula Scopes", markdown);
            Assert.Contains("Pivot Table Extended Field States", markdown);
            Assert.Contains("Pivot Table Additional Classes", markdown);
            Assert.Contains("SxcCache", markdown);
            Assert.Contains("Pivot Table Additional Types", markdown);
            Assert.Contains("SXDId", markdown);
            Assert.Contains("Pivot Table Additional Class Types", markdown);
            Assert.Contains("Pivot Table Additional Cache Ids", markdown);
            Assert.Contains("Pivot Table Additional Class Depths Before", markdown);
            Assert.Contains("Pivot Table Additional Class Depths After", markdown);
            Assert.Contains("Pivot Table Additional Class Transitions", markdown);
            Assert.Contains("Pivot Table Additional Class Transitions By Class Type", markdown);
            Assert.Contains("CacheId:1", markdown);
            Assert.Contains("SxcCache\\|SXDEnd\\|EndClass", markdown);
            Assert.Contains("SxVdEx", markdown);
            Assert.Contains("SxAddl", markdown);
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
            Assert.Equal(43, unsupportedSheet.ChartRecordCount);
            Assert.Equal(4, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Container]);
            Assert.Equal(10, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Axis]);
            Assert.Equal(3, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Series]);
            Assert.Equal(9, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Formatting]);
            Assert.Equal(3, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Layout]);
            Assert.Equal(5, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.ChartType]);
            Assert.Equal(6, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Text]);
            Assert.Equal(3, unsupportedSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.FutureMetadata]);
            Assert.Equal(1, unsupportedSheet.ChartRecordsByChartType["Scatter"]);
            Assert.Equal(44, report.UnsupportedFeatureCount);
            Assert.Equal(43, report.PreservedFeatureRecordCount);
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.ChartSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x02|ChartSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["ChartOnly"]);
            Assert.Equal(1, report.UnsupportedSheetMetadataRecordCount);
            Assert.Equal(1, report.UnsupportedSheetMetadataRecordsByKind[LegacyXlsUnsupportedSheetMetadataKind.ChartTextObject]);
            Assert.Equal(1, report.UnsupportedChartSheetTextObjectCounts["TextObjects:1"]);
            Assert.Equal(1, report.UnsupportedChartSheetChartRecordCounts["ChartRecords:43"]);
            Assert.Equal(1, report.UnsupportedChartSheetChartRecordCountsBySheet["Sheet:ChartOnly;ChartRecords:43"]);
            Assert.Equal(4, report.UnsupportedChartSheetChartRecordKinds["Container"]);
            Assert.Equal(10, report.UnsupportedChartSheetChartRecordKinds["Axis"]);
            Assert.Equal(3, report.UnsupportedChartSheetChartRecordKinds["Series"]);
            Assert.Equal(9, report.UnsupportedChartSheetChartRecordKinds["Formatting"]);
            Assert.Equal(3, report.UnsupportedChartSheetChartRecordKinds["Layout"]);
            Assert.Equal(5, report.UnsupportedChartSheetChartRecordKinds["ChartType"]);
            Assert.Equal(6, report.UnsupportedChartSheetChartRecordKinds["Text"]);
            Assert.Equal(3, report.UnsupportedChartSheetChartRecordKinds["FutureMetadata"]);
            Assert.Equal(1, report.UnsupportedChartSheetChartTypes["Scatter"]);
            Assert.Equal(10, report.UnsupportedChartSheetChartRecordKindsBySheet["Sheet:ChartOnly;Kind:Axis"]);
            Assert.Equal(5, report.UnsupportedChartSheetChartRecordKindsBySheet["Sheet:ChartOnly;Kind:ChartType"]);
            Assert.Equal(1, report.UnsupportedChartSheetChartTypesBySheet["Sheet:ChartOnly;ChartType:Scatter"]);
            Assert.Equal(1, report.UnsupportedChartSheetStates["PrintSize:Missing|TextObjects:Present|ChartRecords:Present|ChartTypes:Present"]);
            Assert.Empty(report.UnsupportedChartSheetPrintSizes);
            Assert.Empty(report.UnsupportedChartSheetPrintSizeKinds);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.ChartSheet]);
            Assert.Equal(43, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(43, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.Chart]);
            Assert.Equal(43, report.ChartRecordCount);
            Assert.Equal(4, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Container]);
            Assert.Equal(10, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Axis]);
            Assert.Equal(3, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Series]);
            Assert.Equal(9, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Formatting]);
            Assert.Equal(3, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Layout]);
            Assert.Equal(5, report.ChartRecordsByKind[LegacyXlsChartRecordKind.ChartType]);
            Assert.Equal(6, report.ChartRecordsByKind[LegacyXlsChartRecordKind.Text]);
            Assert.Equal(3, report.ChartRecordsByKind[LegacyXlsChartRecordKind.FutureMetadata]);
            Assert.Equal(1, report.ChartRecordsByName["ChartFrtInfo"]);
            Assert.Equal(1, report.ChartRecordsByName["Begin"]);
            Assert.Equal(1, report.ChartRecordsByName["StartBlock"]);
            Assert.Equal(1, report.ChartRecordsByName["CatLab"]);
            Assert.Equal(1, report.ChartRecordsByName["EndBlock"]);
            Assert.Equal(1, report.ChartRecordsByName["Units"]);
            Assert.Equal(1, report.ChartRecordsByName["Chart"]);
            Assert.Equal(1, report.ChartRecordsByName["DataFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["Ifmt"]);
            Assert.Equal(1, report.ChartRecordsByName["ChartFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["Axis"]);
            Assert.Equal(1, report.ChartRecordsByName["ValueRange"]);
            Assert.Equal(1, report.ChartRecordsByName["CatSerRange"]);
            Assert.Equal(1, report.ChartRecordsByName["Tick"]);
            Assert.Equal(1, report.ChartRecordsByName["AxesUsed"]);
            Assert.Equal(1, report.ChartRecordsByName["Series"]);
            Assert.Equal(1, report.ChartRecordsByName["SIIndex"]);
            Assert.Equal(1, report.ChartRecordsByName["Scatter"]);
            Assert.Equal(1, report.ChartRecordsByName["LineFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["Frame"]);
            Assert.Equal(1, report.ChartRecordsByName["AreaFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["MarkerFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["PieFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["AttachedLabel"]);
            Assert.Equal(1, report.ChartRecordsByName["DefaultText"]);
            Assert.Equal(1, report.ChartRecordsByName["Text"]);
            Assert.Equal(1, report.ChartRecordsByName["FontX"]);
            Assert.Equal(1, report.ChartRecordsByName["ObjectLink"]);
            Assert.Equal(1, report.ChartRecordsByName["Legend"]);
            Assert.Equal(3, report.ChartRecordsByName["AxisLineFormat"]);
            Assert.Equal(1, report.ChartRecordsByName["AxcExt"]);
            Assert.Equal(1, report.ChartRecordsByName["Dat"]);
            Assert.Equal(1, report.ChartRecordsByName["Pos"]);
            Assert.Equal(1, report.ChartRecordsByName["BRAI"]);
            Assert.Equal(1, report.ChartRecordsByName["PlotGrowth"]);
            Assert.Equal(1, report.ChartRecordsByName["GelFrame"]);
            Assert.Equal(1, report.ChartRecordsByName["BopPopCustom"]);
            Assert.Equal(1, report.ChartRecordsByName["Fbi2"]);
            Assert.Equal(1, report.ChartRecordsByName["Chart3d"]);
            Assert.Equal(1, report.ChartRecordsByName["Chart3DBarShape"]);
            Assert.Equal(1, report.ChartRecordsByName["End"]);
            Assert.Equal(1, report.ChartRecordsByNameAndPayloadLength["Pos|Bytes:20"]);
            Assert.Equal(1, report.ChartRecordsByNameAndPayloadLength["BRAI|Bytes:17"]);
            Assert.Equal(1, report.ChartWorkbookStates["Containers:Present|ChartTypes:Present|Series:Present|Axes:Present|Text:Present|Formatting:Present|Layout:Present|Future:Present|PreserveOnly:Missing|Scopes:ChartSheetsOnly"]);
            Assert.Equal(1, report.ChartRecordsByContainerDepthBefore["Depth:0"]);
            Assert.Equal(42, report.ChartRecordsByContainerDepthBefore["Depth:1"]);
            Assert.Equal(42, report.ChartRecordsByContainerDepthAfter["Depth:1"]);
            Assert.Equal(1, report.ChartRecordsByContainerDepthAfter["Depth:0"]);
            Assert.Equal(1, report.ChartRecordsByContainerTransition["Begin"]);
            Assert.Equal(41, report.ChartRecordsByContainerTransition["InsideContainer"]);
            Assert.Equal(1, report.ChartRecordsByContainerTransition["End"]);
            Assert.Equal(1, report.ChartRecordsByNameAndContainerDepth["Begin|Depth:0"]);
            Assert.Equal(1, report.ChartRecordsByNameAndContainerDepth["Chart|Depth:1"]);
            Assert.Equal(1, report.ChartRecordsByNameAndContainerDepth["End|Depth:1"]);
            Assert.Equal(1, report.ChartRecordsByNameAndContainerTransition["Begin|Begin"]);
            Assert.Equal(1, report.ChartRecordsByNameAndContainerTransition["Chart|InsideContainer"]);
            Assert.Equal(1, report.ChartRecordsByNameAndContainerTransition["End|End"]);
            Assert.Equal(1, report.ChartRecordsByChartType["Scatter"]);
            Assert.Equal(1, report.ChartRecordsByChartType["ThreeDimensional"]);
            Assert.Equal(1, report.ChartRecordsByChartType["ThreeDimensionalBarShape"]);
            Assert.Equal(1, report.ChartRecordsByRectangle["X:100;Y:200;Width:3000;Height:2200"]);
            Assert.Equal(1, report.ChartRecordsByAxisType["ValueOrVerticalValue"]);
            Assert.Equal(1, report.ChartRecordsByAxesUsedCount["AxesUsed:1"]);
            Assert.Equal(1, report.ChartCategorySeriesRangeIntervals["Cross:2;Labels:3;Ticks:4"]);
            Assert.Equal(1, report.ChartCategorySeriesRangeStates["Between:True;MaxCross:False;Reversed:True"]);
            Assert.Equal(1, report.ChartAxisLineFormatTargets["AxisLine"]);
            Assert.Equal(1, report.ChartAxisLineFormatTargets["MajorGridlines"]);
            Assert.Equal(1, report.ChartAxisLineFormatTargets["MinorGridlines"]);
            Assert.Equal(1, report.ChartSeriesCategoryDataTypes["Text"]);
            Assert.Equal(1, report.ChartSeriesValueDataTypes["Numeric"]);
            Assert.Equal(1, report.ChartSeriesBubbleSizeDataTypes["Numeric"]);
            Assert.Equal(1, report.ChartSeriesValueCounts["Categories:4;Values:4;BubbleSizes:0"]);
            Assert.Equal(1, report.ChartSeriesDataCacheIndexes["Index:2"]);
            Assert.Equal(1, report.ChartSeriesDataCacheTypes["CategoryLabelsOrHorizontalValues"]);
            Assert.Equal(1, report.ChartDataSourceIds["ValuesOrHorizontalValues"]);
            Assert.Equal(1, report.ChartDataSourceReferenceTypes["WorksheetRange"]);
            Assert.Equal(1, report.ChartDataSourceNumberFormatIds["NumberFormatId:14"]);
            Assert.Equal(1, report.ChartDataSourceFormulaByteCounts["FormulaBytes:9"]);
            Assert.Equal(1, report.ChartDataSourceFormulaProjectionStates["FormulaTextProjected"]);
            Assert.Equal(1, report.ChartDataSourceStates["Source:ValuesOrHorizontalValues;Reference:WorksheetRange;CustomNumberFormat:True;FormulaBytes:9;FormulaComplete:True;FormulaTextProjected:True"]);
            Assert.Equal(1, report.FormulaTokensByContext["ChartDataSource"]);
            Assert.Equal(1, report.FormulaTokensByContextAndSheet["ChartDataSource|ChartOnly"]);
            Assert.Equal(1, report.FormulaTokensByName["PtgArea"]);
            Assert.Equal(1, report.ChartDataFormatTargets["Series"]);
            Assert.Equal(1, report.ChartDataFormatSeriesIndexes["SeriesIndex:2"]);
            Assert.Equal(1, report.ChartNumberFormatIds["NumberFormatId:14"]);
            Assert.Equal(1, report.ChartFontIndexes["FontIndex:3"]);
            Assert.Equal(1, report.ChartDataTableOptions["HorizontalBorders:True;VerticalBorders:False;Outline:True;SeriesKeys:True"]);
            Assert.Equal(1, report.ChartThreeDimensionalBarShapeRisers["Ellipse"]);
            Assert.Equal(1, report.ChartThreeDimensionalBarShapeTapers["ProjectedPoint"]);
            Assert.Equal(1, report.ChartThreeDimensionalBarShapeStates["Riser:Ellipse;Taper:ProjectedPoint"]);
            Assert.Equal(1, report.ChartScatterBubbleSizeRatios["Ratio:150"]);
            Assert.Equal(1, report.ChartScatterBubbleSizeRepresentations["Width"]);
            Assert.Equal(1, report.ChartScatterBubbleSizeRatioStates["Valid"]);
            Assert.Equal(1, report.ChartScatterStates["Bubble:True;NegativeBubbles:True;Shadow:True;Size:Width"]);
            Assert.Equal(1, report.ChartLineFormatStyles["Dash"]);
            Assert.Equal(1, report.ChartLineFormatWeights["Medium"]);
            Assert.Equal(1, report.ChartAreaFormatPatterns["Solid"]);
            Assert.Equal(1, report.ChartMarkerFormatTypes["Circle"]);
            Assert.Equal(1, report.ChartMarkerFormatSizes["SizeTwips:240"]);
            Assert.Equal(1, report.ChartPieFormatExplosions["ExplosionPercent:25"]);
            Assert.Equal(1, report.ChartAttachedLabelFlags["ShowValue"]);
            Assert.Equal(1, report.ChartAttachedLabelFlags["ShowLabel"]);
            Assert.Equal(1, report.ChartAttachedLabelFlags["ShowSeriesName"]);
            Assert.Equal(1, report.ChartAttachedLabelStates["ShowValue:True;ShowPercent:False;ShowLabelAndPercent:False;ShowLabel:True;ShowBubbleSizes:False;ShowSeriesName:True"]);
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
            Assert.Equal(1, report.ChartValueRangeScales["Min:0;Max:100;Major:25;Minor:5;Cross:10"]);
            Assert.Equal(1, report.ChartValueRangeStates["AutoMin:False;AutoMax:False;AutoMajor:False;AutoMinor:False;AutoCross:False;Log:True;Reversed:True;MaxCross:False"]);
            Assert.Equal(1, report.ChartPositionModePairs["MDCHART/MDABS"]);
            Assert.Equal(1, report.ChartPositionRectangles["X1:15;Y1:25;X2:300;Y2:120"]);
            Assert.Equal(1, report.ChartPositionSemanticTypes["LegendManualSize"]);
            Assert.Equal(1, report.ChartPositionCoordinateMeanings["X1Y1:ChartAreaSprcOffset;X2Y2:PointSize"]);
            Assert.Equal(1, report.ChartPositionIgnoredCoordinateStates["None"]);
            Assert.Equal(1, report.ChartPositionKnownSemanticStates["Known:True"]);
            Assert.Equal(1, report.ChartFrameTypes["ShadowFrame"]);
            Assert.Equal(1, report.ChartFrameAutoStates["AutoSize:True;AutoPosition:True"]);
            Assert.Equal(1, report.ChartPlotGrowthFactors["Horizontal:1.25;Vertical:2.5"]);
            Assert.Equal(43, report.ChartRecordsByLocation["ChartOnly"]);
            Assert.Equal(1, report.DrawingRecordCount);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.TextObject]);
            Assert.Equal(1, report.DrawingRecordsByName["TxO"]);
            Assert.Equal(1, report.DrawingRecordsByLocation["ChartOnly"]);
            Assert.Equal(43, report.UnsupportedFeaturesByLocation["XLS-BIFF-FEATURE-CHART-UNSUPPORTED|ChartOnly"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Begin"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFrtInfo"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:StartBlock"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:CatLab"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:EndBlock"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Units"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DataFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Ifmt"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:LineFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Frame"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AreaFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:MarkerFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:PieFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AttachedLabel"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Axis"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ValueRange"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:CatSerRange"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Tick"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxesUsed"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Series"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:SIIndex"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Scatter"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DefaultText"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Text"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:FontX"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ObjectLink"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Legend"]);
            Assert.Equal(3, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxisLineFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxcExt"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Dat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Pos"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:BRAI"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:PlotGrowth"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:GelFrame"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:BopPopCustom"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Fbi2"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart3d"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart3DBarShape"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:End"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Begin"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFrtInfo"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:StartBlock"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:CatLab"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:EndBlock"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Units"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DataFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Ifmt"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:LineFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Frame"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AreaFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:MarkerFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:PieFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AttachedLabel"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ChartFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Axis"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ValueRange"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:CatSerRange"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Tick"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxesUsed"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Series"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:SIIndex"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Scatter"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:DefaultText"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Text"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:FontX"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:ObjectLink"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Legend"]);
            Assert.Equal(3, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxisLineFormat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:AxcExt"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Dat"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Pos"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:BRAI"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:PlotGrowth"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:GelFrame"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:BopPopCustom"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Fbi2"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart3d"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart3DBarShape"]);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:End"]);
            Assert.Contains(workbook.PreservedFeatureRecords, record => record.SheetName == "ChartOnly" && record.DetailCode == "Chart:Chart");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Begin" && record.ContainerDepthBefore == 0 && record.ContainerDepthAfter == 1 && record.ContainerTransition == "Begin");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "ChartFrtInfo" && record.Kind == LegacyXlsChartRecordKind.FutureMetadata);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "StartBlock" && record.Kind == LegacyXlsChartRecordKind.FutureMetadata);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "CatLab" && record.Kind == LegacyXlsChartRecordKind.Axis);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "EndBlock" && record.Kind == LegacyXlsChartRecordKind.FutureMetadata);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Chart" && record.ChartX == 100 && record.ChartY == 200 && record.ChartWidth == 3000 && record.ChartHeight == 2200);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Axis" && record.AxisType == 0x0001 && record.AxisTypeName == "ValueOrVerticalValue");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "ValueRange" && record.ValueRange != null && record.ValueRange.Minimum == 0 && record.ValueRange.Maximum == 100 && record.ValueRange.MajorUnit == 25 && record.ValueRange.MinorUnit == 5 && record.ValueRange.CrossingValue == 10 && record.ValueRange.Flags == 0x0060 && !record.ValueRange.AutoMinimum && !record.ValueRange.AutoMaximum && !record.ValueRange.AutoMajorUnit && !record.ValueRange.AutoMinorUnit && !record.ValueRange.AutoCrossingValue && record.ValueRange.LogarithmicScale && record.ValueRange.Reversed && !record.ValueRange.MaximumCrossing);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "CatSerRange" && record.CategorySeriesRange != null && record.CategorySeriesRange.CrossingCategory == 2 && record.CategorySeriesRange.LabelInterval == 3 && record.CategorySeriesRange.TickInterval == 4 && record.CategorySeriesRange.Flags == 0x0005 && record.CategorySeriesRange.CrossesBetweenTickMarks && !record.CategorySeriesRange.CrossesAtMaximum && record.CategorySeriesRange.Reversed);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AxesUsed" && record.AxesUsedCount == 1);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Series" && record.SeriesCategoryDataType == 0x0003 && record.SeriesCategoryDataTypeName == "Text" && record.SeriesValueDataType == 0x0001 && record.SeriesValueDataTypeName == "Numeric" && record.SeriesCategoryCount == 4 && record.SeriesValueCount == 4 && record.SeriesBubbleSizeDataType == 0x0001 && record.SeriesBubbleSizeDataTypeName == "Numeric" && record.SeriesBubbleSizeCount == 0);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "SIIndex" && record.SeriesDataCacheIndex == 0x0002 && record.SeriesDataCacheIndexName == "CategoryLabelsOrHorizontalValues");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "DataFormat" && record.DataFormatPointIndex == 0xffff && record.DataFormatSeriesIndex == 2 && record.DataFormatOrder == 1 && record.DataFormatTarget == "Series");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Ifmt" && record.NumberFormatId == 14);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "LineFormat" && record.LineFormat != null && record.LineFormat.RgbHex == "#112233" && record.LineFormat.Style == 0x0001 && record.LineFormat.StyleName == "Dash" && record.LineFormat.Weight == 1 && record.LineFormat.WeightName == "Medium" && !record.LineFormat.Automatic && record.LineFormat.AxisVisible && !record.LineFormat.AutomaticColor && record.LineFormat.ColorIndex == 0x004d);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Frame" && record.Frame != null && record.Frame.FrameType == 0x0004 && record.Frame.FrameTypeName == "ShadowFrame" && record.Frame.Flags == 0x0003 && record.Frame.AutomaticSize && record.Frame.AutomaticPosition);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AreaFormat" && record.AreaFormat != null && record.AreaFormat.ForegroundRgbHex == "#AABBCC" && record.AreaFormat.BackgroundRgbHex == "#102030" && record.AreaFormat.Pattern == 0x0001 && record.AreaFormat.PatternName == "Solid" && record.AreaFormat.Automatic && record.AreaFormat.InvertNegative && record.AreaFormat.ForegroundColorIndex == 0x004e && record.AreaFormat.BackgroundColorIndex == 0x004d);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "MarkerFormat" && record.MarkerFormat != null && record.MarkerFormat.ForegroundRgbHex == "#DEADBE" && record.MarkerFormat.BackgroundRgbHex == "#445566" && record.MarkerFormat.MarkerType == 0x0008 && record.MarkerFormat.MarkerTypeName == "Circle" && record.MarkerFormat.Automatic && !record.MarkerFormat.InteriorHidden && record.MarkerFormat.BorderHidden && record.MarkerFormat.ForegroundColorIndex == 0x004e && record.MarkerFormat.BackgroundColorIndex == 0x004d && record.MarkerFormat.SizeTwips == 240);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "PieFormat" && record.PieFormat != null && record.PieFormat.ExplosionPercentage == 25);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AttachedLabel" && record.AttachedLabel != null && record.AttachedLabel.Flags == 0x0051 && record.AttachedLabel.ShowValue && !record.AttachedLabel.ShowPercent && !record.AttachedLabel.ShowLabelAndPercent && record.AttachedLabel.ShowLabel && !record.AttachedLabel.ShowBubbleSizes && record.AttachedLabel.ShowSeriesName && record.AttachedLabel.FlagNames.SequenceEqual(new[] { "ShowValue", "ShowLabel", "ShowSeriesName" }));
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "DefaultText" && record.DefaultTextId == 0x0002 && record.DefaultTextTargetName == "ChartUnscaledText");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Text" && record.Text != null && record.Text.HorizontalAlignmentName == "Center" && record.Text.VerticalAlignmentName == "Bottom" && record.Text.BackgroundModeName == "Transparent" && record.Text.RgbHex == "#224466" && record.Text.X == 120 && record.Text.Y == 240 && record.Text.Width == 800 && record.Text.Height == 160 && record.Text.Flags == 0x4095 && record.Text.FlagNames.SequenceEqual(new[] { "AutoColor", "ShowValue", "AutoText", "AutoMode", "ShowLabel" }) && record.Text.ColorIndex == 0x004d && record.Text.DataLabelPositionName == "Center" && record.Text.ReadingOrderName == "LeftToRight" && record.Text.Rotation == 45);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "FontX" && record.FontIndex == 3);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "ObjectLink" && record.ObjectLink != null && record.ObjectLink.LinkedObject == 0x0004 && record.ObjectLink.LinkedObjectName == "SeriesOrDataPoint" && record.ObjectLink.SeriesIndex == 2 && record.ObjectLink.DataPointIndex == 0xffff);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Legend" && record.Legend != null && record.Legend.X == 10 && record.Legend.Y == 20 && record.Legend.Width == 300 && record.Legend.Height == 400 && record.Legend.Spacing == 1 && record.Legend.Flags == 0x001d && record.Legend.AutoPosition && record.Legend.AutoPositionX && record.Legend.AutoPositionY && record.Legend.Vertical && !record.Legend.WasDataTable);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AxisLineFormat" && record.AxisLineFormat != null && record.AxisLineFormat.TargetId == 0x0000 && record.AxisLineFormat.TargetName == "AxisLine");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AxisLineFormat" && record.AxisLineFormat != null && record.AxisLineFormat.TargetId == 0x0001 && record.AxisLineFormat.TargetName == "MajorGridlines");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AxisLineFormat" && record.AxisLineFormat != null && record.AxisLineFormat.TargetId == 0x0002 && record.AxisLineFormat.TargetName == "MinorGridlines");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "AxcExt" && record.Kind == LegacyXlsChartRecordKind.Axis);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Dat" && record.DataTableOptions != null && record.DataTableOptions.Flags == 0x000d && record.DataTableOptions.HasHorizontalBorders && !record.DataTableOptions.HasVerticalBorders && record.DataTableOptions.HasOutlineBorder && record.DataTableOptions.ShowSeriesKeys);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Tick" && record.Tick != null && record.Tick.MajorTickLocationName == "Outside" && record.Tick.MinorTickLocationName == "Inside" && record.Tick.LabelLocationName == "NextToAxis" && record.Tick.BackgroundModeName == "Transparent" && record.Tick.RgbHex == "#998877" && record.Tick.Flags == 0x402d && record.Tick.RotationModeName == "RotatedClockwise" && record.Tick.AutoColor && !record.Tick.AutoBackground && record.Tick.AutoRotation && record.Tick.ReadingOrderName == "LeftToRight" && record.Tick.ColorIndex == 0x004d && record.Tick.Rotation == 30);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Pos" && record.Position != null && record.Position.TopLeftMode == 0x0005 && record.Position.TopLeftModeName == "MDCHART" && record.Position.BottomRightMode == 0x0001 && record.Position.BottomRightModeName == "MDABS" && record.Position.SemanticTypeName == "LegendManualSize" && record.Position.X1Y1MeaningName == "ChartAreaSprcOffset" && record.Position.X2Y2MeaningName == "PointSize" && record.Position.IgnoredCoordinateStateName == "None" && record.Position.HasKnownSemanticCombination && record.Position.X1 == 15 && record.Position.Y1 == 25 && record.Position.X2 == 300 && record.Position.Y2 == 120);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "BRAI" && record.DataSource != null && record.DataSource.SourceId == 0x01 && record.DataSource.SourceIdName == "ValuesOrHorizontalValues" && record.DataSource.ReferenceType == 0x02 && record.DataSource.ReferenceTypeName == "WorksheetRange" && record.DataSource.Flags == 0x0001 && record.DataSource.UsesCustomNumberFormat && record.DataSource.NumberFormatId == 14 && record.DataSource.FormulaByteCount == 9 && record.DataSource.FormulaBytesAvailable == 9 && record.DataSource.FormulaByteCountFitsPayload && record.DataSource.FormulaTextProjected && record.DataSource.FormulaText == "$B$1:$B$4");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "PlotGrowth" && record.PlotGrowth != null && record.PlotGrowth.HorizontalIntegral == 1 && record.PlotGrowth.HorizontalFractional == 0x4000 && record.PlotGrowth.HorizontalGrowthPoints == 1.25 && record.PlotGrowth.VerticalIntegral == 2 && record.PlotGrowth.VerticalFractional == 0x8000 && record.PlotGrowth.VerticalGrowthPoints == 2.5);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "GelFrame" && record.Kind == LegacyXlsChartRecordKind.Formatting);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "BopPopCustom" && record.Kind == LegacyXlsChartRecordKind.ChartType && record.ChartTypeName == "CustomBarOfPieOrPieOfPie");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Fbi2" && record.Kind == LegacyXlsChartRecordKind.Text);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Chart3d" && record.Kind == LegacyXlsChartRecordKind.ChartType && record.ChartTypeName == "ThreeDimensional");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Chart3DBarShape" && record.Kind == LegacyXlsChartRecordKind.ChartType && record.ChartTypeName == "ThreeDimensionalBarShape" && record.ThreeDimensionalBarShapeOptions != null && record.ThreeDimensionalBarShapeOptions.Riser == 0x01 && record.ThreeDimensionalBarShapeOptions.RiserName == "Ellipse" && record.ThreeDimensionalBarShapeOptions.Taper == 0x02 && record.ThreeDimensionalBarShapeOptions.TaperName == "ProjectedPoint");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Scatter" && record.Kind == LegacyXlsChartRecordKind.ChartType && record.ChartTypeName == "Scatter" && record.ScatterOptions != null && record.ScatterOptions.BubbleSizeRatio == 150 && record.ScatterOptions.BubbleSizeRepresentation == 0x0002 && record.ScatterOptions.BubbleSizeRepresentationName == "Width" && record.ScatterOptions.HasKnownBubbleSizeRepresentation && record.ScatterOptions.HasValidBubbleSizeRatio && record.ScatterOptions.Flags == 0x0007 && record.ScatterOptions.IsBubbleChart && record.ScatterOptions.ShowNegativeBubbles && record.ScatterOptions.HasShadow);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "End" && record.ContainerDepthBefore == 1 && record.ContainerDepthAfter == 0 && record.ContainerTransition == "End");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.ChartTypeName == "Scatter");
            Assert.Contains(workbook.Diagnostics, d => d.SheetName == "ChartOnly" && d.DetailCode == "Chart:Chart");
            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Records By Rectangle", markdown);
            Assert.Contains("Unsupported Chart Sheet States", markdown);
            Assert.Contains("Chart Records By Name And Payload Length", markdown);
            Assert.Contains("Chart Workbook States", markdown);
            Assert.Contains("Chart Records By Container Depth Before", markdown);
            Assert.Contains("Chart Records By Container Depth After", markdown);
            Assert.Contains("Chart Records By Container Transition", markdown);
            Assert.Contains("Chart Records By Name And Container Depth", markdown);
            Assert.Contains("Chart Records By Name And Container Transition", markdown);
            Assert.Contains("Chart Records By Axis Type", markdown);
            Assert.Contains("Chart Records By Axes Used Count", markdown);
            Assert.Contains("Chart CatSerRange Intervals", markdown);
            Assert.Contains("Chart CatSerRange States", markdown);
            Assert.Contains("Chart AxisLineFormat Targets", markdown);
            Assert.Contains("Chart Series Category Data Types", markdown);
            Assert.Contains("Chart Series Value Data Types", markdown);
            Assert.Contains("Chart Series Bubble Size Data Types", markdown);
            Assert.Contains("Chart Series Value Counts", markdown);
            Assert.Contains("Chart Series Data Cache Indexes", markdown);
            Assert.Contains("Chart Series Data Cache Types", markdown);
            Assert.Contains("Chart DataSource Ids", markdown);
            Assert.Contains("Chart DataSource Reference Types", markdown);
            Assert.Contains("Chart DataSource Number Format Ids", markdown);
            Assert.Contains("Chart DataSource Formula Byte Counts", markdown);
            Assert.Contains("Chart DataSource Formula Projection States", markdown);
            Assert.Contains("Chart DataSource States", markdown);
            Assert.Contains("Chart DataFormat Targets", markdown);
            Assert.Contains("Chart DataFormat Series Indexes", markdown);
            Assert.Contains("Chart Number Format Ids", markdown);
            Assert.Contains("Chart Font Indexes", markdown);
            Assert.Contains("Chart DataTable Options", markdown);
            Assert.Contains("Chart 3D Bar Shape Risers", markdown);
            Assert.Contains("Chart 3D Bar Shape Tapers", markdown);
            Assert.Contains("Chart 3D Bar Shape States", markdown);
            Assert.Contains("Chart Scatter Bubble Size Ratios", markdown);
            Assert.Contains("Chart Scatter Bubble Size Representations", markdown);
            Assert.Contains("Chart Scatter Bubble Size Ratio States", markdown);
            Assert.Contains("Chart Scatter States", markdown);
            Assert.Contains("Chart LineFormat Styles", markdown);
            Assert.Contains("Chart LineFormat Weights", markdown);
            Assert.Contains("Chart AreaFormat Patterns", markdown);
            Assert.Contains("Chart MarkerFormat Types", markdown);
            Assert.Contains("Chart MarkerFormat Sizes", markdown);
            Assert.Contains("Chart PieFormat Explosions", markdown);
            Assert.Contains("Chart AttachedLabel Flags", markdown);
            Assert.Contains("Chart AttachedLabel States", markdown);
            Assert.Contains("Chart DefaultText Targets", markdown);
            Assert.Contains("Chart Text Horizontal Alignments", markdown);
            Assert.Contains("Chart ObjectLink Targets", markdown);
            Assert.Contains("Chart Tick Label Locations", markdown);
            Assert.Contains("Chart ValueRange Scales", markdown);
            Assert.Contains("Chart ValueRange States", markdown);
            Assert.Contains("Chart Position Mode Pairs", markdown);
            Assert.Contains("Chart Position Rectangles", markdown);
            Assert.Contains("Chart Position Semantic Types", markdown);
            Assert.Contains("Chart Position Coordinate Meanings", markdown);
            Assert.Contains("Chart Position Ignored Coordinate States", markdown);
            Assert.Contains("Chart Position Known Semantic States", markdown);
            Assert.Contains("Chart Frame Types", markdown);
            Assert.Contains("Chart Frame Auto States", markdown);
            Assert.Contains("Chart PlotGrowth Factors", markdown);
            Assert.Contains("Unsupported Chart Sheet Chart Record Counts", markdown);
            Assert.Contains("Unsupported Chart Sheet Chart Record Kinds", markdown);
            Assert.Contains("Unsupported Chart Sheet Chart Types", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartDataSourceFormulaProjectionFailures() {
            byte[] payload = {
                0x01,
                0x02,
                0x01, 0x00,
                0x0e, 0x00,
                0x01, 0x00,
                0x01
            };
            var chartRecord = new BiffRecord(0x1051, offset: 123, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "ChartDiag", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartDataSource? dataSource = record.DataSource;
            Assert.NotNull(dataSource);
            Assert.False(dataSource.FormulaTextProjected);
            Assert.True(dataSource.HasFormulaProjectionFailure);
            Assert.Equal("FormulaToken0x01", dataSource.FormulaProjectionFailureCode);
            Assert.Equal((byte)0x01, dataSource.FormulaProjectionFailureToken);
            Assert.Equal("PtgExp", dataSource.FormulaProjectionFailureTokenName);
            Assert.Equal(0, dataSource.FormulaProjectionFailureTokenOffset);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartDataSourceFormulaProjectionStates["FormulaTextUnsupported"]);
            Assert.Equal(1, report.ChartDataSourceFormulaProjectionFailures["FormulaToken0x01"]);
            Assert.Equal(1, report.ChartDataSourceFormulaProjectionFailuresByToken["Token:0x01"]);
            Assert.Equal(1, report.ChartDataSourceFormulaProjectionFailuresByTokenName["PtgExp"]);
            Assert.Equal(1, report.ChartDataSourceFormulaProjectionFailuresByOffset["Offset:0"]);
            Assert.Equal(1, report.ChartDataSourceStates["Source:ValuesOrHorizontalValues;Reference:WorksheetRange;CustomNumberFormat:True;FormulaBytes:1;FormulaComplete:True;FormulaTextProjected:False"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart DataSource Formula Projection Failures", markdown);
            Assert.Contains("FormulaToken0x01", markdown);
            Assert.Contains("PtgExp", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartFontBasisMetadata() {
            byte[] payload = {
                0xc0, 0x12,
                0x80, 0x0c,
                0xdc, 0x00,
                0x01, 0x00,
                0x07, 0x00
            };
            var chartRecord = new BiffRecord(0x1060, offset: 456, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "FontChart", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartFontBasisOptions? fontBasis = record.FontBasisOptions;
            Assert.NotNull(fontBasis);
            Assert.Equal(4800, fontBasis!.WidthTwipsBasis);
            Assert.Equal(3200, fontBasis.HeightTwipsBasis);
            Assert.Equal(220, fontBasis.FontHeightTwips);
            Assert.Equal(0x0001, fontBasis.ScaleBasis);
            Assert.Equal("PlotArea", fontBasis.ScaleBasisName);
            Assert.True(fontBasis.HasKnownScaleBasis);
            Assert.Equal(7, fontBasis.FontIndex);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartFontBasisScaleBasis["PlotArea"]);
            Assert.Equal(1, report.ChartFontBasisFontIndexes["FontIndex:7"]);
            Assert.Equal(1, report.ChartFontBasisStates["Basis:4800x3200;HeightTwips:220;Scale:PlotArea;FontIndex:7"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart FontBasis Scale Basis", markdown);
            Assert.Contains("Basis:4800x3200;HeightTwips:220;Scale:PlotArea;FontIndex:7", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartGroupAndPivotViewReferences() {
            var chartRecords = new List<LegacyXlsChartRecord>();
            byte[] chartFormatPayload = {
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x01, 0x00,
                0x02, 0x00
            };
            byte[] seriesChartGroupPayload = {
                0x03, 0x00
            };
            byte[] pivotViewReferencePayload = {
                0x01, 0x00,
                0x03, 0x00,
                0x02, 0x00,
                0x04, 0x00
            };

            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x1014, offset: 100, chartFormatPayload), "ChartGroup", chartRecords));
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x1044, offset: 120, seriesChartGroupPayload), "ChartGroup", chartRecords));
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x1046, offset: 140, pivotViewReferencePayload), "ChartGroup", chartRecords));

            LegacyXlsChartRecord chartGroupRecord = chartRecords[0];
            LegacyXlsChartRecord seriesLinkRecord = chartRecords[1];
            LegacyXlsChartRecord pivotViewRecord = chartRecords[2];

            Assert.NotNull(chartGroupRecord.ChartGroupOptions);
            Assert.True(chartGroupRecord.ChartGroupOptions!.VariedDataPointColors);
            Assert.Equal(2, chartGroupRecord.ChartGroupOptions.DrawingOrder);
            Assert.NotNull(seriesLinkRecord.SeriesChartGroupReference);
            Assert.Equal(3, seriesLinkRecord.SeriesChartGroupReference!.ChartGroupIndex);
            Assert.NotNull(pivotViewRecord.PivotViewReference);
            Assert.Equal("C2:E4", pivotViewRecord.PivotViewReference!.Reference);

            var workbook = new LegacyXlsWorkbook();
            foreach (LegacyXlsChartRecord record in chartRecords) {
                workbook.MutableChartRecords.Add(record);
            }

            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartGroupVariedColorStates["VariedDataPointColors"]);
            Assert.Equal(1, report.ChartGroupDrawingOrders["DrawingOrder:2"]);
            Assert.Equal(1, report.ChartSeriesChartGroupIndexes["ChartGroupIndex:3"]);
            Assert.Equal(1, report.ChartPivotViewReferences["C2:E4"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Group Varied Color States", markdown);
            Assert.Contains("Chart Series Chart Group Indexes", markdown);
            Assert.Contains("Chart Pivot View References", markdown);
            Assert.Contains("C2:E4", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartCategoryLabelOptions() {
            byte[] payload = {
                0x56, 0x08,
                0x00, 0x00,
                0x96, 0x00,
                0x02, 0x00,
                0x01, 0x00,
                0x00, 0x00
            };
            var chartRecord = new BiffRecord(0x0856, offset: 160, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "CategoryLabels", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartCategoryLabelOptions? options = record.CategoryLabelOptions;
            Assert.NotNull(options);
            Assert.Equal(150, options!.OffsetPercentage);
            Assert.Equal(0x0002, options.Alignment);
            Assert.Equal("Center", options.AlignmentName);
            Assert.True(options.UseAutomaticLabelCount);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartCategoryLabelAlignments["Center"]);
            Assert.Equal(1, report.ChartCategoryLabelOffsets["Offset:150%"]);
            Assert.Equal(1, report.ChartCategoryLabelCountStates["AutomaticLabelCount"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart CatLab Alignments", markdown);
            Assert.Contains("Offset:150%", markdown);
            Assert.Contains("AutomaticLabelCount", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartLayout12Metadata() {
            byte[] payload = new byte[60];
            WriteUInt16(payload, 0, 0x089D);
            WriteUInt32(payload, 12, 0x12345678);
            WriteUInt16(payload, 16, 0x0006);
            WriteUInt16(payload, 18, 0x0001);
            WriteUInt16(payload, 20, 0x0002);
            WriteUInt16(payload, 22, 0x0000);
            WriteUInt16(payload, 24, 0x0001);
            WriteDouble(payload, 26, 0.125);
            WriteDouble(payload, 34, 0.25);
            WriteDouble(payload, 42, 0.5);
            WriteDouble(payload, 50, 0.75);
            var chartRecord = new BiffRecord(0x089D, offset: 176, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "Layout", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartLayout12? layout = record.Layout12;
            Assert.NotNull(layout);
            Assert.Equal(0x12345678u, layout!.Checksum);
            Assert.Equal(0x03, layout.AutomaticLayoutType);
            Assert.Equal("Right", layout.AutomaticLayoutTypeName);
            Assert.Equal("Factor", layout.XModeName);
            Assert.Equal("Edge", layout.YModeName);
            Assert.Equal("Automatic", layout.WidthModeName);
            Assert.Equal("Factor", layout.HeightModeName);
            Assert.Equal(0.125, layout.X);
            Assert.Equal(0.25, layout.Y);
            Assert.Equal(0.5, layout.Width);
            Assert.Equal(0.75, layout.Height);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartLayout12ModePairs["X:Factor;Y:Edge;Width:Automatic;Height:Factor"]);
            Assert.Equal(1, report.ChartLayout12AutoLayoutTypes["Right"]);
            Assert.Equal(1, report.ChartLayout12Checksums["Checksum:0x12345678"]);
            Assert.Equal(1, report.ChartLayout12Rectangles["X:0.125;Y:0.25;Width:0.5;Height:0.75"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart CrtLayout12 Mode Pairs", markdown);
            Assert.Contains("Checksum:0x12345678", markdown);
            Assert.Contains("X:0.125;Y:0.25;Width:0.5;Height:0.75", markdown);

            static void WriteUInt16(byte[] buffer, int offset, ushort value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)(value >> 8);
            }

            static void WriteUInt32(byte[] buffer, int offset, uint value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)((value >> 8) & 0xff);
                buffer[offset + 2] = (byte)((value >> 16) & 0xff);
                buffer[offset + 3] = (byte)(value >> 24);
            }

            static void WriteDouble(byte[] buffer, int offset, double value) {
                byte[] bytes = BitConverter.GetBytes(value);
                Array.Copy(bytes, 0, buffer, offset, bytes.Length);
            }
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartPlotAreaLayout12Metadata() {
            byte[] payload = new byte[68];
            WriteUInt16(payload, 0, 0x08A7);
            WriteUInt32(payload, 12, 0x00000001);
            WriteUInt16(payload, 16, 0x0001);
            WriteInt16(payload, 18, 10);
            WriteInt16(payload, 20, 20);
            WriteInt16(payload, 22, 300);
            WriteInt16(payload, 24, 400);
            WriteUInt16(payload, 26, 0x0001);
            WriteUInt16(payload, 28, 0x0002);
            WriteUInt16(payload, 30, 0x0000);
            WriteUInt16(payload, 32, 0x0001);
            WriteDouble(payload, 34, 0.1);
            WriteDouble(payload, 42, 0.2);
            WriteDouble(payload, 50, 0.3);
            WriteDouble(payload, 58, 0.4);
            var chartRecord = new BiffRecord(0x08A7, offset: 224, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "CrtLayout12A", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            Assert.Equal("CrtLayout12A", record.RecordName);
            Assert.Equal(LegacyXlsChartRecordKind.Layout, record.Kind);
            LegacyXlsChartPlotAreaLayout12? layout = record.PlotAreaLayout12;
            Assert.NotNull(layout);
            Assert.Equal(0x00000001u, layout!.Checksum);
            Assert.True(layout.TargetsInnerPlotArea);
            Assert.Equal("InnerPlotArea", layout.TargetName);
            Assert.Equal(10, layout.UpperLeftX);
            Assert.Equal(20, layout.UpperLeftY);
            Assert.Equal(300, layout.WidthSprc);
            Assert.Equal(400, layout.HeightSprc);
            Assert.Equal("Factor", layout.XModeName);
            Assert.Equal("Edge", layout.YModeName);
            Assert.Equal("Automatic", layout.WidthModeName);
            Assert.Equal("Factor", layout.HeightModeName);
            Assert.Equal(0.1, layout.X);
            Assert.Equal(0.2, layout.Y);
            Assert.Equal(0.3, layout.Width);
            Assert.Equal(0.4, layout.Height);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartPlotAreaLayout12Targets["InnerPlotArea"]);
            Assert.Equal(1, report.ChartPlotAreaLayout12ModePairs["X:Factor;Y:Edge;Width:Automatic;Height:Factor"]);
            Assert.Equal(1, report.ChartPlotAreaLayout12Checksums["Checksum:0x00000001"]);
            Assert.Equal(1, report.ChartPlotAreaLayout12Bounds["X:10;Y:20;Width:300;Height:400"]);
            Assert.Equal(1, report.ChartPlotAreaLayout12Rectangles["X:0.1;Y:0.2;Width:0.3;Height:0.4"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart CrtLayout12A Targets", markdown);
            Assert.Contains("InnerPlotArea", markdown);
            Assert.Contains("X:10;Y:20;Width:300;Height:400", markdown);

            static void WriteUInt16(byte[] buffer, int offset, ushort value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)(value >> 8);
            }

            static void WriteInt16(byte[] buffer, int offset, short value) {
                WriteUInt16(buffer, offset, unchecked((ushort)value));
            }

            static void WriteUInt32(byte[] buffer, int offset, uint value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)((value >> 8) & 0xff);
                buffer[offset + 2] = (byte)((value >> 16) & 0xff);
                buffer[offset + 3] = (byte)(value >> 24);
            }

            static void WriteDouble(byte[] buffer, int offset, double value) {
                byte[] bytes = BitConverter.GetBytes(value);
                Array.Copy(bytes, 0, buffer, offset, bytes.Length);
            }
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartFutureRecordInfo() {
            byte[] payload = {
                0x50, 0x08,
                0x00, 0x00,
                0x0E,
                0x0E,
                0x04, 0x00,
                0x50, 0x08, 0x5A, 0x08,
                0x61, 0x08, 0x61, 0x08,
                0x6A, 0x08, 0x6B, 0x08,
                0x9D, 0x08, 0xA6, 0x08
            };
            var chartRecord = new BiffRecord(0x0850, offset: 208, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "FutureInfo", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartFutureRecordInfo? info = record.FutureRecordInfo;
            Assert.NotNull(info);
            Assert.Equal(0x0E, info!.OriginatorVersion);
            Assert.Equal(0x0E, info.WriterVersion);
            Assert.Equal("Version:0x0E", info.OriginatorVersionName);
            Assert.Equal("Version:0x0E", info.WriterVersionName);
            Assert.Collection(info.Ranges,
                range => Assert.Equal("0x0850-0x085A", range.RangeKey),
                range => Assert.Equal("0x0861-0x0861", range.RangeKey),
                range => Assert.Equal("0x086A-0x086B", range.RangeKey),
                range => Assert.Equal("0x089D-0x08A6", range.RangeKey));

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartFutureRecordInfoVersions["Originator:Version:0x0E;Writer:Version:0x0E"]);
            Assert.Equal(1, report.ChartFutureRecordInfoRangeCounts["Ranges:4"]);
            Assert.Equal(1, report.ChartFutureRecordInfoRanges["0x089D-0x08A6"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Future Record Info Versions", markdown);
            Assert.Contains("Ranges:4", markdown);
            Assert.Contains("0x089D-0x08A6", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartUnitsMetadata() {
            byte[] payload = {
                0x00, 0x00
            };
            var chartRecord = new BiffRecord(0x1001, offset: 284, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "Units", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            Assert.Equal("Units", record.RecordName);
            Assert.Equal(LegacyXlsChartRecordKind.Container, record.Kind);
            LegacyXlsChartUnits? units = record.Units;
            Assert.NotNull(units);
            Assert.Equal(0, units!.Reserved);
            Assert.True(units.HasZeroReservedValue);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartUnitsReservedValues["Reserved:0x0000"]);
            Assert.Equal(1, report.ChartUnitsReservedStates["ReservedZero"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Units Reserved Values", markdown);
            Assert.Contains("Reserved:0x0000", markdown);
            Assert.Contains("ReservedZero", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartFutureBlockMetadata() {
            byte[] startPayload = {
                0x52, 0x08,
                0x00, 0x00,
                0x0C, 0x00,
                0x00, 0x00,
                0x02, 0x00,
                0x00, 0x00
            };
            byte[] endPayload = {
                0x53, 0x08,
                0x00, 0x00,
                0x0C, 0x00,
                0x00, 0x00,
                0x00, 0x00,
                0x00, 0x00
            };
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x0852, offset: 256, startPayload), "StartBlock", chartRecords));
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x0853, offset: 268, endPayload), "EndBlock", chartRecords));

            Assert.Equal(2, chartRecords.Count);
            LegacyXlsChartFutureBlock? startBlock = chartRecords[0].FutureBlock;
            Assert.NotNull(startBlock);
            Assert.True(startBlock!.IsStart);
            Assert.Equal("StartBlock", startBlock.DirectionName);
            Assert.Equal(0x000C, startBlock.ObjectKind);
            Assert.Equal("Series", startBlock.ObjectKindName);
            Assert.Equal((ushort?)0x0000, startBlock.ObjectContext);
            Assert.Equal((ushort?)0x0002, startBlock.ObjectInstance1);
            Assert.Equal((ushort?)0x0000, startBlock.ObjectInstance2);

            LegacyXlsChartFutureBlock? endBlock = chartRecords[1].FutureBlock;
            Assert.NotNull(endBlock);
            Assert.True(endBlock!.IsEnd);
            Assert.Equal("EndBlock", endBlock.DirectionName);
            Assert.Equal("Series", endBlock.ObjectKindName);
            Assert.Null(endBlock.ObjectContext);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.AddRange(chartRecords);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartFutureBlockDirections["StartBlock"]);
            Assert.Equal(1, report.ChartFutureBlockDirections["EndBlock"]);
            Assert.Equal(2, report.ChartFutureBlockObjectKinds["Series"]);
            Assert.Equal(1, report.ChartFutureBlockScopes["Kind:Series;Context:0x0000;Instance1:0x0002;Instance2:0x0000"]);
            Assert.Equal(1, report.ChartFutureBlockScopes["Kind:Series"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Future Block Directions", markdown);
            Assert.Contains("Chart Future Block Object Kinds", markdown);
            Assert.Contains("Kind:Series;Context:0x0000;Instance1:0x0002;Instance2:0x0000", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartXmlTokenChainMetadata() {
            byte[] payload = {
                0x9E, 0x08,
                0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x08, 0x00, 0x00, 0x00,
                0x01, 0x00, 0x00, 0x00,
                0x02, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00
            };
            var chartRecord = new BiffRecord(0x089E, offset: 232, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "XmlTokens", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartXmlTokenChain? chain = record.XmlTokenChain;
            Assert.NotNull(chain);
            Assert.Equal(8u, chain!.DeclaredByteCount);
            Assert.Equal(8, chain.FirstSegmentByteCount);
            Assert.Equal(0u, chain.TrailingUnusedValue);
            Assert.True(chain.IsCompleteInRecord);
            Assert.True(chain.HasZeroTrailingUnusedValue);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartXmlTokenChainDeclaredByteCounts["DeclaredBytes:8"]);
            Assert.Equal(1, report.ChartXmlTokenChainFirstSegmentByteCounts["FirstSegmentBytes:8"]);
            Assert.Equal(1, report.ChartXmlTokenChainCompletionStates["CompleteInRecord"]);
            Assert.Equal(1, report.ChartXmlTokenChainTrailingStates["TrailingUnusedZero"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart XmlTkChain Declared Byte Counts", markdown);
            Assert.Contains("FirstSegmentBytes:8", markdown);
            Assert.Contains("CompleteInRecord", markdown);
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
            Assert.Equal(3, report.DataValidationsByErrorStyle["Stop"]);
            Assert.Equal(2, report.DataValidationsByAllowBlankState["AllowBlank"]);
            Assert.Equal(1, report.DataValidationsByAllowBlankState["RejectBlank"]);
            Assert.Equal(1, report.DataValidationsByInputMessageState["ShowInputMessage"]);
            Assert.Equal(2, report.DataValidationsByInputMessageState["HideInputMessage"]);
            Assert.Equal(3, report.DataValidationsByErrorMessageState["ShowErrorMessage"]);
            Assert.Equal(1, report.DataValidationsByPromptTextState["Present"]);
            Assert.Equal(2, report.DataValidationsByPromptTextState["Missing"]);
            Assert.Equal(3, report.DataValidationsByErrorTextState["Present"]);
            Assert.Equal(3, report.DataValidationsByDropDownState["NotList"]);
            Assert.Equal(3, report.DataValidationsBySheet["TypedValidation"]);
            Assert.Equal(3, report.DataValidationsByRangeCount["Ranges:1"]);
            Assert.Equal(1, report.DataValidationsByRange["E2:E5"]);
            Assert.Equal(1, report.DataValidationsByRange["F2:F5"]);
            Assert.Equal(1, report.DataValidationsByRange["G2:G5"]);
            Assert.Equal(1, report.DataValidationsBySheetAndRange["TypedValidation!E2:E5"]);
            Assert.Equal(1, report.DataValidationsBySheetAndRange["TypedValidation!F2:F5"]);
            Assert.Equal(1, report.DataValidationsBySheetAndRange["TypedValidation!G2:G5"]);
            Assert.Equal(3, report.DataValidationsByFormula1State["Present"]);
            Assert.Equal(1, report.DataValidationsByFormula2State["Present"]);
            Assert.Equal(2, report.DataValidationsByFormula2State["Missing"]);
            Assert.Equal(1, report.DataValidationsByFormulaPairState["Formula1:Present|Formula2:Present"]);
            Assert.Equal(2, report.DataValidationsByFormulaPairState["Formula1:Present|Formula2:Missing"]);
            Assert.Empty(report.DataValidationListSourcesByKind);
            string markdown = report.ToMarkdown();
            Assert.Contains("Data validations: 3", markdown);
            Assert.Contains("Data Validations By Type", markdown);
            Assert.Contains("Data Validations By Operator", markdown);
            Assert.Contains("Data Validations By Error Style", markdown);
            Assert.Contains("Data Validations By Allow Blank State", markdown);
            Assert.Contains("Data Validations By Input Message State", markdown);
            Assert.Contains("Data Validations By Error Message State", markdown);
            Assert.Contains("Data Validations By Prompt Text State", markdown);
            Assert.Contains("Data Validations By Error Text State", markdown);
            Assert.Contains("Data Validations By Drop Down State", markdown);
            Assert.Contains("Data Validations By Sheet", markdown);
            Assert.Contains("Data Validations By Range Count", markdown);
            Assert.Contains("Data Validations By Range", markdown);
            Assert.Contains("Data Validations By Sheet And Range", markdown);
            Assert.Contains("Data Validations By Formula1 State", markdown);
            Assert.Contains("Data Validations By Formula2 State", markdown);
            Assert.Contains("Data Validations By Formula Pair State", markdown);
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
            Assert.Equal(1, report.ConditionalFormattingsByPriorityState["Missing"]);
            Assert.Empty(report.ConditionalFormattingsByPriority);
            Assert.Equal(1, report.ConditionalFormattingsBySheet["ConditionalRule"]);
            Assert.Equal(1, report.ConditionalFormattingsByRangeCount["Ranges:1"]);
            Assert.Equal(1, report.ConditionalFormattingsByRange["A1:A3"]);
            Assert.Equal(1, report.ConditionalFormattingsBySheetAndRange["ConditionalRule!A1:A3"]);
            Assert.Equal(1, report.ConditionalFormattingsByFormula1State["Present"]);
            Assert.Equal(1, report.ConditionalFormattingsByFormula2State["Missing"]);
            Assert.Equal(1, report.ConditionalFormattingsByFormulaPairState["Formula1:Present|Formula2:Missing"]);
            Assert.Equal(1, report.ConditionalFormattingsByStopIfTrueState["Continue"]);
            Assert.Equal(1, report.ConditionalFormattingsByDifferentialFormatState["Missing"]);
            Assert.Empty(report.ConditionalFormattingsByDifferentialFill);
            string markdown = report.ToMarkdown();
            Assert.Contains("Conditional formatting rules: 1", markdown);
            Assert.Contains("Conditional Formatting By Type", markdown);
            Assert.Contains("Conditional Formatting By Operator", markdown);
            Assert.Contains("Conditional Formatting By Sheet", markdown);
            Assert.Contains("Conditional Formatting By Range Count", markdown);
            Assert.Contains("Conditional Formatting By Range", markdown);
            Assert.Contains("Conditional Formatting By Sheet And Range", markdown);
            Assert.Contains("Conditional Formatting By Formula1 State", markdown);
            Assert.Contains("Conditional Formatting By Formula2 State", markdown);
            Assert.Contains("Conditional Formatting By Formula Pair State", markdown);
            Assert.Contains("Conditional Formatting By Priority State", markdown);
            Assert.Contains("Conditional Formatting By Stop If True State", markdown);
            Assert.Contains("Conditional Formatting By Differential Format State", markdown);
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
            Assert.Equal(1, report.AutoFilterCriteriaByValueKind["Text"]);
            Assert.Equal(1, report.AutoFilterCriteriaByValueKind["Number"]);
            Assert.Equal(1, report.AutoFilterCriteriaByTextPattern["ExactText"]);
            Assert.Equal(2, report.AutoFilterCriteriaByJoinOperator["Single"]);
            Assert.Equal(2, report.AutoFilterCriteriaBySheet["Filtered"]);
            Assert.Equal(1, report.AutoFilterCriteriaByColumn["Column:0"]);
            Assert.Equal(1, report.AutoFilterCriteriaByColumn["Column:1"]);
            Assert.Equal(1, report.AutoFilterCriteriaBySheetAndColumn["Filtered!Column:0"]);
            Assert.Equal(1, report.AutoFilterCriteriaBySheetAndColumn["Filtered!Column:1"]);
            Assert.Equal(2, report.AutoFilterCriteriaByConditionCount["Conditions:1"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("AutoFilter criteria columns: 2", markdown);
            Assert.Contains("AutoFilter Criteria By Sheet", markdown);
            Assert.Contains("AutoFilter Criteria By Operator", markdown);
            Assert.Contains("AutoFilter Criteria By Value Kind", markdown);
            Assert.Contains("AutoFilter Criteria By Text Pattern", markdown);
            Assert.Contains("AutoFilter Criteria By Join Operator", markdown);
            Assert.Contains("AutoFilter Criteria By Column", markdown);
            Assert.Contains("AutoFilter Criteria By Sheet And Column", markdown);
            Assert.Contains("AutoFilter Criteria By Condition Count", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_CountsAutoFilterWildcardTextPatterns() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AutoFilterWildcardTextWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(1, report.AutoFilterCriteriaCount);
            Assert.Equal(1, report.AutoFilterCriteriaByValueKind["Text"]);
            Assert.Equal(1, report.AutoFilterCriteriaByTextPattern["Contains"]);
            Assert.Equal(1, report.AutoFilterCriteriaByJoinOperator["Single"]);
            Assert.Equal(1, report.AutoFilterCriteriaByConditionCount["Conditions:1"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("AutoFilter Criteria By Text Pattern", markdown);
            Assert.Contains("Contains", markdown);
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
            Assert.Equal(0, compoundRecord.EntrySizes["_VBA_PROJECT_CUR"]);
            Assert.Equal(LegacyXlsCompoundFeatureEntryObjectType.Storage, compoundRecord.EntryObjectTypes["_VBA_PROJECT_CUR"]);
            LegacyXlsCompoundFeatureEntryInfo entry = Assert.Single(compoundRecord.EntryDetails);
            Assert.Equal("_VBA_PROJECT_CUR", entry.Path);
            Assert.Equal(LegacyXlsCompoundFeatureEntryRole.VbaProjectStorage, entry.Role);
            Assert.Equal(LegacyXlsCompoundFeatureEntryObjectType.Storage, entry.ObjectType);
            Assert.True(entry.IsStorage);
            Assert.False(entry.IsStream);
            Assert.Equal(0, entry.SizeBytes);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntryCount);
            Assert.Equal(0, result.ImportReport.CompoundFeatureEntryByteCount);
            Assert.Equal(0, result.ImportReport.CompoundVbaModuleByteCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordsByKind[LegacyXlsCompoundFeatureRecordKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKind[LegacyXlsCompoundFeatureRecordKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByName["_VBA_PROJECT_CUR"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRole["VbaProjectStorage"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKindAndRole["VbaProject|VbaProjectStorage"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByObjectType["Storage"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRoleAndObjectType["VbaProjectStorage|Storage"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesBySize["Bytes:0"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRoleAndSize["VbaProjectStorage|Bytes:0"]);
            Assert.Empty(compoundRecord.VbaModuleNames);
            Assert.Equal(0, result.ImportReport.CompoundVbaModuleCount);
            Assert.Equal(1, result.ImportReport.CompoundVbaProjectsByModuleCount["Modules:0"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaProjectsByModuleByteCount["Bytes:0"]);
            Assert.Equal(1, result.ImportReport.VbaProjectWorkbookStates["BiffMarker:Missing|NoMacrosMarker:Missing|CompoundProject:Present|Modules:Missing"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["VbaProject|XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED|Compound:VbaProjectStorage"]);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("VbaProject", markdown);
            Assert.Contains("Compound Feature Entries By Name", markdown);
            Assert.Contains("Compound Feature Entries By Role", markdown);
            Assert.Contains("Compound Feature Entries By Object Type", markdown);
            Assert.Contains("Compound Feature Entries By Size", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_SummarizesCompoundVbaModuleNames() {
            var workbook = new LegacyXlsWorkbook();
            workbook.SetHasVbaProjectMarker();
            workbook.SetCodeName("ThisWorkbook");
            var sheet = new LegacyXlsWorksheet("Data", 0, 0, 0);
            sheet.SetCodeName("Sheet1");
            workbook.MutableWorksheets.Add(sheet);
            var entryRoles = new Dictionary<string, LegacyXlsCompoundFeatureEntryRole>(StringComparer.OrdinalIgnoreCase) {
                ["_VBA_PROJECT_CUR"] = LegacyXlsCompoundFeatureEntryRole.VbaProjectStorage,
                ["_VBA_PROJECT_CUR/VBA"] = LegacyXlsCompoundFeatureEntryRole.VbaStorage,
                ["_VBA_PROJECT_CUR/VBA/dir"] = LegacyXlsCompoundFeatureEntryRole.VbaDirStream,
                ["_VBA_PROJECT_CUR/VBA/Sheet1"] = LegacyXlsCompoundFeatureEntryRole.VbaModuleStream,
                ["_VBA_PROJECT_CUR/VBA/ThisWorkbook"] = LegacyXlsCompoundFeatureEntryRole.VbaModuleStream,
                ["_VBA_PROJECT_CUR/VBA/LooseModule"] = LegacyXlsCompoundFeatureEntryRole.VbaModuleStream,
                ["_VBA_PROJECT_CUR/VBA/_VBA_PROJECT"] = LegacyXlsCompoundFeatureEntryRole.VbaProjectStream
            };
            var entrySizes = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase) {
                ["_VBA_PROJECT_CUR"] = 0,
                ["_VBA_PROJECT_CUR/VBA"] = 0,
                ["_VBA_PROJECT_CUR/VBA/dir"] = 50,
                ["_VBA_PROJECT_CUR/VBA/Sheet1"] = 100,
                ["_VBA_PROJECT_CUR/VBA/ThisWorkbook"] = 200,
                ["_VBA_PROJECT_CUR/VBA/LooseModule"] = 300,
                ["_VBA_PROJECT_CUR/VBA/_VBA_PROJECT"] = 20
            };
            var entryObjectTypes = new Dictionary<string, LegacyXlsCompoundFeatureEntryObjectType>(StringComparer.OrdinalIgnoreCase) {
                ["_VBA_PROJECT_CUR"] = LegacyXlsCompoundFeatureEntryObjectType.Storage,
                ["_VBA_PROJECT_CUR/VBA"] = LegacyXlsCompoundFeatureEntryObjectType.Storage,
                ["_VBA_PROJECT_CUR/VBA/dir"] = LegacyXlsCompoundFeatureEntryObjectType.Stream,
                ["_VBA_PROJECT_CUR/VBA/Sheet1"] = LegacyXlsCompoundFeatureEntryObjectType.Stream,
                ["_VBA_PROJECT_CUR/VBA/ThisWorkbook"] = LegacyXlsCompoundFeatureEntryObjectType.Stream,
                ["_VBA_PROJECT_CUR/VBA/LooseModule"] = LegacyXlsCompoundFeatureEntryObjectType.Stream,
                ["_VBA_PROJECT_CUR/VBA/_VBA_PROJECT"] = LegacyXlsCompoundFeatureEntryObjectType.Stream
            };
            var record = new LegacyXlsCompoundFeatureRecord(
                LegacyXlsCompoundFeatureRecordKind.VbaProject,
                entryRoles.Keys.ToArray(),
                entryRoles,
                entrySizes,
                entryObjectTypes);
            workbook.MutableCompoundFeatureRecords.Add(record);

            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(new[] { "Sheet1", "ThisWorkbook", "LooseModule" }, record.VbaModuleNames);
            Assert.Equal(3, record.VbaModuleCount);
            Assert.Equal(670, record.EntryByteCount);
            Assert.Equal(600, record.VbaModuleByteCount);
            Assert.Equal(3, report.CompoundVbaModuleCount);
            Assert.Equal(670, report.CompoundFeatureEntryByteCount);
            Assert.Equal(600, report.CompoundVbaModuleByteCount);
            Assert.Equal(1, report.CompoundVbaModulesByName["Sheet1"]);
            Assert.Equal(1, report.CompoundVbaModulesByName["ThisWorkbook"]);
            Assert.Equal(1, report.CompoundVbaModulesByName["LooseModule"]);
            Assert.Equal(1, report.CompoundVbaModulesByPath["_VBA_PROJECT_CUR/VBA/Sheet1"]);
            Assert.Equal(1, report.CompoundVbaModulesByPath["_VBA_PROJECT_CUR/VBA/ThisWorkbook"]);
            Assert.Equal(1, report.CompoundVbaModulesByPath["_VBA_PROJECT_CUR/VBA/LooseModule"]);
            Assert.Equal(2, report.CompoundFeatureEntriesByObjectType["Storage"]);
            Assert.Equal(5, report.CompoundFeatureEntriesByObjectType["Stream"]);
            Assert.Equal(3, report.CompoundFeatureEntriesByRoleAndObjectType["VbaModuleStream|Stream"]);
            Assert.Equal(2, report.CompoundFeatureEntriesBySize["Bytes:0"]);
            Assert.Equal(1, report.CompoundFeatureEntriesBySize["Bytes:100"]);
            Assert.Equal(1, report.CompoundFeatureEntriesBySize["Bytes:200"]);
            Assert.Equal(1, report.CompoundFeatureEntriesBySize["Bytes:300"]);
            Assert.Equal(1, report.CompoundVbaModulesBySize["Bytes:100"]);
            Assert.Equal(1, report.CompoundVbaModulesBySize["Bytes:200"]);
            Assert.Equal(1, report.CompoundVbaModulesBySize["Bytes:300"]);
            Assert.Equal(1, report.CompoundVbaModulesByNameAndSize["Sheet1|Bytes:100"]);
            Assert.Equal(1, report.CompoundVbaModulesByNameAndSize["ThisWorkbook|Bytes:200"]);
            Assert.Equal(1, report.CompoundVbaModulesByNameAndSize["LooseModule|Bytes:300"]);
            Assert.Equal(1, report.CompoundVbaModulesByCodeNameMatch["WorksheetCodeName"]);
            Assert.Equal(1, report.CompoundVbaModulesByCodeNameMatch["WorkbookCodeName"]);
            Assert.Equal(1, report.CompoundVbaModulesByCodeNameMatch["UnmatchedCodeName"]);
            Assert.Equal(1, report.CompoundVbaModulesByCodeNameMatchAndName["WorksheetCodeName|Sheet1"]);
            Assert.Equal(1, report.CompoundVbaModulesByCodeNameMatchAndName["WorkbookCodeName|ThisWorkbook"]);
            Assert.Equal(1, report.CompoundVbaModulesByCodeNameMatchAndName["UnmatchedCodeName|LooseModule"]);
            Assert.Equal(1, report.CompoundVbaProjectsByModuleCount["Modules:3"]);
            Assert.Equal(1, report.CompoundVbaProjectsByModuleByteCount["Bytes:600"]);
            Assert.Equal(1, report.CompoundVbaProjectsByStructure["Modules:3|DirStreams:1|ProjectStreams:1|Storages:2"]);
            Assert.Equal(1, report.VbaProjectWorkbookStates["BiffMarker:Present|NoMacrosMarker:Missing|CompoundProject:Present|Modules:Present"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("Compound VBA modules: 3", markdown);
            Assert.Contains("Compound VBA module bytes: 600", markdown);
            Assert.Contains("Compound VBA Modules By Name", markdown);
            Assert.Contains("Compound VBA Modules By Path", markdown);
            Assert.Contains("Compound VBA Modules By Size", markdown);
            Assert.Contains("Compound VBA Modules By Name And Size", markdown);
            Assert.Contains("Compound VBA Modules By CodeName Match", markdown);
            Assert.Contains("Compound VBA Modules By CodeName Match And Name", markdown);
            Assert.Contains("Compound VBA Projects By Module Count", markdown);
            Assert.Contains("Compound VBA Projects By Module Byte Count", markdown);
            Assert.Contains("Compound VBA Projects By Structure", markdown);
            Assert.Contains("VBA Project Workbook States", markdown);
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
