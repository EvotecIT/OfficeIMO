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
                ReportUnsupportedContent = true
            });

            LegacyXlsImportReport report = result.ImportReport;
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(1, report.ChartSheetCount);
            Assert.Equal(2, report.UnsupportedSheetCount);
            Assert.Equal(1, report.CellCount);
            Assert.Equal(0, report.FormulaCellCount);
            Assert.Equal(0, report.CommentCount);
            Assert.Equal(0, report.HyperlinkCount);
            Assert.Equal(0, report.DataValidationCount);
            Assert.Equal(0, report.ConditionalFormattingCount);
            Assert.Equal(0, report.AutoFilterCriteriaCount);
            Assert.Equal(2, report.UnsupportedFeatureCount);
            Assert.Equal(0, report.PreservedFeatureRecordCount);
            Assert.Equal(2, report.UnsupportedProjectionGapCount);
            Assert.False(report.HasImportErrors);
            Assert.True(report.HasUnsupportedFeatures);
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.MacroSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByKind[LegacyXlsUnsupportedSheetKind.VbaModuleSheet]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x01|MacroSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByType["0x06|VbaModuleSheet"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["Macro1"]);
            Assert.Equal(1, report.UnsupportedSheetsByName["Module1"]);
            Assert.Equal(1, report.ChartSheetsByType["0x02|ChartSheet"]);
            Assert.Equal(1, report.ChartSheetsByName["Chart1"]);
            Assert.Equal(1, report.WorksheetsByVisibility["Visible"]);
            Assert.Equal(2, report.UnsupportedSheetsByVisibility["Visible"]);
            Assert.Equal(1, report.ChartSheetsByVisibility["Visible"]);
            Assert.Equal(1, report.UnsupportedSheetsByKindAndVisibility["MacroSheet|Visible"]);
            Assert.Equal(1, report.UnsupportedSheetsByKindAndVisibility["VbaModuleSheet|Visible"]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.MacroSheet]);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.VbaModuleSheet]);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.ChartSheet, report.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.ChartSheet, report.PreservedFeatureRecordsByKind.Keys);
            Assert.Equal(1, report.UnsupportedProjectionGapsByKind[LegacyXlsUnsupportedFeatureKind.MacroSheet]);
            Assert.Equal(1, report.UnsupportedProjectionGapsByKind[LegacyXlsUnsupportedFeatureKind.VbaModuleSheet]);
            Assert.DoesNotContain(report.UnsupportedProjectionGapsByKind, entry => entry.Key == LegacyXlsUnsupportedFeatureKind.ChartSheet);
            Assert.Equal(1, report.UnsupportedFeaturesByCode["XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED"]);
            Assert.Equal(1, report.UnsupportedFeaturesByRecordType["MacroSheet|XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED|0x0085"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["MacroSheet|XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED|Sheet:MacroSheet"]);
            Assert.Equal(1, report.UnsupportedFeaturesByLocation["XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED|Macro1"]);
            Assert.Equal(1, report.DiagnosticsByCode["XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Worksheets: 1", markdown);
            Assert.Contains("Chart sheets: 1", markdown);
            Assert.Contains("Unsupported sheets: 2", markdown);
            Assert.Contains("Unsupported projection gaps: 2", markdown);
            Assert.Contains("XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED", markdown);
            Assert.Contains("Unsupported Feature Record Types", markdown);
            Assert.Contains("Unsupported Feature Details", markdown);
            Assert.DoesNotContain("Preserved Feature Records By Kind", markdown);
            Assert.Contains("Unsupported Sheets By Kind", markdown);
            Assert.Contains("Unsupported Sheets By Visibility", markdown);
            Assert.Contains("Chart Sheets By Name", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_NamesPreserveOnlyFeatureDetails() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5PreserveOnlyFeatureDetailsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(1, report.WorksheetCount);
            Assert.Equal(1, report.UnsupportedFeatureCount);
            Assert.Equal(1, report.PreservedFeatureRecordCount);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.DrawingObject, report.UnsupportedFeaturesByKind.Keys);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.Chart, report.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.DrawingObject, report.PreservedFeatureRecordsByKind.Keys);
            Assert.Equal(1, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.Chart, report.PreservedFeatureRecordsByKind.Keys);
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
            Assert.Equal(1, report.DrawingOfficeArtRecordsByPayloadLength["PayloadLength:16"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByPayloadLength["PayloadLength:22"]);
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
            Assert.Equal(3, report.DrawingShapePropertyCount);
            Assert.Equal(1, report.DrawingShapePropertiesById["PropertyId:0x0104"]);
            Assert.Equal(1, report.DrawingShapePropertiesById["PropertyId:0x00BF"]);
            Assert.Equal(1, report.DrawingShapePropertiesById["PropertyId:0x0005"]);
            Assert.Equal(1, report.DrawingShapePropertiesByName["pib"]);
            Assert.Equal(1, report.DrawingShapePropertiesByName["TextBooleanProperties"]);
            Assert.Equal(1, report.DrawingShapePropertiesByName["PropertyId:0x0005"]);
            Assert.Equal(1, report.DrawingShapePropertiesByGroup["Blip"]);
            Assert.Equal(1, report.DrawingShapePropertiesByGroup["Text"]);
            Assert.Equal(1, report.DrawingShapePropertiesByGroup["Protection"]);
            Assert.Equal(1, report.DrawingShapePropertiesByFlagState["Simple"]);
            Assert.Equal(1, report.DrawingShapePropertiesByFlagState["Blip"]);
            Assert.Equal(1, report.DrawingShapePropertiesByFlagState["Complex"]);
            Assert.Equal(1, report.DrawingShapePropertiesByValue["PropertyId:0x0104;Value:0x00000001"]);
            Assert.Equal(1, report.DrawingShapePropertiesByValue["PropertyId:0x00BF;Value:0x00000001"]);
            Assert.Equal(1, report.DrawingShapeComplexPropertiesByDeclaredLength["PropertyId:0x0005;DeclaredBytes:4"]);
            Assert.Equal(1, report.DrawingShapeComplexPropertiesByAvailableLength["PropertyId:0x0005;AvailableBytes:4"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByType["Png"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByEmbeddedRecordType["OfficeArtBlipPNG"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesBySize["SizeBytes:75"]);
            Assert.Equal(1, report.DrawingBlipStoreEntriesByReferenceCount["References:1"]);
            Assert.Equal(1, report.DrawingPictureStates["PictureObjects:Present|BlipStore:Present|PictureBlipReferences:Present|ReferencedBlips:Resolved"]);
            Assert.Equal(1, report.DrawingPictureCountStates["PictureObjects:1|BlipStoreEntries:1|PictureBlipReferences:1|PictureFrames:1|ObjectBlipParity:Balanced|ObjectFrameCoverage:Balanced|ReferencedBlips:Resolved"]);
            Assert.Equal(1, report.DrawingShapeEntriesByType["PictureFrame"]);
            Assert.Equal(1, report.DrawingShapeEntriesById["ShapeId:1024"]);
            Assert.Equal(1, report.DrawingShapeEntriesByFlags["Flags:0x00000A02"]);
            Assert.Equal(1, report.DrawingShapeEntriesByReservedState["ReservedClear"]);
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
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawingGroup", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawing", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:ShapePropsStream", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:TextPropsStream", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:RichTextStream", report.UnsupportedFeaturesByDetail.Keys);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView"]);
            Assert.DoesNotContain("Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawingGroup", report.PreservedFeatureRecordsByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj", report.PreservedFeatureRecordsByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:MsoDrawing", report.PreservedFeatureRecordsByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:ShapePropsStream", report.PreservedFeatureRecordsByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:TextPropsStream", report.PreservedFeatureRecordsByDetail.Keys);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:RichTextStream", report.PreservedFeatureRecordsByDetail.Keys);
            Assert.Equal(1, report.PreservedFeatureRecordsByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView"]);
            Assert.DoesNotContain("Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|Chart:Chart", report.PreservedFeatureRecordsByDetail.Keys);
            Assert.DoesNotContain(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:MsoDrawingGroup" && record.SheetName == null);
            Assert.DoesNotContain(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:Obj" && record.SheetName == "FeatureMap");
            Assert.DoesNotContain(workbook.PreservedFeatureRecords, record => record.DetailCode == "Drawing:ShapePropsStream" && record.SheetName == "FeatureMap");
            Assert.Contains(workbook.DrawingRecords, record => record.SheetName == "FeatureMap" && record.ObjectType == 0x0008 && record.ObjectTypeKind == LegacyXlsDrawingObjectType.Picture && record.ObjectTypeName == "Picture" && record.ObjectId == 1 && record.ObjectFlags == 0x4011 && record.IsObjectLocked && record.IsObjectPrintable);
            LegacyXlsDrawingRecord objectRecord = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "Obj");
            Assert.True(objectRecord.HasSupportedObjectMetadata);
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
            Assert.Equal((uint)159, drawingGroup.EscherPayloadLength);
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
            Assert.Equal((uint)75, blipEntry.SizeBytes);
            Assert.Equal((uint)1, blipEntry.ReferenceCount);
            Assert.Equal((ushort)0xf01e, blipEntry.EmbeddedBlipRecordType);
            Assert.Equal("OfficeArtBlipPNG", blipEntry.EmbeddedBlipRecordTypeName);
            Assert.Equal((uint)67, blipEntry.EmbeddedBlipPayloadLength);
            Assert.Equal("image/png", blipEntry.EmbeddedBlipContentType);
            Assert.True(blipEntry.HasEmbeddedBlipPayloadBytes);
            LegacyXlsDrawingRecord drawing = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "MsoDrawing");
            Assert.Equal((ushort)0xf002, drawing.EscherRecordType);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer, drawing.EscherRecordTypeKind);
            Assert.Equal("OfficeArtDgContainer", drawing.EscherRecordTypeName);
            Assert.Equal((ushort)1, drawing.EscherRecordInstance);
            Assert.Equal((byte)0x0f, drawing.EscherRecordVersion);
            Assert.Equal((uint)120, drawing.EscherPayloadLength);
            Assert.True(drawing.OfficeArtPayloadFullyTraversed);
            Assert.True(drawing.HasSupportedOfficeArtMetadata);
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
            Assert.Equal(3, drawing.ShapeProperties.Count);
            LegacyXlsDrawingShapeProperty blipProperty = drawing.ShapeProperties[0];
            Assert.Equal(0, blipProperty.Index);
            Assert.Equal((ushort)0x4104, blipProperty.RawOperationId);
            Assert.Equal((ushort)0x0104, blipProperty.PropertyId);
            Assert.Equal("PropertyId:0x0104", blipProperty.PropertyIdKey);
            Assert.Equal("pib", blipProperty.PropertyName);
            Assert.Equal("Blip", blipProperty.PropertyGroupName);
            Assert.True(blipProperty.IsBlipId);
            Assert.False(blipProperty.IsComplex);
            Assert.Equal((uint)1, blipProperty.Value);
            Assert.Null(blipProperty.DeclaredComplexDataLength);
            Assert.Null(blipProperty.AvailableComplexDataLength);
            LegacyXlsDrawingShapeProperty simpleProperty = drawing.ShapeProperties[1];
            Assert.Equal(1, simpleProperty.Index);
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
            LegacyXlsDrawingShapeProperty complexProperty = drawing.ShapeProperties[2];
            Assert.Equal(2, complexProperty.Index);
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
            Assert.Equal(0u, shape.ReservedFlags);
            Assert.False(shape.HasReservedFlags);
            Assert.Equal("ReservedClear", shape.ReservedState);
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
            Assert.True(shapeStream.HasSupportedFutureDrawingStreamMetadata);
            Assert.Equal((ushort)0x08a3, shapeStream.FutureRecordHeader?.WrappedRecordType);
            Assert.False(shapeStream.FutureRecordHeader?.HasRange);
            Assert.Equal(5, shapeStream.FutureRecordHeader?.StreamByteCount);
            LegacyXlsDrawingRecord textStream = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "TextPropsStream");
            Assert.Equal(LegacyXlsDrawingRecordKind.TextPropertiesStream, textStream.Kind);
            Assert.True(textStream.HasFutureRecordHeader);
            Assert.True(textStream.HasSupportedFutureDrawingStreamMetadata);
            Assert.Equal((ushort)0x08a4, textStream.FutureRecordHeader?.WrappedRecordType);
            Assert.True(textStream.FutureRecordHeader?.HasRange);
            Assert.True(textStream.FutureRecordHeader?.HasCompleteRangeReference);
            Assert.Equal((ushort)1, textStream.FutureRecordHeader?.FirstRow);
            Assert.Equal((ushort)2, textStream.FutureRecordHeader?.LastRow);
            Assert.Equal((ushort)0, textStream.FutureRecordHeader?.FirstColumn);
            Assert.Equal((ushort)1, textStream.FutureRecordHeader?.LastColumn);
            Assert.Equal(4, textStream.FutureRecordHeader?.StreamByteCount);
            LegacyXlsDrawingRecord richStream = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "RichTextStream");
            Assert.Equal(LegacyXlsDrawingRecordKind.RichTextStream, richStream.Kind);
            Assert.True(richStream.HasFutureRecordHeader);
            Assert.True(richStream.HasSupportedFutureDrawingStreamMetadata);
            Assert.Equal((ushort)0x08a5, richStream.FutureRecordHeader?.WrappedRecordType);
            Assert.False(richStream.FutureRecordHeader?.HasRange);
            Assert.Equal(4, richStream.FutureRecordHeader?.StreamByteCount);
            Assert.DoesNotContain(workbook.PreservedFeatureRecords, record => record.DetailCode == "Chart:Chart" && record.RecordType == 0x1002);
            Assert.DoesNotContain(workbook.Diagnostics, d => d.DetailCode == "Chart:Chart");
            Assert.Contains(workbook.Diagnostics, d => d.DetailCode == "PivotTable:SxView");
            string markdown = report.ToMarkdown();
            Assert.Contains("Preserved feature records: 1", markdown);
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
            Assert.Contains("Drawing shape properties: 3", markdown);
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

        [Fact]
        public void LegacyXls_ImportReport_TracksTableStyleMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5TableStyleMetadataWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsWorkbook workbook = result.Workbook;
            LegacyXlsImportReport report = result.ImportReport;

            Assert.DoesNotContain(workbook.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsTableStyleCollection collection = Assert.Single(workbook.TableStyleCollections);
            Assert.Equal(145U, collection.TotalStyleCount);
            Assert.Equal("TableStyleMedium2", collection.DefaultTableStyleName);
            Assert.Equal("PivotStyleLight16", collection.DefaultPivotStyleName);

            LegacyXlsTableStyle style = Assert.Single(workbook.TableStyles);
            Assert.Equal("OfficeIMO Custom", style.Name);
            Assert.True(style.AppliesToTables);
            Assert.True(style.AppliesToPivotTables);
            Assert.Equal(2U, style.DeclaredElementCount);
            Assert.Collection(style.Elements,
                element => {
                    Assert.Equal("HeaderRow", element.ElementTypeName);
                    Assert.Equal(0U, element.StripeSize);
                    Assert.Equal(3U, element.DifferentialFormatIndex);
                },
                element => {
                    Assert.Equal("RowStripe1", element.ElementTypeName);
                    Assert.Equal(2U, element.StripeSize);
                    Assert.Equal(4U, element.DifferentialFormatIndex);
                });

            Assert.Equal(1, report.TableStyleCollectionRecordCount);
            Assert.Equal(1, report.TableStyleDefinitionCount);
            Assert.Equal(2, report.TableStyleElementRecordCount);
            Assert.Equal(1, report.TableStyleCollectionsByDefaultTableStyle["TableStyleMedium2"]);
            Assert.Equal(1, report.TableStyleCollectionsByDefaultPivotStyle["PivotStyleLight16"]);
            Assert.Equal(1, report.TableStyleCollectionsByTotalStyleCount["Styles:145"]);
            Assert.Equal(1, report.TableStylesByName["OfficeIMO Custom"]);
            Assert.Equal(1, report.TableStylesByApplicability["TableAndPivot"]);
            Assert.Equal(1, report.TableStylesByDeclaredElementCount["Declared:2"]);
            Assert.Equal(1, report.TableStylesByParsedElementCount["Parsed:2"]);
            Assert.Equal(1, report.TableStyleElementsByType["HeaderRow"]);
            Assert.Equal(1, report.TableStyleElementsByType["RowStripe1"]);
            Assert.Equal(1, report.TableStyleElementsByDifferentialFormatIndex["Dxf:3"]);
            Assert.Equal(1, report.TableStyleElementsByDifferentialFormatIndex["Dxf:4"]);
            Assert.Equal(1, report.TableStyleElementsByStripeSize["Size:2"]);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.TableStyle, report.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.TableStyle, report.PreservedFeatureRecordsByKind.Keys);

            DocumentFormat.OpenXml.Spreadsheet.TableStyles projectedTableStyles = result.Document.WorkbookPartRoot
                .WorkbookStylesPart!
                .Stylesheet!
                .TableStyles!;
            Assert.Equal("TableStyleMedium2", projectedTableStyles.DefaultTableStyle!.Value);
            Assert.Equal("PivotStyleLight16", projectedTableStyles.DefaultPivotStyle!.Value);
            Assert.Equal(1U, projectedTableStyles.Count!.Value);

            DocumentFormat.OpenXml.Spreadsheet.TableStyle projectedStyle = Assert.Single(projectedTableStyles.Elements<DocumentFormat.OpenXml.Spreadsheet.TableStyle>());
            Assert.Equal("OfficeIMO Custom", projectedStyle.Name!.Value);
            Assert.True(projectedStyle.Table!.Value);
            Assert.True(projectedStyle.Pivot!.Value);
            Assert.Equal(2U, projectedStyle.Count!.Value);
            Assert.Collection(
                projectedStyle.Elements<DocumentFormat.OpenXml.Spreadsheet.TableStyleElement>(),
                element => {
                    Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.TableStyleValues.HeaderRow, element.Type!.Value);
                    Assert.Equal(3U, element.FormatId!.Value);
                },
                element => {
                    Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.TableStyleValues.FirstRowStripe, element.Type!.Value);
                    Assert.Equal(2U, element.Size!.Value);
                    Assert.Equal(4U, element.FormatId!.Value);
                });
        }

        [Fact]
        public void LegacyXls_ImportReport_ProjectsCustomTableStylesWithoutDefaults() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5CustomTableStyleOnlyWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(result.Workbook.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Empty(result.Workbook.TableStyleCollections);
            LegacyXlsTableStyle style = Assert.Single(result.Workbook.TableStyles);
            Assert.Equal("OfficeIMO Custom Only", style.Name);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.TableStyle, result.ImportReport.UnsupportedFeaturesByKind.Keys);

            DocumentFormat.OpenXml.Spreadsheet.TableStyles projectedTableStyles = result.Document.WorkbookPartRoot
                .WorkbookStylesPart!
                .Stylesheet!
                .TableStyles!;
            Assert.Equal(1U, projectedTableStyles.Count!.Value);

            DocumentFormat.OpenXml.Spreadsheet.TableStyle projectedStyle = Assert.Single(projectedTableStyles.Elements<DocumentFormat.OpenXml.Spreadsheet.TableStyle>());
            Assert.Equal("OfficeIMO Custom Only", projectedStyle.Name!.Value);
            DocumentFormat.OpenXml.Spreadsheet.TableStyleElement projectedElement = Assert.Single(projectedStyle.Elements<DocumentFormat.OpenXml.Spreadsheet.TableStyleElement>());
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.TableStyleValues.HeaderRow, projectedElement.Type!.Value);
            Assert.Equal(3U, projectedElement.FormatId!.Value);
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

        [Fact]
        public void LegacyXlsDrawingTextObject_DecodesAlignmentRotationAndFlags() {
            var textObject = new LegacyXlsDrawingTextObject(
                0xc212,
                rotation: 2,
                textCharacterCount: 19,
                formattingRunByteCount: 16,
                emptyFontIndex: 0,
                formulaByteCount: 4);

            Assert.Equal("Left", textObject.HorizontalAlignmentName);
            Assert.Equal("Top", textObject.VerticalAlignmentName);
            Assert.Equal("RotatedCounterClockwise90", textObject.RotationName);
            Assert.True(textObject.LockedText);
            Assert.True(textObject.JustifyLastLine);
            Assert.True(textObject.SecretEdit);
            Assert.True(textObject.HasTextInContinueRecords);
            Assert.True(textObject.HasFormattingRunsInContinueRecords);
            Assert.Equal(4, textObject.FormulaByteCount);
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
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(32, workbook.PivotTableRecords.Count);
            Assert.Equal(32, report.PivotTableRecordCount);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.View]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Field]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.FieldIndexList]);
            Assert.Equal(2, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.LineItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.PageItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Item]);
            Assert.Equal(2, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.DataItem]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Cache]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.CacheStream]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.CacheSource]);
            Assert.Equal(8, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.CacheItem]);
            Assert.Equal(3, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Formula]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Rule]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Table]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.GroupingRange]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Filter]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Format]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.ExtendedPivotField]);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.PivotChart]);
            Assert.Equal(2, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.Additional]);
            Assert.DoesNotContain(report.PivotTableRecordsByKind, entry => entry.Key == LegacyXlsPivotTableRecordKind.PreserveOnly);
            Assert.Equal(29, workbook.PivotTableRecords.Count(record => record.HasSupportedPivotTableMetadata));
            Assert.Equal(3, workbook.PivotTableRecords.Count(record => !record.HasSupportedPivotTableMetadata));
            Assert.Equal(6, report.FormulaTokenRecordCount);
            Assert.Equal(3, report.FormulaTokensByContext["PivotTableFormula"]);
            Assert.Equal(3, report.FormulaTokensByContext["PivotTableCalculatedFieldFormula"]);
            Assert.Equal(3, report.FormulaTokensByContextAndSheet["PivotTableFormula|(workbook)"]);
            Assert.Equal(3, report.FormulaTokensByContextAndSheet["PivotTableCalculatedFieldFormula|(workbook)"]);
            Assert.Equal(3, report.FormulaTokensByName["PtgInt"]);
            Assert.Equal(1, report.FormulaTokensByName["PtgAdd"]);
            Assert.Equal(1, report.FormulaTokensByName["PtgLe"]);
            Assert.Equal(1, report.FormulaTokensByName["FormulaToken0xFF"]);
            Assert.Equal(2, report.FormulaTokensByContextAndOperandKind["PivotTableFormula|IntegerLiteral"]);
            Assert.Equal(1, report.FormulaTokensByContextAndOperandKind["PivotTableCalculatedFieldFormula|IntegerLiteral"]);
            Assert.Equal(1, report.FormulaTokensByNameAndOperandText["PtgInt|7"]);
            Assert.Equal(1, report.FormulaTokensByNameAndOperandText["PtgInt|8"]);
            Assert.Equal(1, report.FormulaTokensByNameAndOperandText["PtgInt|40980"]);
            Assert.Equal(0, report.UnsupportedProjectionGapCount);
            Assert.Empty(report.UnsupportedProjectionGapsByKind);
            Assert.Empty(report.UnsupportedProjectionGapsByRecordType);
            Assert.Empty(report.UnsupportedProjectionGapsByDetail);
            Assert.Equal(1, report.PivotTableRecordsByName["SxView"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvd"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxIvd"]);
            Assert.Equal(2, report.PivotTableRecordsByName["Sxli"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxpi"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvi"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxdi"]);
            Assert.Equal(2, report.PivotTableRecordsByName["Sxdb"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxStreamId"]);
            Assert.Equal(1, report.PivotTableRecordsByName["Sxvs"]);
            Assert.Equal(2, report.PivotTableRecordsByName["SxFormula"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxFmla"]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxRule"]);
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
            Assert.Equal(17, report.PivotTableRecordsByLocation["(workbook)"]);
            Assert.Equal(15, report.PivotTableRecordsByLocation["PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByKindAndLocation["View|(workbook)"]);
            Assert.Equal(2, report.PivotTableRecordsByKindAndLocation["DataItem|(workbook)"]);
            Assert.Equal(3, report.PivotTableRecordsByKindAndLocation["Formula|(workbook)"]);
            Assert.Equal(5, report.PivotTableRecordsByKindAndLocation["CacheItem|(workbook)"]);
            Assert.Equal(3, report.PivotTableRecordsByKindAndLocation["CacheItem|PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByKindAndLocation["Field|PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByKindAndLocation["FieldIndexList|PivotMeta"]);
            Assert.Equal(2, report.PivotTableRecordsByKindAndLocation["LineItem|PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByKindAndLocation["PageItem|PivotMeta"]);
            Assert.Equal(2, report.PivotTableRecordsByKindAndLocation["Additional|PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByNameAndLocation["SxView|(workbook)"]);
            Assert.Equal(2, report.PivotTableRecordsByNameAndLocation["Sxdb|(workbook)"]);
            Assert.Equal(2, report.PivotTableRecordsByNameAndLocation["SxFormula|(workbook)"]);
            Assert.Equal(1, report.PivotTableRecordsByNameAndLocation["SxFmla|(workbook)"]);
            Assert.Equal(1, report.PivotTableRecordsByNameAndLocation["SxRule|(workbook)"]);
            Assert.Equal(1, report.PivotTableRecordsByNameAndLocation["SxIvd|PivotMeta"]);
            Assert.Equal(2, report.PivotTableRecordsByNameAndLocation["Sxli|PivotMeta"]);
            Assert.Equal(1, report.PivotTableRecordsByNameAndLocation["Sxpi|PivotMeta"]);
            Assert.Equal(2, report.PivotTableRecordsByNameAndLocation["SxAddl|PivotMeta"]);
            Assert.Equal(1, report.PivotTableWorkbookStates["View:Present|Cache:Present|CacheSource:Present|CacheItems:Present|Fields:Present|Items:Present|DataItems:Present|Grouping:Present|Formulas:Present|Additional:Present|Locations:WorkbookAndSheets"]);
            Assert.Equal(1, report.PivotTableViewRanges["A3:C11"]);
            Assert.Equal(1, report.PivotTableViewNames["SalesPivot"]);
            Assert.Equal(1, report.PivotTableViewDataNames["Data"]);
            Assert.Equal(1, report.PivotTableViewFieldCounts["Fields:5;Rows:2;Columns:0;Pages:1;Data:1"]);
            Assert.Equal(1, report.PivotTableViewLineCounts["Rows:7;Columns:1"]);
            Assert.Equal(1, report.PivotTableViewDataAxes["Row"]);
            Assert.Equal(1, report.PivotTableViewDataPositions["Position:-1"]);
            Assert.Equal(1, report.PivotTableViewCacheIndexes["CacheIndex:0"]);
            Assert.Equal(1, report.PivotTableViewGrandTotalStates["Rows:True;Columns:True"]);
            Assert.Equal(1, report.PivotTableViewAutoFormatStates["AutoFormat:True;Id:1"]);
            Assert.Equal(1, report.PivotTableFieldAxes["Row"]);
            Assert.Equal(1, report.PivotTableFieldItemCounts["Items:3"]);
            Assert.Equal(1, report.PivotTableFieldSubtotalCounts["Subtotals:1;Flags:0x0001"]);
            Assert.Equal(1, report.PivotTableFieldSubtotalFunctions["Default"]);
            Assert.Equal(1, report.PivotTableFieldNames["Region"]);
            Assert.Equal(1, report.PivotTableFieldIndexListLengths["Indexes:2"]);
            Assert.Equal(1, report.PivotTableFieldIndexReferences["FieldIndex:0"]);
            Assert.Equal(1, report.PivotTableFieldIndexReferences["FieldIndex:2"]);
            Assert.Equal(1, report.PivotTableFieldIndexSequences["FieldIndexes:0,2"]);
            Assert.Equal(2, report.PivotTableLineItemCounts["LineItems:1"]);
            Assert.Equal(1, report.PivotTableLineItemTypes["LineItemType:0"]);
            Assert.Equal(1, report.PivotTableLineItemTypes["LineItemType:13"]);
            Assert.Equal(1, report.PivotTableLineItemTypeKinds["Data"]);
            Assert.Equal(1, report.PivotTableLineItemTypeKinds["GrandTotal"]);
            Assert.Equal(1, report.PivotTableLineItemEntryCounts["Entries:1"]);
            Assert.Equal(1, report.PivotTableLineItemEntryCounts["Entries:2"]);
            Assert.Equal(1, report.PivotTableLineItemEntrySlotCounts["Slots:1;Entries:1"]);
            Assert.Equal(1, report.PivotTableLineItemEntrySlotCounts["Slots:2;Entries:2"]);
            Assert.Equal(1, report.PivotTableLineItemEntryIndexes["EntryIndex:0"]);
            Assert.Equal(1, report.PivotTableLineItemEntryIndexes["EntryIndex:2"]);
            Assert.Equal(1, report.PivotTableLineItemEntryIndexes["BlankEntry"]);
            Assert.Equal(2, report.PivotTableLineItemDataIndexes["DataIndex:0"]);
            Assert.Equal(1, report.PivotTableLineItemFlagStates["Subtotal:False;Block:False;Grand:False;MultiDataName:False;MultiDataOnAxis:False"]);
            Assert.Equal(1, report.PivotTableLineItemFlagStates["Subtotal:False;Block:False;Grand:True;MultiDataName:False;MultiDataOnAxis:False"]);
            Assert.Equal(1, report.PivotTableLineItemSequences["LineItems:Type:Data;Entries:EntryIndex:0,EntryIndex:2"]);
            Assert.Equal(1, report.PivotTableLineItemSequences["LineItems:Type:GrandTotal;Entries:BlankEntry"]);
            Assert.Equal(1, report.PivotTablePageItemCounts["PageItems:1"]);
            Assert.Equal(1, report.PivotTablePageItemFieldIndexes["FieldIndex:1"]);
            Assert.Equal(1, report.PivotTablePageItemIndexes["AllItems"]);
            Assert.Equal(1, report.PivotTablePageItemObjectIds["ObjectId:42"]);
            Assert.Equal(1, report.PivotTablePageItemSequences["PageItems:FieldIndex:1;AllItems;ObjectId:42"]);
            Assert.Equal(1, report.PivotTableItemTypes["ItemType:0"]);
            Assert.Equal(1, report.PivotTableItemTypeKinds["Data"]);
            Assert.Equal(1, report.PivotTableItemCacheIndexes["CacheItem:1"]);
            Assert.Equal(1, report.PivotTableItemFlagStates["Hidden:False;HideDetail:True;Formula:False;Missing:False"]);
            Assert.Equal(1, report.PivotTableItemNames["East"]);
            Assert.Equal(1, report.PivotTableFormulaPayloadLengths["SxFormula|Bytes:4"]);
            Assert.Equal(1, report.PivotTableFormulaPayloadLengths["SxFormula|Bytes:20"]);
            Assert.Equal(1, report.PivotTableFormulaPayloadLengths["SxFmla|Bytes:9"]);
            Assert.Equal(1, report.PivotTableFormulaPayloadKinds["Scope"]);
            Assert.Equal(1, report.PivotTableFormulaPayloadKinds["FormulaTokens"]);
            Assert.Equal(1, report.PivotTableFormulaPayloadKinds["CalculatedFieldFormulaTokens"]);
            Assert.Equal(1, report.PivotTableFormulaTokenByteCounts["TokenBytes:7"]);
            Assert.Equal(1, report.PivotTableCalculatedFieldFormulaTokenByteCounts["TokenBytes:20"]);
            Assert.Equal(1, report.PivotTableFormulaTrailingByteCounts["TrailingBytes:0"]);
            Assert.Equal(1, report.PivotTableRuleAxes["Row"]);
            Assert.Equal(1, report.PivotTableRuleTypes["DataCells"]);
            Assert.Equal(1, report.PivotTableRuleFieldReferences["FilteredFields"]);
            Assert.Equal(1, report.PivotTableRuleFilterCounts["Filters:2"]);
            Assert.Equal(1, report.PivotTableRuleOptionStates["Partial:False;DataOnly:False;LabelOnly:False;CacheBased:False"]);
            Assert.Equal(1, report.PivotTableRuleFilterEntryCounts["Filters:2"]);
            Assert.Equal(1, report.PivotTableRuleFilterAxes["Row"]);
            Assert.Equal(1, report.PivotTableRuleFilterAxes["Data"]);
            Assert.Equal(2, report.PivotTableRuleFilterFieldPositions["Position:0"]);
            Assert.Equal(1, report.PivotTableRuleFilterFieldReferences["FieldIndex:0"]);
            Assert.Equal(1, report.PivotTableRuleFilterFieldReferences["DataField"]);
            Assert.Equal(1, report.PivotTableRuleFilterSelectedStates["Selected:True"]);
            Assert.Equal(1, report.PivotTableRuleFilterSelectedStates["Selected:False"]);
            Assert.Equal(1, report.PivotTableRuleFilterSubtotalFlags["Flags:0x0001"]);
            Assert.Equal(1, report.PivotTableRuleFilterSubtotalFlags["Flags:0x0002"]);
            Assert.Equal(1, report.PivotTableRuleFilterSubtotalFunctions["Data"]);
            Assert.Equal(1, report.PivotTableRuleFilterSubtotalFunctions["Default"]);
            Assert.Equal(1, report.PivotTableRuleFilterItemIndexCounts["Indexes:1"]);
            Assert.Equal(1, report.PivotTableRuleFilterItemIndexCounts["Indexes:0"]);
            Assert.Equal(1, report.PivotTableRuleFilterStates["Axis:Row|Position:0|Field:FieldIndex:0|Selected:True|Subtotals:0x0001|Indexes:1"]);
            Assert.Equal(1, report.PivotTableRuleFilterStates["Axis:Data|Position:0|Field:DataField|Selected:False|Subtotals:0x0002|Indexes:0"]);
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
            Assert.Equal(2, report.PivotTableDataItemAggregations["AggregationFunction:0"]);
            Assert.Equal(2, report.PivotTableDataItemAggregationKinds["Sum"]);
            Assert.Equal(1, report.PivotTableDataItemFieldIndexes["FieldIndex:2"]);
            Assert.Equal(1, report.PivotTableDataItemFieldIndexes["FieldIndex:3"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculationIds["DisplayCalculation:0"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculationIds["DisplayCalculation:7"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculations["Value"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculations["PercentOfGrandTotal"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculationReferenceStates["Value|Field:NoFieldReference|Item:NoItemReference"]);
            Assert.Equal(1, report.PivotTableDataItemDisplayCalculationReferenceStates["PercentOfGrandTotal|Field:NoFieldReference|Item:NoItemReference"]);
            Assert.Equal(2, report.PivotTableDataItemDisplayCalculationFieldIndexes["FieldIndex:-1"]);
            Assert.Equal(2, report.PivotTableDataItemDisplayCalculationItemIndexes["ItemIndex:-1"]);
            Assert.Equal(1, report.PivotTableDataItemNumberFormats["NumberFormatId:0"]);
            Assert.Equal(1, report.PivotTableDataItemNumberFormats["NumberFormatId:14"]);
            Assert.Equal(1, report.PivotTableDataItemNames["Sales"]);
            Assert.Equal(1, report.PivotTableDataItemNames["Sum of Amount"]);
            Assert.Equal(1, report.PivotTableGroupingKinds["Months"]);
            Assert.Equal(1, report.PivotTableGroupingBoundaryStates["AutoStart:True;AutoEnd:True"]);
            Assert.Equal(1, report.PivotTableGroupingCompletionStates["CompleteDateRange"]);
            Assert.Equal(1, report.PivotTableGroupingStates["Kind:Months|AutoStart:True|AutoEnd:True|CompleteDateRange"]);
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
            Assert.Equal(1, report.PivotTableExtendedFieldPermissionStates["ShowAllItems:True|Row:True|Column:True|Page:True|Hide:True|PreventData:False|ServerBased:True"]);
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
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.FieldIndexList && record.RecordName == "SxIvd" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.LineItem && record.RecordName == "Sxli" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.PageItem && record.RecordName == "Sxpi" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Item && record.RecordName == "Sxvi" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Cache && record.RecordName == "Sxdb" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.CacheStream && record.RecordName == "SxStreamId" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.CacheSource && record.RecordName == "Sxvs" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Formula && record.RecordName == "SxFormula" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Formula && record.RecordName == "SxFmla" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Rule && record.RecordName == "SxRule" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.CacheItem && record.RecordName == "Sxnum" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Table && record.RecordName == "Sxtbl" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Filter && record.RecordName == "SxFilt" && record.SheetName == null);
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Format && record.RecordName == "SxFormat" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.PivotChart && record.RecordName == "PivotChartBits" && record.SheetName == "PivotMeta");
            Assert.Contains(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Additional && record.RecordName == "SxAddl" && record.SheetName == "PivotMeta");

            LegacyXlsPivotTableRecord view = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.View);
            Assert.Null(view.SheetName);
            Assert.Equal("SxView", view.RecordName);
            Assert.Equal("A3:C11", view.ViewRange);
            Assert.Equal((ushort)2, view.ViewFirstRow);
            Assert.Equal((ushort)10, view.ViewLastRow);
            Assert.Equal((ushort)0, view.ViewFirstColumn);
            Assert.Equal((ushort)2, view.ViewLastColumn);
            Assert.Equal((ushort)4, view.ViewFirstHeaderRow);
            Assert.Equal((ushort)4, view.ViewFirstDataRow);
            Assert.Equal((ushort)2, view.ViewFirstDataColumn);
            Assert.Equal((short)0, view.ViewCacheIndex);
            Assert.Equal("Row", view.ViewDataAxisName);
            Assert.Equal((short)-1, view.ViewDataPosition);
            Assert.Equal((short)5, view.ViewFieldCount);
            Assert.Equal((ushort)2, view.ViewRowFieldCount);
            Assert.Equal((ushort)0, view.ViewColumnFieldCount);
            Assert.Equal((ushort)1, view.ViewPageFieldCount);
            Assert.Equal((short)1, view.ViewDataFieldCount);
            Assert.Equal((ushort)7, view.ViewRowLineCount);
            Assert.Equal((ushort)1, view.ViewColumnLineCount);
            Assert.True(view.ViewRowGrandTotals);
            Assert.True(view.ViewColumnGrandTotals);
            Assert.True(view.ViewAutoFormat);
            Assert.Equal((ushort)1, view.ViewAutoFormatId);
            Assert.Equal("SalesPivot", view.ViewTableName);
            Assert.Equal("Data", view.ViewDataName);

            LegacyXlsPivotTableRecord field = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Field);
            Assert.Equal("PivotMeta", field.SheetName);
            Assert.Equal("Sxvd", field.RecordName);
            Assert.Equal("Row", field.FieldAxisName);
            Assert.Equal((ushort)1, field.FieldSubtotalCount);
            Assert.Equal((ushort)0x0001, field.FieldSubtotalFlags);
            Assert.Equal("Default", Assert.Single(field.FieldSubtotalFunctionNames));
            Assert.Equal((short)3, field.FieldItemCount);
            Assert.Equal("Region", field.FieldName);

            LegacyXlsPivotTableRecord fieldIndexList = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.FieldIndexList);
            Assert.Equal("PivotMeta", fieldIndexList.SheetName);
            Assert.Equal("SxIvd", fieldIndexList.RecordName);
            Assert.Equal(new short[] { 0, 2 }, fieldIndexList.FieldIndexReferences);

            LegacyXlsPivotTableRecord[] lineItemRecords = workbook.PivotTableRecords
                .Where(record => record.Kind == LegacyXlsPivotTableRecordKind.LineItem)
                .ToArray();
            Assert.Equal(2, lineItemRecords.Length);
            LegacyXlsPivotLineItem dataLineItem = Assert.Single(lineItemRecords[0].LineItems);
            Assert.Equal((ushort)0, dataLineItem.ItemType);
            Assert.Equal(LegacyXlsPivotLineItemType.Data, dataLineItem.ItemTypeKind);
            Assert.Equal("Data", dataLineItem.ItemTypeName);
            Assert.Equal(new short[] { 0, 2 }, dataLineItem.EntryIndexes);
            Assert.Equal(2, dataLineItem.EntrySlotCount);
            Assert.False(dataLineItem.Subtotal);
            Assert.False(dataLineItem.GrandTotal);
            LegacyXlsPivotLineItem grandLineItem = Assert.Single(lineItemRecords[1].LineItems);
            Assert.Equal((ushort)13, grandLineItem.ItemType);
            Assert.Equal(LegacyXlsPivotLineItemType.GrandTotal, grandLineItem.ItemTypeKind);
            Assert.Equal("GrandTotal", grandLineItem.ItemTypeName);
            Assert.Equal("BlankEntry", Assert.Single(grandLineItem.EntryIndexNames));
            Assert.Equal(1, grandLineItem.EntrySlotCount);
            Assert.True(grandLineItem.GrandTotal);

            LegacyXlsPivotTableRecord pageItemRecord = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.PageItem);
            Assert.Equal("PivotMeta", pageItemRecord.SheetName);
            Assert.Equal("Sxpi", pageItemRecord.RecordName);
            LegacyXlsPivotPageItem pageItem = Assert.Single(pageItemRecord.PageItems);
            Assert.Equal((short)1, pageItem.FieldIndex);
            Assert.Equal((short)0x7ffd, pageItem.ItemIndex);
            Assert.Equal("AllItems", pageItem.ItemIndexName);
            Assert.Equal((short)42, pageItem.ObjectId);

            LegacyXlsPivotTableRecord item = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Item && record.RecordName == "Sxvi");
            Assert.Equal("PivotMeta", item.SheetName);
            Assert.Equal((short)0, item.ItemType);
            Assert.Equal(LegacyXlsPivotItemType.Data, item.ItemTypeKind);
            Assert.Equal("Data", item.ItemTypeName);
            Assert.False(item.ItemHidden);
            Assert.True(item.ItemHideDetail);
            Assert.False(item.ItemFormula);
            Assert.False(item.ItemMissing);
            Assert.Equal((short)1, item.ItemCacheIndex);
            Assert.Equal("CacheItem:1", item.ItemCacheIndexName);
            Assert.Equal("East", item.ItemName);

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

            LegacyXlsPivotTableRecord dataItem = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.DataItem && record.RecordName == "Sxdi");
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
            Assert.Equal("NoFieldReference", dataItem.DisplayCalculationFieldReferenceName);
            Assert.Equal((short)-1, dataItem.DisplayCalculationItemIndex);
            Assert.Equal("NoItemReference", dataItem.DisplayCalculationItemReferenceName);
            Assert.Equal((ushort)14, dataItem.NumberFormatId);
            Assert.Equal("Sales", dataItem.Name);

            LegacyXlsPivotTableRecord legacyDataItem = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.DataItem && record.RecordName == "Sxdb");
            Assert.Null(legacyDataItem.SheetName);
            Assert.Equal((short)3, legacyDataItem.DataItemFieldIndex);
            Assert.Equal((short)0, legacyDataItem.AggregationFunction);
            Assert.Equal(LegacyXlsPivotAggregationFunction.Sum, legacyDataItem.AggregationFunctionKind);
            Assert.Equal("Sum", legacyDataItem.AggregationFunctionName);
            Assert.Equal((short)0, legacyDataItem.DisplayCalculation);
            Assert.Equal(LegacyXlsPivotDisplayCalculation.Value, legacyDataItem.DisplayCalculationKind);
            Assert.Equal("Value", legacyDataItem.DisplayCalculationName);
            Assert.Equal((short)-1, legacyDataItem.DisplayCalculationFieldIndex);
            Assert.Equal("NoFieldReference", legacyDataItem.DisplayCalculationFieldReferenceName);
            Assert.Equal((short)-1, legacyDataItem.DisplayCalculationItemIndex);
            Assert.Equal("NoItemReference", legacyDataItem.DisplayCalculationItemReferenceName);
            Assert.Equal((ushort)0, legacyDataItem.NumberFormatId);
            Assert.Equal("Sum of Amount", legacyDataItem.Name);

            LegacyXlsPivotTableRecord cache = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Cache && record.RecordName == "Sxdb");
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

            LegacyXlsPivotTableRecord formulaScope = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "SxFormula" && record.HasCalculatedItemFormulaScope);
            Assert.Equal(LegacyXlsPivotTableRecordKind.Formula, formulaScope.Kind);
            Assert.Equal((ushort)0, formulaScope.CalculatedItemFormulaReserved);
            Assert.Equal((short)-1, formulaScope.CalculatedItemFormulaCacheFieldIndex);
            Assert.True(formulaScope.HasCalculatedItemFormulaScope);
            Assert.True(formulaScope.CalculatedItemFormulaAppliesToAllCacheFields);
            Assert.Equal("AllCacheFields", formulaScope.CalculatedItemFormulaScopeName);
            Assert.Equal("Scope", formulaScope.CalculatedItemFormulaPayloadKind);

            LegacyXlsPivotTableRecord calculatedFieldFormula = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "SxFormula" && record.HasCalculatedFieldFormula);
            Assert.Equal(LegacyXlsPivotTableRecordKind.Formula, calculatedFieldFormula.Kind);
            Assert.Equal(20, calculatedFieldFormula.CalculatedFieldFormulaTokenByteCount);
            Assert.False(calculatedFieldFormula.HasCalculatedItemFormulaScope);
            Assert.Equal("CalculatedFieldFormulaTokens", calculatedFieldFormula.CalculatedItemFormulaPayloadKind);

            LegacyXlsPivotTableRecord formulaTokens = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "SxFmla" && record.CalculatedItemFormulaTokenByteCount.HasValue);
            Assert.Equal(LegacyXlsPivotTableRecordKind.Formula, formulaTokens.Kind);
            Assert.Equal((ushort)7, formulaTokens.CalculatedItemFormulaTokenByteCount);
            Assert.Equal(0, formulaTokens.CalculatedItemFormulaTrailingByteCount);
            Assert.Equal("FormulaTokens", formulaTokens.CalculatedItemFormulaPayloadKind);
            Assert.False(formulaTokens.HasCalculatedItemFormulaScope);

            LegacyXlsPivotTableRecord rule = Assert.Single(workbook.PivotTableRecords, record => record.RecordName == "SxRule");
            Assert.Equal(LegacyXlsPivotTableRecordKind.Rule, rule.Kind);
            Assert.Equal((byte)0, rule.RuleFieldPosition);
            Assert.Equal((byte)0xff, rule.RuleFieldIndex);
            Assert.Equal("FilteredFields", rule.RuleFieldReferenceName);
            Assert.Equal("Row", rule.RuleAxisName);
            Assert.Equal((byte)0x02, rule.RuleType);
            Assert.Equal("DataCells", rule.RuleTypeName);
            Assert.Equal((byte)0, rule.RuleOptionFlags);
            Assert.False(rule.RulePartialArea);
            Assert.False(rule.RuleDataOnly);
            Assert.False(rule.RuleLabelOnly);
            Assert.False(rule.RuleCacheBased);
            Assert.Equal((ushort)2, rule.RuleFilterCount);

            LegacyXlsPivotTableRecord filter = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.Filter);
            Assert.Null(filter.SheetName);
            Assert.Equal("SxFilt", filter.RecordName);
            Assert.Equal(2, filter.RuleFilters.Count);
            LegacyXlsPivotRuleFilter rowFilter = filter.RuleFilters[0];
            Assert.Equal("Row", rowFilter.AxisName);
            Assert.Equal((short)0, rowFilter.FieldPosition);
            Assert.Equal((short)0, rowFilter.FieldReferenceIndex);
            Assert.Equal("FieldIndex:0", rowFilter.FieldReferenceName);
            Assert.True(rowFilter.Selected);
            Assert.Equal((ushort)0x0001, rowFilter.SubtotalFlags);
            Assert.Equal("Data", Assert.Single(rowFilter.SubtotalFunctionNames));
            Assert.Equal((ushort)1, rowFilter.ItemIndexCount);
            LegacyXlsPivotRuleFilter dataFilter = filter.RuleFilters[1];
            Assert.Equal("Data", dataFilter.AxisName);
            Assert.Equal((short)0, dataFilter.FieldPosition);
            Assert.Equal((short)-2, dataFilter.FieldReferenceIndex);
            Assert.Equal("DataField", dataFilter.FieldReferenceName);
            Assert.False(dataFilter.Selected);
            Assert.Equal((ushort)0x0002, dataFilter.SubtotalFlags);
            Assert.Equal("Default", Assert.Single(dataFilter.SubtotalFunctionNames));
            Assert.Equal((ushort)0, dataFilter.ItemIndexCount);

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
            Assert.True(view.HasSupportedPivotTableMetadata);
            Assert.True(field.HasSupportedPivotTableMetadata);
            Assert.True(dataItem.HasSupportedPivotTableMetadata);
            Assert.True(legacyDataItem.HasSupportedPivotTableMetadata);
            Assert.True(cache.HasSupportedPivotTableMetadata);
            Assert.True(filter.HasSupportedPivotTableMetadata);
            Assert.True(additional.HasSupportedPivotTableMetadata);
            Assert.Equal(3, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(3, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.PivotTable]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:Sxtbl"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxFormat"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:PivotChartBits"]);
            Assert.DoesNotContain("PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxView", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:Sxdb", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxFilt", report.UnsupportedFeaturesByDetail.Keys);
            Assert.DoesNotContain("PivotTable|XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED|PivotTable:SxAddl", report.UnsupportedFeaturesByDetail.Keys);

            string markdown = report.ToMarkdown();
            Assert.Contains("Pivot table records: 32", markdown);
            Assert.Contains("Pivot Table Records By Kind", markdown);
            Assert.Contains("Pivot Table Records By Location", markdown);
            Assert.Contains("Pivot Table Records By Kind And Location", markdown);
            Assert.Contains("CacheItem\\|(workbook)", markdown);
            Assert.Contains("Pivot Table Records By Name And Location", markdown);
            Assert.Contains("SxAddl\\|PivotMeta", markdown);
            Assert.Contains("Pivot Table Workbook States", markdown);
            Assert.Contains("View:Present\\|Cache:Present\\|CacheSource:Present", markdown);
            Assert.Contains("Pivot Table View Ranges", markdown);
            Assert.Contains("A3:C11", markdown);
            Assert.Contains("Pivot Table View Field Counts", markdown);
            Assert.Contains("Fields:5;Rows:2;Columns:0;Pages:1;Data:1", markdown);
            Assert.Contains("Pivot Table Field Axes", markdown);
            Assert.Contains("Pivot Table Field Names", markdown);
            Assert.Contains("Region", markdown);
            Assert.Contains("Pivot Table Field Index Sequences", markdown);
            Assert.Contains("FieldIndexes:0,2", markdown);
            Assert.Contains("Pivot Table Item Type Kinds", markdown);
            Assert.Contains("CacheItem:1", markdown);
            Assert.Contains("Pivot Table Formula Payload Lengths", markdown);
            Assert.Contains("SxFmla\\|Bytes:9", markdown);
            Assert.Contains("Pivot Table Formula Payload Kinds", markdown);
            Assert.Contains("FormulaTokens", markdown);
            Assert.Contains("Pivot Table Formula Token Byte Counts", markdown);
            Assert.Contains("TokenBytes:7", markdown);
            Assert.Contains("Pivot Table Formula Trailing Byte Counts", markdown);
            Assert.Contains("TrailingBytes:0", markdown);
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
            Assert.Contains("Sum of Amount", markdown);
            Assert.Contains("Pivot Table Rule Filter Entry Counts", markdown);
            Assert.Contains("Pivot Table Rule Filter States", markdown);
            Assert.Contains("Pivot Table Calculated Field Formula Token Byte Counts", markdown);
            Assert.Contains("CalculatedFieldFormulaTokens", markdown);
            Assert.Contains("Axis:Row\\|Position:0\\|Field:FieldIndex:0\\|Selected:True\\|Subtotals:0x0001\\|Indexes:1", markdown);
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
            Assert.Contains("Unsupported projection gaps: 0", markdown);
            Assert.Contains("SxVdEx", markdown);
            Assert.Contains("SxAddl", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_DecodesAxisWidthPivotTableLineItems() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5PivotTableLineItemAxisSlotWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.False(report.HasUnsupportedFeatures);
            Assert.Equal(3, workbook.PivotTableRecords.Count);
            Assert.Equal(3, report.PivotTableRecordCount);
            Assert.Equal(1, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.View]);
            Assert.Equal(2, report.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.LineItem]);
            Assert.Equal(1, report.PivotTableRecordsByName["SxView"]);
            Assert.Equal(2, report.PivotTableRecordsByName["Sxli"]);
            Assert.Equal(3, report.PivotTableRecordsByLocation["PivotSlots"]);
            Assert.Equal(1, report.PivotTableViewFieldCounts["Fields:5;Rows:1;Columns:2;Pages:1;Data:1"]);
            Assert.Equal(1, report.PivotTableViewLineCounts["Rows:4;Columns:8"]);
            Assert.Equal(1, report.PivotTableLineItemCounts["LineItems:4"]);
            Assert.Equal(1, report.PivotTableLineItemCounts["LineItems:8"]);
            Assert.Equal(9, report.PivotTableLineItemTypes["LineItemType:0"]);
            Assert.Equal(3, report.PivotTableLineItemTypes["LineItemType:13"]);
            Assert.Equal(6, report.PivotTableLineItemEntryCounts["Entries:1"]);
            Assert.Equal(6, report.PivotTableLineItemEntryCounts["Entries:2"]);
            Assert.Equal(4, report.PivotTableLineItemEntrySlotCounts["Slots:1;Entries:1"]);
            Assert.Equal(6, report.PivotTableLineItemEntrySlotCounts["Slots:2;Entries:2"]);
            Assert.Equal(2, report.PivotTableLineItemEntrySlotCounts["Slots:2;Entries:1"]);
            Assert.Equal(9, report.PivotTableLineItemEntryIndexes["EntryIndex:0"]);
            Assert.Equal(6, report.PivotTableLineItemEntryIndexes["EntryIndex:1"]);
            Assert.Equal(3, report.PivotTableLineItemEntryIndexes["EntryIndex:2"]);
            Assert.Equal(3, report.PivotTableLineItemFlagStates["Subtotal:False;Block:False;Grand:False;MultiDataName:False;MultiDataOnAxis:False"]);
            Assert.Equal(6, report.PivotTableLineItemFlagStates["Subtotal:False;Block:False;Grand:False;MultiDataName:False;MultiDataOnAxis:True"]);
            Assert.Equal(1, report.PivotTableLineItemFlagStates["Subtotal:True;Block:False;Grand:True;MultiDataName:False;MultiDataOnAxis:False"]);
            Assert.Equal(2, report.PivotTableLineItemFlagStates["Subtotal:True;Block:False;Grand:True;MultiDataName:True;MultiDataOnAxis:True"]);

            LegacyXlsPivotTableRecord view = Assert.Single(workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.View);
            Assert.Equal("PivotSlots", view.SheetName);
            Assert.Equal((ushort)1, view.ViewRowFieldCount);
            Assert.Equal((ushort)2, view.ViewColumnFieldCount);
            Assert.Equal((ushort)4, view.ViewRowLineCount);
            Assert.Equal((ushort)8, view.ViewColumnLineCount);
            Assert.True(view.HasSupportedPivotTableMetadata);

            LegacyXlsPivotTableRecord[] lineItemRecords = workbook.PivotTableRecords
                .Where(record => record.Kind == LegacyXlsPivotTableRecordKind.LineItem)
                .ToArray();
            Assert.Equal(2, lineItemRecords.Length);
            Assert.Equal(4, lineItemRecords[0].LineItems.Count);
            Assert.Equal(8, lineItemRecords[1].LineItems.Count);
            Assert.All(lineItemRecords, record => Assert.True(record.HasSupportedPivotTableMetadata));

            LegacyXlsPivotLineItem rowGrandTotal = lineItemRecords[0].LineItems[3];
            Assert.Equal(LegacyXlsPivotLineItemType.GrandTotal, rowGrandTotal.ItemTypeKind);
            Assert.Equal(new short[] { 0 }, rowGrandTotal.EntryIndexes);
            Assert.Equal(1, rowGrandTotal.EntrySlotCount);

            LegacyXlsPivotLineItem firstColumnGrandTotal = lineItemRecords[1].LineItems[6];
            Assert.Equal(LegacyXlsPivotLineItemType.GrandTotal, firstColumnGrandTotal.ItemTypeKind);
            Assert.Equal(new short[] { 0 }, firstColumnGrandTotal.EntryIndexes);
            Assert.Equal(2, firstColumnGrandTotal.EntrySlotCount);
            Assert.True(firstColumnGrandTotal.GrandTotal);
            Assert.True(firstColumnGrandTotal.MultiDataOnAxis);

            LegacyXlsPivotLineItem secondColumnGrandTotal = lineItemRecords[1].LineItems[7];
            Assert.Equal(LegacyXlsPivotLineItemType.GrandTotal, secondColumnGrandTotal.ItemTypeKind);
            Assert.Equal(new short[] { 0 }, secondColumnGrandTotal.EntryIndexes);
            Assert.Equal(2, secondColumnGrandTotal.EntrySlotCount);
            Assert.True(secondColumnGrandTotal.GrandTotal);
            Assert.True(secondColumnGrandTotal.MultiDataOnAxis);

            string markdown = report.ToMarkdown();
            Assert.Contains("Pivot Table Line Item Entry Slot Counts", markdown);
            Assert.Contains("Slots:2;Entries:1", markdown);
            Assert.Contains("Unsupported features: 0", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_CountsCalculationSettings() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateCalculationSettingsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsWorkbook workbook = result.Workbook;
            LegacyXlsImportReport report = result.ImportReport;

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
            DocumentFormat.OpenXml.Spreadsheet.CalculationProperties projectedProperties = result.Document.WorkbookRoot
                .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CalculationProperties>()!;
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CalculateModeValues.Auto, projectedProperties.CalculationMode!.Value);
            Assert.Equal(42U, projectedProperties.IterateCount!.Value);
            Assert.True(projectedProperties.FullPrecision!.Value);
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.ReferenceModeValues.A1, projectedProperties.ReferenceMode!.Value);
            Assert.Equal(0.001d, projectedProperties.IterateDelta!.Value);
            Assert.True(projectedProperties.Iterate!.Value);
            Assert.True(projectedProperties.CalculationOnSave!.Value);
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
        public void LegacyXls_ImportReport_ScansChartSheetSubstreamsAsSupportedMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ChartSheetSubstreamWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.DoesNotContain(workbook.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(workbook.Worksheets);
            Assert.Equal("Data", sheet.Name);
            Assert.Empty(workbook.UnsupportedSheets);
            LegacyXlsChartSheet chartSheet = Assert.Single(workbook.ChartSheets);
            Assert.Equal("ChartOnly", chartSheet.Name);
            Assert.Equal(0x02, chartSheet.SheetType);
            Assert.Equal(1, chartSheet.ChartTextObjectCount);
            Assert.Equal(43, chartSheet.ChartRecordCount);
            Assert.Equal(4, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Container]);
            Assert.Equal(10, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Axis]);
            Assert.Equal(3, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Series]);
            Assert.Equal(9, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Formatting]);
            Assert.Equal(3, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Layout]);
            Assert.Equal(5, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.ChartType]);
            Assert.Equal(6, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.Text]);
            Assert.Equal(3, chartSheet.ChartRecordsByKind[LegacyXlsChartRecordKind.FutureMetadata]);
            Assert.Equal(1, chartSheet.ChartRecordsByChartType["Scatter"]);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.Equal(0, report.PreservedFeatureRecordCount);
            Assert.Equal(0, report.UnsupportedProjectionGapCount);
            Assert.Equal(1, report.ChartSheetCount);
            Assert.Equal(1, report.ChartSheetsByType["0x02|ChartSheet"]);
            Assert.Equal(1, report.ChartSheetsByName["ChartOnly"]);
            Assert.Equal(1, report.ChartSheetsByVisibility["Visible"]);
            Assert.Equal(1, report.ChartSheetMetadataRecordCount);
            Assert.Equal(1, report.ChartSheetMetadataRecordsByKind["ChartTextObject"]);
            Assert.Equal(1, report.ChartSheetTextObjectCounts["TextObjects:1"]);
            Assert.Equal(1, report.ChartSheetChartRecordCounts["ChartRecords:43"]);
            Assert.Equal(1, report.ChartSheetChartRecordCountsBySheet["Sheet:ChartOnly;ChartRecords:43"]);
            Assert.Equal(4, report.ChartSheetChartRecordKinds["Container"]);
            Assert.Equal(10, report.ChartSheetChartRecordKinds["Axis"]);
            Assert.Equal(3, report.ChartSheetChartRecordKinds["Series"]);
            Assert.Equal(9, report.ChartSheetChartRecordKinds["Formatting"]);
            Assert.Equal(3, report.ChartSheetChartRecordKinds["Layout"]);
            Assert.Equal(5, report.ChartSheetChartRecordKinds["ChartType"]);
            Assert.Equal(6, report.ChartSheetChartRecordKinds["Text"]);
            Assert.Equal(3, report.ChartSheetChartRecordKinds["FutureMetadata"]);
            Assert.Equal(1, report.ChartSheetChartTypes["Scatter"]);
            Assert.Equal(10, report.ChartSheetChartRecordKindsBySheet["Sheet:ChartOnly;Kind:Axis"]);
            Assert.Equal(5, report.ChartSheetChartRecordKindsBySheet["Sheet:ChartOnly;Kind:ChartType"]);
            Assert.Equal(1, report.ChartSheetChartTypesBySheet["Sheet:ChartOnly;ChartType:Scatter"]);
            Assert.Equal(1, report.ChartSheetStates["PrintSize:Missing|TextObjects:Present|ChartRecords:Present|ChartTypes:Present"]);
            Assert.Empty(report.UnsupportedChartSheetPrintSizes);
            Assert.Empty(report.UnsupportedChartSheetPrintSizeKinds);
            Assert.Empty(report.UnsupportedChartSheetTextObjectCounts);
            Assert.Empty(report.UnsupportedChartSheetChartRecordCounts);
            Assert.Empty(report.UnsupportedChartSheetStates);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.ChartSheet, report.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.ChartSheet, report.PreservedFeatureRecordsByKind.Keys);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.Chart, report.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.Chart, report.PreservedFeatureRecordsByKind.Keys);
            Assert.Empty(report.UnsupportedProjectionGapsByKind);
            Assert.Empty(report.UnsupportedProjectionGapsByDetail);
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
            Assert.Equal(1, report.ChartAxisExtensionDateRanges["Min:10;Max:120;Cross:35"]);
            Assert.Equal(1, report.ChartAxisExtensionDateUnits["Major:2 Months;Minor:7 Days;Base:Months"]);
            Assert.Equal(1, report.ChartAxisExtensionStates["AutoMin:False;AutoMax:False;AutoMajor:True;AutoMinor:True;DateAxis:True;AutoBase:False;AutoCross:False;AutoDate:True"]);
            Assert.Equal(1, report.ChartAxisExtensionReservedStates["ReservedZero"]);
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
            Assert.Equal(1, report.FormulaTokensByContextAndOperandKind["ChartDataSource|AreaReference"]);
            Assert.Equal(1, report.FormulaTokensByName["PtgArea"]);
            Assert.Equal(1, report.ChartDataFormatTargets["Series"]);
            Assert.Equal(1, report.ChartDataFormatSeriesIndexes["SeriesIndex:2"]);
            Assert.Equal(1, report.ChartDataFormatPointIndexes["PointIndex:65535"]);
            Assert.Equal(1, report.ChartDataFormatOrders["Order:1"]);
            Assert.Equal(1, report.ChartDataFormatStates["Target:Series;PointIndex:65535;SeriesIndex:2;Order:1"]);
            Assert.Equal(1, report.ChartNumberFormatIds["NumberFormatId:14"]);
            Assert.Equal(1, report.ChartFontIndexes["FontIndex:3"]);
            Assert.Equal(1, report.ChartDataTableOptions["HorizontalBorders:True;VerticalBorders:False;Outline:True;SeriesKeys:True"]);
            Assert.Equal(1, report.ChartDataTableReservedStates["ReservedClear"]);
            Assert.Equal(1, report.ChartThreeDimensionalViewAngles["Rotation:30;Elevation:20"]);
            Assert.Equal(1, report.ChartThreeDimensionalScaleValues["FieldOfView:45;Height:120;Depth:100;Gap:150"]);
            Assert.Equal(1, report.ChartThreeDimensionalStates["Perspective:True;Clustered:True;AutoScale:False;Shape:NotPie;Walls2D:False"]);
            Assert.Equal(1, report.ChartThreeDimensionalReservedStates["ReservedZero"]);
            Assert.Equal(1, report.ChartThreeDimensionalBarShapeRisers["Ellipse"]);
            Assert.Equal(1, report.ChartThreeDimensionalBarShapeTapers["ProjectedPoint"]);
            Assert.Equal(1, report.ChartThreeDimensionalBarShapeStates["Riser:Ellipse;Taper:ProjectedPoint"]);
            Assert.Equal(1, report.ChartScatterBubbleSizeRatios["Ratio:150"]);
            Assert.Equal(1, report.ChartScatterBubbleSizeRepresentations["Width"]);
            Assert.Equal(1, report.ChartScatterBubbleSizeRatioStates["Valid"]);
            Assert.Equal(1, report.ChartScatterStates["Bubble:True;NegativeBubbles:True;Shadow:True;Size:Width"]);
            Assert.Equal(1, report.ChartFontBasisScaleBasis["PlotArea"]);
            Assert.Equal(1, report.ChartFontBasisFontIndexes["FontIndex:3"]);
            Assert.Equal(1, report.ChartFontBasisStates["Basis:640x480;HeightTwips:220;Scale:PlotArea;FontIndex:3"]);
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
            Assert.Equal(1, report.DrawingTextObjectAlignments["Horizontal:Unknown:0;Vertical:Unknown:0"]);
            Assert.Equal(1, report.DrawingTextObjectRotations["None"]);
            Assert.Equal(1, report.DrawingTextObjectTextLengths["Characters:0"]);
            Assert.Equal(1, report.DrawingTextObjectFormattingRunByteCounts["RunBytes:0"]);
            Assert.Equal(1, report.DrawingTextObjectFormulaByteCounts["FormulaBytes:2"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["TextInContinueRecords:Missing"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["FormattingRunsInContinueRecords:Missing"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["LockedText:False"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["JustifyLastLine:False"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["SecretEdit:False"]);
            Assert.Equal(1, report.DrawingRecordsByLocation["ChartOnly"]);
            LegacyXlsDrawingRecord textObjectRecord = Assert.Single(workbook.DrawingRecords, record => record.RecordName == "TxO");
            Assert.NotNull(textObjectRecord.TextObject);
            Assert.Equal((ushort)0, textObjectRecord.TextObject!.RawOptions);
            Assert.Equal("Unknown:0", textObjectRecord.TextObject.HorizontalAlignmentName);
            Assert.Equal("Unknown:0", textObjectRecord.TextObject.VerticalAlignmentName);
            Assert.Equal("None", textObjectRecord.TextObject.RotationName);
            Assert.Equal(2, textObjectRecord.TextObject.FormulaByteCount);
            Assert.DoesNotContain("XLS-BIFF-FEATURE-CHART-UNSUPPORTED|ChartOnly", report.UnsupportedFeaturesByLocation.Keys);
            Assert.DoesNotContain(report.UnsupportedFeaturesByDetail.Keys, key => key.StartsWith("Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|", StringComparison.Ordinal));
            Assert.DoesNotContain(report.PreservedFeatureRecordsByDetail.Keys, key => key.StartsWith("Chart|XLS-BIFF-FEATURE-CHART-UNSUPPORTED|", StringComparison.Ordinal));
            Assert.DoesNotContain(workbook.PreservedFeatureRecords, record => record.SheetName == "ChartOnly" && record.DetailCode.StartsWith("Chart:", StringComparison.Ordinal));
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
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly"
                && record.RecordName == "AxcExt"
                && record.Kind == LegacyXlsChartRecordKind.Axis
                && record.AxisExtension != null
                && record.AxisExtension.MinimumDate == 10
                && record.AxisExtension.MaximumDate == 120
                && record.AxisExtension.MajorInterval == 2
                && record.AxisExtension.MajorUnitName == "Months"
                && record.AxisExtension.MinorInterval == 7
                && record.AxisExtension.MinorUnitName == "Days"
                && record.AxisExtension.BaseUnitName == "Months"
                && record.AxisExtension.CrossingDate == 35
                && record.AxisExtension.Flags == 0x9c
                && !record.AxisExtension.AutoMinimum
                && !record.AxisExtension.AutoMaximum
                && record.AxisExtension.AutoMajor
                && record.AxisExtension.AutoMinor
                && record.AxisExtension.DateAxis
                && !record.AxisExtension.AutoBase
                && !record.AxisExtension.AutoCrossing
                && record.AxisExtension.AutoDateAxis
                && record.AxisExtension.HasZeroReservedByte);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Dat" && record.DataTableOptions != null && record.DataTableOptions.Flags == 0x000d && record.DataTableOptions.HasHorizontalBorders && !record.DataTableOptions.HasVerticalBorders && record.DataTableOptions.HasOutlineBorder && record.DataTableOptions.ShowSeriesKeys && record.DataTableOptions.ReservedFlags == 0 && !record.DataTableOptions.HasReservedFlags && record.DataTableOptions.ReservedState == "ReservedClear");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Tick" && record.Tick != null && record.Tick.MajorTickLocationName == "Outside" && record.Tick.MinorTickLocationName == "Inside" && record.Tick.LabelLocationName == "NextToAxis" && record.Tick.BackgroundModeName == "Transparent" && record.Tick.RgbHex == "#998877" && record.Tick.Flags == 0x402d && record.Tick.RotationModeName == "RotatedClockwise" && record.Tick.AutoColor && !record.Tick.AutoBackground && record.Tick.AutoRotation && record.Tick.ReadingOrderName == "LeftToRight" && record.Tick.ColorIndex == 0x004d && record.Tick.Rotation == 30);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Pos" && record.Position != null && record.Position.TopLeftMode == 0x0005 && record.Position.TopLeftModeName == "MDCHART" && record.Position.BottomRightMode == 0x0001 && record.Position.BottomRightModeName == "MDABS" && record.Position.SemanticTypeName == "LegendManualSize" && record.Position.X1Y1MeaningName == "ChartAreaSprcOffset" && record.Position.X2Y2MeaningName == "PointSize" && record.Position.IgnoredCoordinateStateName == "None" && record.Position.HasKnownSemanticCombination && record.Position.X1 == 15 && record.Position.Y1 == 25 && record.Position.X2 == 300 && record.Position.Y2 == 120);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "BRAI" && record.DataSource != null && record.DataSource.SourceId == 0x01 && record.DataSource.SourceIdName == "ValuesOrHorizontalValues" && record.DataSource.ReferenceType == 0x02 && record.DataSource.ReferenceTypeName == "WorksheetRange" && record.DataSource.Flags == 0x0001 && record.DataSource.UsesCustomNumberFormat && record.DataSource.NumberFormatId == 14 && record.DataSource.FormulaByteCount == 9 && record.DataSource.FormulaBytesAvailable == 9 && record.DataSource.FormulaByteCountFitsPayload && record.DataSource.FormulaTextProjected && record.DataSource.FormulaText == "$B$1:$B$4");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "PlotGrowth" && record.PlotGrowth != null && record.PlotGrowth.HorizontalIntegral == 1 && record.PlotGrowth.HorizontalFractional == 0x4000 && record.PlotGrowth.HorizontalGrowthPoints == 1.25 && record.PlotGrowth.VerticalIntegral == 2 && record.PlotGrowth.VerticalFractional == 0x8000 && record.PlotGrowth.VerticalGrowthPoints == 2.5);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "GelFrame" && record.Kind == LegacyXlsChartRecordKind.Formatting);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "BopPopCustom" && record.Kind == LegacyXlsChartRecordKind.ChartType && record.ChartTypeName == "CustomBarOfPieOrPieOfPie");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly"
                && record.RecordName == "Fbi2"
                && record.Kind == LegacyXlsChartRecordKind.Text
                && record.FontBasisOptions != null
                && record.FontBasisOptions.WidthTwipsBasis == 640
                && record.FontBasisOptions.HeightTwipsBasis == 480
                && record.FontBasisOptions.FontHeightTwips == 220
                && record.FontBasisOptions.ScaleBasis == 0x0001
                && record.FontBasisOptions.ScaleBasisName == "PlotArea"
                && record.FontBasisOptions.HasKnownScaleBasis
                && record.FontBasisOptions.FontIndex == 3);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly"
                && record.RecordName == "Chart3d"
                && record.Kind == LegacyXlsChartRecordKind.ChartType
                && record.ChartTypeName == "ThreeDimensional"
                && record.ThreeDimensionalOptions != null
                && record.ThreeDimensionalOptions.RotationDegrees == 30
                && record.ThreeDimensionalOptions.ElevationDegrees == 20
                && record.ThreeDimensionalOptions.FieldOfViewDegrees == 45
                && record.ThreeDimensionalOptions.HeightOrThicknessPercent == 120
                && record.ThreeDimensionalOptions.HeightPercent == 120
                && record.ThreeDimensionalOptions.DepthPercent == 100
                && record.ThreeDimensionalOptions.GapWidthPercent == 150
                && record.ThreeDimensionalOptions.Flags == 0x0013
                && record.ThreeDimensionalOptions.UsesPerspective
                && record.ThreeDimensionalOptions.IsClustered
                && !record.ThreeDimensionalOptions.UsesAutomaticScaling
                && record.ThreeDimensionalOptions.IsNotPieChart
                && record.ThreeDimensionalOptions.ChartGroupShapeName == "NotPie"
                && !record.ThreeDimensionalOptions.UsesTwoDimensionalWalls
                && record.ThreeDimensionalOptions.HasZeroReservedBits);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Chart3DBarShape" && record.Kind == LegacyXlsChartRecordKind.ChartType && record.ChartTypeName == "ThreeDimensionalBarShape" && record.ThreeDimensionalBarShapeOptions != null && record.ThreeDimensionalBarShapeOptions.Riser == 0x01 && record.ThreeDimensionalBarShapeOptions.RiserName == "Ellipse" && record.ThreeDimensionalBarShapeOptions.Taper == 0x02 && record.ThreeDimensionalBarShapeOptions.TaperName == "ProjectedPoint");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "Scatter" && record.Kind == LegacyXlsChartRecordKind.ChartType && record.ChartTypeName == "Scatter" && record.ScatterOptions != null && record.ScatterOptions.BubbleSizeRatio == 150 && record.ScatterOptions.BubbleSizeRepresentation == 0x0002 && record.ScatterOptions.BubbleSizeRepresentationName == "Width" && record.ScatterOptions.HasKnownBubbleSizeRepresentation && record.ScatterOptions.HasValidBubbleSizeRatio && record.ScatterOptions.Flags == 0x0007 && record.ScatterOptions.IsBubbleChart && record.ScatterOptions.ShowNegativeBubbles && record.ScatterOptions.HasShadow);
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.RecordName == "End" && record.ContainerDepthBefore == 1 && record.ContainerDepthAfter == 0 && record.ContainerTransition == "End");
            Assert.Contains(workbook.ChartRecords, record => record.SheetName == "ChartOnly" && record.ChartTypeName == "Scatter");
            Assert.DoesNotContain(workbook.Diagnostics, d => d.SheetName == "ChartOnly" && d.DetailCode != null && d.DetailCode.StartsWith("Chart:", StringComparison.Ordinal));
            Assert.DoesNotContain(workbook.Diagnostics, d => d.SheetName == "ChartOnly" && d.DetailCode == "Sheet:ChartSheet");
            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Records By Rectangle", markdown);
            Assert.Contains("Chart Sheet States", markdown);
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
            Assert.Contains("Chart AxcExt Date Ranges", markdown);
            Assert.Contains("Chart AxcExt Date Units", markdown);
            Assert.Contains("Chart AxcExt States", markdown);
            Assert.Contains("Chart AxcExt Reserved States", markdown);
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
            Assert.Contains("Chart DataFormat Point Indexes", markdown);
            Assert.Contains("Chart DataFormat Orders", markdown);
            Assert.Contains("Chart DataFormat States", markdown);
            Assert.Contains("Chart Number Format Ids", markdown);
            Assert.Contains("Chart Font Indexes", markdown);
            Assert.Contains("Chart DataTable Options", markdown);
            Assert.Contains("Chart 3D View Angles", markdown);
            Assert.Contains("Chart 3D Scale Values", markdown);
            Assert.Contains("Chart 3D States", markdown);
            Assert.Contains("Chart 3D Reserved States", markdown);
            Assert.Contains("Chart 3D Bar Shape Risers", markdown);
            Assert.Contains("Chart 3D Bar Shape Tapers", markdown);
            Assert.Contains("Chart 3D Bar Shape States", markdown);
            Assert.Contains("Chart Scatter Bubble Size Ratios", markdown);
            Assert.Contains("Chart Scatter Bubble Size Representations", markdown);
            Assert.Contains("Chart Scatter Bubble Size Ratio States", markdown);
            Assert.Contains("Chart Scatter States", markdown);
            Assert.Contains("Chart FontBasis Scale Basis", markdown);
            Assert.Contains("Basis:640x480;HeightTwips:220;Scale:PlotArea;FontIndex:3", markdown);
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
            Assert.Contains("Chart Sheet Chart Record Counts", markdown);
            Assert.Contains("Chart Sheet Chart Record Kinds", markdown);
            Assert.Contains("Chart Sheet Chart Types", markdown);
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
        public void LegacyXls_ImportReport_GroupsBopPopMetadata() {
            byte[] bopPopPayload = new byte[22];
            bopPopPayload[0] = 0x02;
            bopPopPayload[1] = 0x00;
            WriteUInt16(bopPopPayload, 2, 0x0003);
            WriteUInt16(bopPopPayload, 4, 2);
            WriteUInt16(bopPopPayload, 6, 5);
            WriteUInt16(bopPopPayload, 8, 120);
            WriteUInt16(bopPopPayload, 10, 40);
            Buffer.BlockCopy(BitConverter.GetBytes(12.5d), 0, bopPopPayload, 12, 8);
            WriteUInt16(bopPopPayload, 20, 0x0001);

            byte[] customPayload = {
                0x05, 0x00,
                0x14
            };

            var chartRecords = new List<LegacyXlsChartRecord>();
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x1061, offset: 512, bopPopPayload), "BopPopChart", chartRecords));
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x1067, offset: 538, customPayload), "BopPopChart", chartRecords));

            LegacyXlsChartRecord bopPopRecord = Assert.Single(chartRecords, record => record.RecordName == "BopPop");
            LegacyXlsChartBopPopOptions? bopPop = bopPopRecord.BopPopOptions;
            Assert.NotNull(bopPop);
            Assert.Equal(0x02, bopPop!.Subtype);
            Assert.Equal("BarOfPie", bopPop.SubtypeName);
            Assert.True(bopPop.HasKnownSubtype);
            Assert.False(bopPop.AutomaticSplit);
            Assert.Equal(0x0003, bopPop.Split);
            Assert.Equal("Custom", bopPop.SplitName);
            Assert.True(bopPop.HasKnownSplit);
            Assert.Equal(2, bopPop.SplitPosition);
            Assert.Equal(5, bopPop.SplitPercent);
            Assert.Equal(120, bopPop.SecondaryPieSizePercent);
            Assert.Equal(40, bopPop.GapPercent);
            Assert.Equal(12.5d, bopPop.SplitValue);
            Assert.True(bopPop.HasShadow);
            Assert.True(bopPop.HasZeroReservedBits);

            LegacyXlsChartRecord customRecord = Assert.Single(chartRecords, record => record.RecordName == "BopPopCustom");
            LegacyXlsChartBopPopCustomSplit? customSplit = customRecord.BopPopCustomSplit;
            Assert.NotNull(customSplit);
            Assert.Equal(5, customSplit!.BitCount);
            Assert.Equal(4, customSplit.DataPointCount);
            Assert.Equal(1, customSplit.ExpectedBitmapByteCount);
            Assert.Equal(1, customSplit.BitmapBytesAvailable);
            Assert.True(customSplit.HasCompleteBitmap);
            Assert.Equal(new[] { 0, 2 }, customSplit.SecondaryDataPointIndexes);
            Assert.False(customSplit.NoSecondaryDataPointsMarker);
            Assert.True(customSplit.HasConsistentNoSecondaryDataPointsMarker);

            var workbook = new LegacyXlsWorkbook();
            foreach (LegacyXlsChartRecord record in chartRecords) {
                workbook.MutableChartRecords.Add(record);
            }

            LegacyXlsImportReport report = workbook.CreateImportReport();
            Assert.Equal(1, report.ChartBopPopSubtypes["BarOfPie"]);
            Assert.Equal(1, report.ChartBopPopSplitTypes["Custom"]);
            Assert.Equal(1, report.ChartBopPopSplitValues["Position:2;Percent:5;Size:120;Gap:40;Value:12.5"]);
            Assert.Equal(1, report.ChartBopPopStates["Subtype:BarOfPie;Split:Custom;Auto:False;Shadow:True"]);
            Assert.Equal(1, report.ChartBopPopReservedStates["ReservedZero"]);
            Assert.Equal(1, report.ChartBopPopCustomDataPointCounts["DataPoints:4"]);
            Assert.Equal(1, report.ChartBopPopCustomSecondaryCounts["Secondary:2"]);
            Assert.Equal(1, report.ChartBopPopCustomSecondaryIndexes["Secondary:0,2"]);
            Assert.Equal(1, report.ChartBopPopCustomCompletionStates["Complete"]);
            Assert.Equal(1, report.ChartBopPopCustomStates["DataPoints:4;Secondary:2;NoSecondary:False;Consistent:True"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart BopPop Subtypes", markdown);
            Assert.Contains("Chart BopPopCustom Secondary Indexes", markdown);
            Assert.Contains("Secondary:0,2", markdown);

            static void WriteUInt16(byte[] buffer, int offset, ushort value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)(value >> 8);
            }
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartErrorBarMetadata() {
            byte[] payload = new byte[14];
            payload[0] = 0x03;
            payload[1] = 0x02;
            payload[2] = 0x01;
            payload[3] = 0x01;
            WriteDouble(payload, 4, 12.5d);
            WriteUInt16(payload, 12, 4);

            var chartRecords = new List<LegacyXlsChartRecord>();
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x105B, offset: 560, payload), "ErrorBars", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartErrorBarOptions? errorBars = record.ErrorBarOptions;
            Assert.NotNull(errorBars);
            Assert.Equal(0x03, errorBars!.Direction);
            Assert.Equal("VerticalPlus", errorBars.DirectionName);
            Assert.True(errorBars.HasKnownDirection);
            Assert.Equal(0x02, errorBars.ValueSource);
            Assert.Equal("FixedValue", errorBars.ValueSourceName);
            Assert.True(errorBars.HasKnownValueSource);
            Assert.True(errorBars.HasTeeTop);
            Assert.Equal(0x01, errorBars.Reserved);
            Assert.True(errorBars.HasExpectedReservedValue);
            Assert.Equal(12.5d, errorBars.Value);
            Assert.Equal(4, errorBars.CustomValueCount);
            Assert.True(errorBars.UsesValue);
            Assert.False(errorBars.UsesCustomValueCount);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);

            LegacyXlsImportReport report = workbook.CreateImportReport();
            Assert.Equal(1, report.ChartErrorBarDirections["VerticalPlus"]);
            Assert.Equal(1, report.ChartErrorBarValueSources["FixedValue"]);
            Assert.Equal(1, report.ChartErrorBarValues["Value:12.5;CustomCount:4"]);
            Assert.Equal(1, report.ChartErrorBarStates["Direction:VerticalPlus;Source:FixedValue;Tee:True;UsesValue:True;UsesCustomCount:False"]);
            Assert.Equal(1, report.ChartErrorBarReservedStates["ReservedExpected"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Error Bar Directions", markdown);
            Assert.Contains("Chart Error Bar Value Sources", markdown);
            Assert.Contains("Direction:VerticalPlus;Source:FixedValue", markdown);

            static void WriteUInt16(byte[] buffer, int offset, ushort value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)(value >> 8);
            }

            static void WriteDouble(byte[] buffer, int offset, double value) {
                Buffer.BlockCopy(BitConverter.GetBytes(value), 0, buffer, offset, 8);
            }
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsAreaChartGroupMetadata() {
            byte[] payload = {
                0x05, 0x00
            };

            var chartRecords = new List<LegacyXlsChartRecord>();
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x101A, offset: 584, payload), "AreaChart", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            Assert.Equal("Area", record.ChartTypeName);
            LegacyXlsChartAreaOptions? area = record.AreaOptions;
            Assert.NotNull(area);
            Assert.Equal(0x0005, area!.Flags);
            Assert.True(area.IsStacked);
            Assert.False(area.IsPercentStacked);
            Assert.True(area.HasShadow);
            Assert.True(area.HasValidPercentStackedState);
            Assert.True(area.HasZeroReservedBits);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);

            LegacyXlsImportReport report = workbook.CreateImportReport();
            Assert.Equal(1, report.ChartRecordsByChartType["Area"]);
            Assert.Equal(1, report.ChartAreaStates["Stacked:True;Percent:False;Shadow:True"]);
            Assert.Equal(1, report.ChartAreaReservedStates["ReservedZero"]);
            Assert.Equal(1, report.ChartAreaPercentStackedStates["ValidPercentState"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Area States", markdown);
            Assert.Contains("Stacked:True;Percent:False;Shadow:True", markdown);
            Assert.Contains("Chart Area Percent Stacked States", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsLegendValidityMetadata() {
            byte[] payload = {
                0x0a, 0x00, 0x00, 0x00,
                0x14, 0x00, 0x00, 0x00,
                0x2c, 0x01, 0x00, 0x00,
                0x90, 0x01, 0x00, 0x00,
                0x00, 0x01,
                0x1f, 0x00
            };

            var chartRecords = new List<LegacyXlsChartRecord>();
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x1015, offset: 596, payload), "LegendChart", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            LegacyXlsChartLegend? legend = record.Legend;
            Assert.NotNull(legend);
            Assert.True(legend!.HasExpectedSpacing);
            Assert.True(legend.HasRequiredReservedBit);
            Assert.True(legend.HasZeroReservedBits);
            Assert.True(legend.HasValidReservedBits);
            Assert.True(legend.HasValidAutoPositionState);
            Assert.True(legend.HasValidDataTableState);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);

            LegacyXlsImportReport report = workbook.CreateImportReport();
            Assert.Equal(1, report.ChartLegendLayouts["Vertical"]);
            Assert.Equal(1, report.ChartLegendSpacingStates["ExpectedSpacing"]);
            Assert.Equal(1, report.ChartLegendReservedStates["ReservedExpected"]);
            Assert.Equal(1, report.ChartLegendAutoPositionStates["AutoPositionConsistent"]);
            Assert.Equal(1, report.ChartLegendDataTableStates["DataTableConsistent"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Legend Reserved States", markdown);
            Assert.Contains("ReservedExpected", markdown);
            Assert.Contains("Chart Legend Data Table States", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsLineChartGroupMetadata() {
            byte[] payload = {
                0x07, 0x00
            };

            var chartRecords = new List<LegacyXlsChartRecord>();
            Assert.True(BiffChartMetadataReader.TryRead(new BiffRecord(0x1018, offset: 572, payload), "LineChart", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            Assert.Equal("Line", record.ChartTypeName);
            LegacyXlsChartLineOptions? line = record.LineOptions;
            Assert.NotNull(line);
            Assert.Equal(0x0007, line!.Flags);
            Assert.True(line.IsStacked);
            Assert.True(line.IsPercentStacked);
            Assert.True(line.HasShadow);
            Assert.True(line.HasValidPercentStackedState);
            Assert.True(line.HasZeroReservedBits);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);

            LegacyXlsImportReport report = workbook.CreateImportReport();
            Assert.Equal(1, report.ChartRecordsByChartType["Line"]);
            Assert.Equal(1, report.ChartLineStates["Stacked:True;Percent:True;Shadow:True"]);
            Assert.Equal(1, report.ChartLineReservedStates["ReservedZero"]);
            Assert.Equal(1, report.ChartLinePercentStackedStates["ValidPercentState"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart Line States", markdown);
            Assert.Contains("Stacked:True;Percent:True;Shadow:True", markdown);
            Assert.Contains("Chart Line Percent Stacked States", markdown);
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
            Assert.Equal(LegacyXlsChartRecordKind.Layout, record.Kind);
            Assert.True(record.HasSupportedChartMetadata);
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
        public void LegacyXls_ImportReport_GroupsChartSeriesListMetadata() {
            byte[] payload = {
                0x03, 0x00,
                0x01, 0x00,
                0x00, 0x00
            };
            var chartRecord = new BiffRecord(0x1016, offset: 292, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "SeriesList", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            Assert.Equal("SeriesList", record.RecordName);
            Assert.Equal(LegacyXlsChartRecordKind.Series, record.Kind);
            LegacyXlsChartSeriesList? seriesList = record.SeriesList;
            Assert.NotNull(seriesList);
            Assert.Equal(3, seriesList!.DeclaredSeriesCount);
            Assert.Equal(2, seriesList.DecodedSeriesCount);
            Assert.False(seriesList.HasCompleteSeriesIndexList);
            Assert.False(seriesList.HasOnlyValidSeriesIndexes);
            Assert.Equal(new ushort[] { 1, 0 }, seriesList.SeriesIndexes);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartSeriesListDeclaredCounts["Declared:3"]);
            Assert.Equal(1, report.ChartSeriesListDecodedCounts["Decoded:2"]);
            Assert.Equal(1, report.ChartSeriesListCompletenessStates["Truncated"]);
            Assert.Equal(1, report.ChartSeriesListIndexValidityStates["ContainsInvalid"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart SeriesList Declared Counts", markdown);
            Assert.Contains("Declared:3", markdown);
            Assert.Contains("ContainsInvalid", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartSeriesFormatMetadata() {
            byte[] payload = {
                0x0D, 0x00
            };
            var chartRecord = new BiffRecord(0x105D, offset: 296, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "SerFmt", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            Assert.Equal("SerFmt", record.RecordName);
            Assert.Equal(LegacyXlsChartRecordKind.Formatting, record.Kind);
            LegacyXlsChartSeriesFormat? seriesFormat = record.SeriesFormat;
            Assert.NotNull(seriesFormat);
            Assert.Equal(0x000D, seriesFormat!.Flags);
            Assert.True(seriesFormat.SmoothLine);
            Assert.False(seriesFormat.ThreeDimensionalBubbles);
            Assert.True(seriesFormat.Shadow);
            Assert.Equal(0x0008, seriesFormat.Reserved);
            Assert.False(seriesFormat.HasZeroReservedBits);
            Assert.Equal(new[] { "SmoothLine", "Shadow" }, seriesFormat.FlagNames);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartSeriesFormatFlags["SmoothLine"]);
            Assert.Equal(1, report.ChartSeriesFormatFlags["Shadow"]);
            Assert.Equal(1, report.ChartSeriesFormatStates["SmoothLine:True;ThreeDimensionalBubbles:False;Shadow:True"]);
            Assert.Equal(1, report.ChartSeriesFormatReservedValues["Reserved:0x0008"]);
            Assert.Equal(1, report.ChartSeriesFormatReservedStates["ReservedNonZero"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart SerFmt Flags", markdown);
            Assert.Contains("SmoothLine", markdown);
            Assert.Contains("ReservedNonZero", markdown);
        }

        [Fact]
        public void LegacyXls_ImportReport_GroupsChartClientColorPaletteMetadata() {
            byte[] payload = {
                0x03, 0x00,
                0x11, 0x22, 0x33, 0x00,
                0x44, 0x55, 0x66, 0x00,
                0x00, 0x00, 0x00, 0x00
            };
            var chartRecord = new BiffRecord(0x105C, offset: 300, payload);
            var chartRecords = new List<LegacyXlsChartRecord>();

            Assert.True(BiffChartMetadataReader.TryRead(chartRecord, "ClrtClient", chartRecords));

            LegacyXlsChartRecord record = Assert.Single(chartRecords);
            Assert.Equal("ClrtClient", record.RecordName);
            Assert.Equal(LegacyXlsChartRecordKind.Formatting, record.Kind);
            LegacyXlsChartClientColorPalette? palette = record.ClientColorPalette;
            Assert.NotNull(palette);
            Assert.Equal(3, palette!.DeclaredColorCount);
            Assert.Equal(3, palette.DecodedColorCount);
            Assert.True(palette.HasCompleteColorList);
            Assert.True(palette.HasExpectedColorCount);
            Assert.Equal("#112233", palette.ForegroundColor);
            Assert.Equal("#445566", palette.BackgroundColor);
            Assert.Equal("#000000", palette.NeutralColor);

            var workbook = new LegacyXlsWorkbook();
            workbook.MutableChartRecords.Add(record);
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Equal(1, report.ChartClientColorPaletteDeclaredCounts["Declared:3"]);
            Assert.Equal(1, report.ChartClientColorPaletteDecodedCounts["Decoded:3"]);
            Assert.Equal(1, report.ChartClientColorPaletteCompletenessStates["Complete"]);
            Assert.Equal(1, report.ChartClientColorPaletteExpectedCountStates["ExpectedThreeColors"]);
            Assert.Equal(1, report.ChartClientColorPaletteColors["Foreground:#112233"]);
            Assert.Equal(1, report.ChartClientColorPaletteColors["Background:#445566"]);
            Assert.Equal(1, report.ChartClientColorPaletteColors["Neutral:#000000"]);

            string markdown = report.ToMarkdown();
            Assert.Contains("Chart ClrtClient Declared Counts", markdown);
            Assert.Contains("ExpectedThreeColors", markdown);
            Assert.Contains("Neutral:#000000", markdown);
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
            Assert.Equal(LegacyXlsChartRecordKind.Extension, record.Kind);
            Assert.True(record.HasSupportedChartMetadata);
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
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = true
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
        public void LegacyXls_Load_ReportsVbaProjectStorageAsSupportedCompoundMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithVbaProjectStorage(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Single(result.Document.Sheets);
            Assert.DoesNotContain(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.VbaProject);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED");
            Assert.False(result.ImportReport.HasUnsupportedFeatures);
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
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.VbaProject, result.ImportReport.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain("XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED", result.ImportReport.UnsupportedFeaturesByCode.Keys);
            Assert.DoesNotContain("VbaProject|XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED|Compound:VbaProjectStorage", result.ImportReport.UnsupportedFeaturesByDetail.Keys);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("VbaProject", markdown);
            Assert.Contains("Compound Feature Entries By Name", markdown);
            Assert.Contains("Compound Feature Entries By Role", markdown);
            Assert.Contains("Compound Feature Entries By Object Type", markdown);
            Assert.Contains("Compound Feature Entries By Content Kind", markdown);
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
            var entryContentKinds = new Dictionary<string, LegacyXlsCompoundFeatureEntryContentKind>(StringComparer.OrdinalIgnoreCase) {
                ["_VBA_PROJECT_CUR"] = LegacyXlsCompoundFeatureEntryContentKind.Storage,
                ["_VBA_PROJECT_CUR/VBA"] = LegacyXlsCompoundFeatureEntryContentKind.Storage,
                ["_VBA_PROJECT_CUR/VBA/dir"] = LegacyXlsCompoundFeatureEntryContentKind.VbaCompressedContainer,
                ["_VBA_PROJECT_CUR/VBA/Sheet1"] = LegacyXlsCompoundFeatureEntryContentKind.VbaCompressedContainer,
                ["_VBA_PROJECT_CUR/VBA/ThisWorkbook"] = LegacyXlsCompoundFeatureEntryContentKind.VbaCompressedContainer,
                ["_VBA_PROJECT_CUR/VBA/LooseModule"] = LegacyXlsCompoundFeatureEntryContentKind.BinaryStream,
                ["_VBA_PROJECT_CUR/VBA/_VBA_PROJECT"] = LegacyXlsCompoundFeatureEntryContentKind.VbaProjectMetadataStream
            };
            var record = new LegacyXlsCompoundFeatureRecord(
                LegacyXlsCompoundFeatureRecordKind.VbaProject,
                entryRoles.Keys.ToArray(),
                entryRoles,
                entrySizes,
                entryObjectTypes,
                entryContentKinds);
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
            Assert.Equal(2, report.CompoundFeatureEntriesByContentKind["Storage"]);
            Assert.Equal(3, report.CompoundFeatureEntriesByContentKind["VbaCompressedContainer"]);
            Assert.Equal(1, report.CompoundFeatureEntriesByContentKind["BinaryStream"]);
            Assert.Equal(1, report.CompoundFeatureEntriesByContentKind["VbaProjectMetadataStream"]);
            Assert.Equal(2, report.CompoundFeatureEntriesByRoleAndContentKind["VbaModuleStream|VbaCompressedContainer"]);
            Assert.Equal(1, report.CompoundFeatureEntriesByRoleAndContentKind["VbaModuleStream|BinaryStream"]);
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
            Assert.Equal(2, report.CompoundVbaModulesByContentKind["VbaCompressedContainer"]);
            Assert.Equal(1, report.CompoundVbaModulesByContentKind["BinaryStream"]);
            Assert.Equal(1, report.CompoundVbaModulesByNameAndContentKind["Sheet1|VbaCompressedContainer"]);
            Assert.Equal(1, report.CompoundVbaModulesByNameAndContentKind["ThisWorkbook|VbaCompressedContainer"]);
            Assert.Equal(1, report.CompoundVbaModulesByNameAndContentKind["LooseModule|BinaryStream"]);
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
            Assert.Contains("Compound VBA Modules By Content Kind", markdown);
            Assert.Contains("Compound VBA Modules By Name And Content Kind", markdown);
            Assert.Contains("Compound VBA Modules By CodeName Match", markdown);
            Assert.Contains("Compound VBA Modules By CodeName Match And Name", markdown);
            Assert.Contains("Compound VBA Projects By Module Count", markdown);
            Assert.Contains("Compound VBA Projects By Module Byte Count", markdown);
            Assert.Contains("Compound VBA Projects By Structure", markdown);
            Assert.Contains("VBA Project Workbook States", markdown);
        }

        [Fact]
        public void LegacyXls_Load_ReportsOleObjectStorageAsSupportedCompoundMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithOleObjectStorage(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Single(result.Document.Sheets);
            Assert.DoesNotContain(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.OleObject);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED");
            Assert.False(result.ImportReport.HasUnsupportedFeatures);
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
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.OleObject, result.ImportReport.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain("XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED", result.ImportReport.UnsupportedFeaturesByCode.Keys);
            Assert.DoesNotContain("OleObject|XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED|Compound:OleObjectStorage", result.ImportReport.UnsupportedFeaturesByDetail.Keys);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("OleObject", markdown);
            Assert.Contains("Compound Feature Entries By Name", markdown);
            Assert.Contains("Compound Feature Entries By Role", markdown);
        }

        [Fact]
        public void LegacyXls_Load_ReportsDigitalSignatureStreamAsDiagnosed() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithDigitalSignatureStream(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Single(result.Document.Sheets);
            LegacyXlsUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DigitalSignature);
            Assert.Equal("XLS-COMPOUND-FEATURE-DIGITAL-SIGNATURE-DIAGNOSED", feature.Code);
            Assert.Contains("_signatures", feature.Description);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-COMPOUND-FEATURE-DIGITAL-SIGNATURE-DIAGNOSED");
            LegacyXlsCompoundFeatureRecord compoundRecord = Assert.Single(result.Workbook.CompoundFeatureRecords);
            Assert.Equal(LegacyXlsCompoundFeatureRecordKind.DigitalSignature, compoundRecord.Kind);
            Assert.Contains("_signatures", compoundRecord.Entries);
            Assert.Equal(LegacyXlsCompoundFeatureEntryRole.DigitalSignatureStream, compoundRecord.EntryRoles["_signatures"]);
            Assert.Equal(LegacyXlsCompoundFeatureEntryObjectType.Stream, compoundRecord.EntryObjectTypes["_signatures"]);
            Assert.Equal(LegacyXlsCompoundFeatureEntryContentKind.DigitalSignatureStream, compoundRecord.EntryContentKinds["_signatures"]);
            LegacyXlsCompoundFeatureEntryInfo entry = Assert.Single(compoundRecord.EntryDetails);
            Assert.Equal("_signatures", entry.Path);
            Assert.Equal(LegacyXlsCompoundFeatureEntryRole.DigitalSignatureStream, entry.Role);
            Assert.Equal(LegacyXlsCompoundFeatureEntryObjectType.Stream, entry.ObjectType);
            Assert.Equal(LegacyXlsCompoundFeatureEntryContentKind.DigitalSignatureStream, entry.ContentKind);
            Assert.False(entry.IsStorage);
            Assert.True(entry.IsStream);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntryCount);
            Assert.Equal(1, result.ImportReport.CompoundFeatureRecordsByKind[LegacyXlsCompoundFeatureRecordKind.DigitalSignature]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKind[LegacyXlsCompoundFeatureRecordKind.DigitalSignature]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByName["_signatures"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRole["DigitalSignatureStream"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByKindAndRole["DigitalSignature|DigitalSignatureStream"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByObjectType["Stream"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRoleAndObjectType["DigitalSignatureStream|Stream"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByContentKind["DigitalSignatureStream"]);
            Assert.Equal(1, result.ImportReport.CompoundFeatureEntriesByRoleAndContentKind["DigitalSignatureStream|DigitalSignatureStream"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DigitalSignature]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-COMPOUND-FEATURE-DIGITAL-SIGNATURE-DIAGNOSED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["DigitalSignature|XLS-COMPOUND-FEATURE-DIGITAL-SIGNATURE-DIAGNOSED|Compound:DigitalSignature"]);
            Assert.Equal(0, result.ImportReport.UnsupportedProjectionGapCount);
            Assert.Empty(result.ImportReport.UnsupportedProjectionGapsByKind);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("DigitalSignature", markdown);
            Assert.Contains("Compound Feature Entries By Content Kind", markdown);
        }
    }
}
