using OpenXmlBreak = DocumentFormat.OpenXml.Spreadsheet.Break;
using OpenXmlColumnBreaks = DocumentFormat.OpenXml.Spreadsheet.ColumnBreaks;
using OpenXmlRowBreaks = DocumentFormat.OpenXml.Spreadsheet.RowBreaks;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValue;
using OpenXmlCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;
using OpenXmlAutoFilter = DocumentFormat.OpenXml.Spreadsheet.AutoFilter;
using OpenXmlConditionalFormattingOperatorValues = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingOperatorValues;
using OpenXmlCustomFilter = DocumentFormat.OpenXml.Spreadsheet.CustomFilter;
using OpenXmlDataValidationErrorStyleValues = DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues;
using OpenXmlDataValidationOperatorValues = DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues;
using OpenXmlFilter = DocumentFormat.OpenXml.Spreadsheet.Filter;
using OpenXmlFilterColumn = DocumentFormat.OpenXml.Spreadsheet.FilterColumn;
using OpenXmlFilterOperatorValues = DocumentFormat.OpenXml.Spreadsheet.FilterOperatorValues;
using OpenXmlWorksheetPart = DocumentFormat.OpenXml.Packaging.WorksheetPart;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Compound;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Drawing.Internal;
using System.Globalization;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_SavesSamePathCreatedXlsWorkbook() {
            string xlsPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(xlsPath)) {
                    ExcelSheet sheet = document.AddWorksheet("SamePath");
                    sheet.CellValue(1, 1, "Native XLS");

                    document.Save();

                    sheet = document.Sheets.Single(candidate => candidate.Name == "SamePath");
                    sheet.CellValue(1, 2, "After reopen");
                    document.Save();
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell firstCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCell secondCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
                Assert.Equal("Native XLS", Assert.IsType<string>(firstCell.Value));
                Assert.Equal("After reopen", Assert.IsType<string>(secondCell.Value));
            } finally {
                TryDelete(xlsPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCompleteBofRecordsAndActualWorkbookStreamSize() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Biff");
                    sheet.CellValue(1, 1, "BIFF8 BOF");

                    document.Save(xlsOutputPath);
                }

                IReadOnlyList<byte[]> bofPayloads = GetBiffRecordPayloads(xlsOutputPath, 0x0809);
                Assert.Equal(2, bofPayloads.Count);
                Assert.All(bofPayloads, payload => {
                    Assert.Equal(16, payload.Length);
                    Assert.Equal((ushort)0x0600, ReadUInt16(payload, 0));
                    Assert.Equal(0x00000041U, ReadUInt32(payload, 8));
                    Assert.Equal(0x00000006U, ReadUInt32(payload, 12));
                });
                Assert.Equal((ushort)0x0005, ReadUInt16(bofPayloads[0], 2));
                Assert.Equal((ushort)0x0010, ReadUInt16(bofPayloads[1], 2));

                byte[] fileBytes = File.ReadAllBytes(xlsOutputPath);
                Assert.Equal(4096U, ReadUInt32(fileBytes, 56));
                CompoundDirectoryEntry workbookEntry = Assert.Single(
                    ReadFirstCompoundDirectorySector(fileBytes),
                    entry => string.Equals(entry.Name, "Workbook", StringComparison.OrdinalIgnoreCase));
                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.True(result.Workbook.Worksheets.Count > 0);

                byte[] workbookStream = ReadCompoundStream(fileBytes, "Workbook");
                Assert.Equal(
                    checked((ulong)GetBiffContentLength(workbookStream, expectedEndOfFileRecords: result.Workbook.Worksheets.Count + 1)),
                    workbookEntry.StreamSize);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_LoadedAutoSaveWorkbookDoesNotCopyOpenXmlPackageOverXlsOnDispose() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CopyBack");
                    sheet.CellValue(1, 1, "OpenXML source");
                    document.Save();
                }

                using (ExcelDocument document = ExcelDocument.Load(openXmlPath, new OfficeIMO.Excel.ExcelLoadOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                    ExcelSheet sheet = document.Sheets.Single();
                    sheet.CellValue(1, 1, "Native XLS target");
                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsCell cell = Assert.Single(Assert.Single(result.Workbook.Worksheets).Cells, item => item.Row == 1 && item.Column == 1);
                Assert.Equal("Native XLS target", Assert.IsType<string>(cell.Value));

                using ExcelDocument sourceDocument = ExcelDocument.Load(openXmlPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
                ExcelSheet sourceSheet = Assert.Single(sourceDocument.Sheets);
                Assert.True(sourceSheet.TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? sourceValue));
                Assert.Equal("OpenXML source", sourceValue!.Text);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeStreamSave_DisablesAutoSaveCopyBackOnDispose() {
            using var stream = new MemoryStream();

            using (ExcelDocument document = ExcelDocument.Create(stream, new OfficeIMO.Excel.ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                ExcelSheet sheet = document.AddWorksheet("StreamXls");
                sheet.CellValue(1, 1, "Native stream XLS");

                document.Save(stream, ExcelFileFormat.Xls);
            }

            AssertLegacyXlsStreamCell(stream, 1, 1, "Native stream XLS");
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksLossyImportedXlsWithoutExplicitOptIn() {
            string sourcePath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "objects.xls");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using ExcelDocument document = ExcelDocument.Load(sourcePath);
                Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
                Assert.Empty(document.LegacyXlsUnsupportedFeatures);
                Assert.NotEmpty(document.LegacyXlsChartSheets);

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(xlsOutputPath));
                Assert.Contains(nameof(ExcelSaveOptions.LossPolicy), exception.Message);
                Assert.False(File.Exists(xlsOutputPath));
            } finally {
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_PreservesImportedVbaProjectAndWritesWorkbookMarker() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] vbaPayload = { 0x01, 0x00, 0x4d, 0x41, 0x43, 0x52, 0x4f };
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithVbaProjectPayload(
                workbookStream,
                vbaPayload);

            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound));
            Assert.True(document.SourceFormat == ExcelFileFormat.Xls);
            Assert.Empty(document.LegacyXlsUnsupportedFeatures);
            LegacyXlsCompoundFeatureRecord feature = Assert.Single(document.LegacyXlsCompoundFeatures);
            Assert.Equal(LegacyXlsCompoundFeatureRecordKind.VbaProject, feature.Kind);
            ExcelSheet sheet = Assert.Single(document.Sheets);
            sheet.CellValue(1, 1, "Edited with VBA preserved");

            using var conversionOutput = new MemoryStream(new byte[] { 7, 8, 9 }, writable: true);
            NotSupportedException conversionException = Assert.Throws<NotSupportedException>(() =>
                document.Save(conversionOutput, ExcelFileFormat.Xlsx));
            Assert.Contains(nameof(ExcelSaveOptions.LossPolicy), conversionException.Message);
            Assert.Equal(new byte[] { 7, 8, 9 }, conversionOutput.ToArray());

            using var output = new MemoryStream();
            document.Save(output, ExcelFileFormat.Xls);
            byte[] savedBytes = output.ToArray();

            Assert.Equal(vbaPayload, ReadCompoundStream(savedBytes, "_VBA_PROJECT_CUR/VBA/dir"));
            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(savedBytes));
            result.EnsureNoImportErrors();
            Assert.True(result.Workbook.HasVbaProjectMarker);
            Assert.Contains(result.CompoundFeatures, item => item.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject);
            LegacyXlsCell cell = Assert.Single(Assert.Single(result.Workbook.Worksheets).Cells,
                item => item.Row == 1 && item.Column == 1);
            Assert.Equal("Edited with VBA preserved", Assert.IsType<string>(cell.Value));
        }

        [Fact]
        public void LegacyXls_NativeSave_StillBlocksImportedOleObjectsWithoutExplicitOptIn() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] olePayload = { 0x01, 0x4f, 0x4c, 0x45 };
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithOleObjectPayload(
                workbookStream,
                olePayload);

            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound));
            Assert.Contains(document.LegacyXlsCompoundFeatures,
                feature => feature.Kind == LegacyXlsCompoundFeatureRecordKind.OleObject);
            using var output = new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: true);

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(output, ExcelFileFormat.Xls));

            Assert.Contains(nameof(ExcelSaveOptions.LossPolicy), exception.Message);
            Assert.Equal(new byte[] { 1, 2, 3, 4 }, output.ToArray());

            using var allowedOutput = new MemoryStream();
            document.Save(allowedOutput, ExcelFileFormat.Xls, new ExcelSaveOptions {
                LossPolicy = ExcelConversionLossPolicy.Allow
            });
            Assert.Equal(
                olePayload,
                ReadCompoundStream(allowedOutput.ToArray(), "ObjectPool/OLEPackage/\u0001Ole10Native"));
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksLegacyDigitalSignaturesByDefault() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithDigitalSignatureStream(workbookStream);
            byte[] signatureStream = ReadCompoundStream(compound, "_signatures");

            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound));
            Assert.Contains(document.LegacyXlsUnsupportedFeatures,
                feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DigitalSignature);
            Assert.Contains(document.LegacyXlsCompoundFeatures,
                feature => feature.Kind == LegacyXlsCompoundFeatureRecordKind.DigitalSignature);
            using var blockedOutput = new MemoryStream(new byte[] { 4, 3, 2, 1 }, writable: true);

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(blockedOutput, ExcelFileFormat.Xls));

            Assert.Contains("DigitalSignature", exception.Message);
            Assert.Equal(new byte[] { 4, 3, 2, 1 }, blockedOutput.ToArray());

            using var allowedOutput = new MemoryStream();
            document.Save(allowedOutput, ExcelFileFormat.Xls, new ExcelSaveOptions {
                LossPolicy = ExcelConversionLossPolicy.Allow
            });
            Assert.Equal(
                signatureStream,
                ReadCompoundStream(allowedOutput.ToArray(), "_signatures"));
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetLayoutMetadata() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Layout");
                    sheet.CellValue(1, 1, "Merged header");
                    sheet.CellValue(2, 2, 12.5d);
                    sheet.CellValue(3, 1, "Hidden row marker");
                    sheet.MergeRange("A1:C1");
                    sheet.SetDefaultColumnWidth(11d);
                    sheet.SetDefaultRowHeight(18.5d, hidden: true);
                    sheet.SetColumnWidth(2, 12.5d);
                    sheet.SetColumnOutline(2, 2, collapsed: true);
                    sheet.SetColumnHidden(3, true);
                    sheet.SetRowHeight(2, 21d);
                    sheet.SetRowOutline(2, 1, collapsed: true);
                    sheet.SetRowHidden(3, true);
                    sheet.Freeze(2, 1);
                    sheet.SetGridlinesVisible(false);
                    sheet.SetRowColumnHeadingsVisible(false);
                    sheet.SetZeroValuesVisible(false);
                    sheet.SetRightToLeft(true);
                    sheet.SetZoomScale(125);
                    sheet.SetViewOptions(zoomScaleNormal: 90, view: ExcelWorksheetViewKind.PageBreakPreview);
                    DocumentFormat.OpenXml.Spreadsheet.SheetView sheetView = sheet.WorksheetPart.Worksheet
                        .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>()!
                        .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetView>()!;
                    sheetView.ShowFormulas = true;
                    sheetView.DefaultGridColor = false;
                    sheetView.ColorId = 22U;
                    sheetView.ShowOutlineSymbols = false;
                    sheetView.TabSelected = true;
                    sheetView.TopLeftCell = "D6";
                    sheetView.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Pane>()!.State = DocumentFormat.OpenXml.Spreadsheet.PaneStateValues.FrozenSplit;

                    document.Save(xlsOutputPath);
                }

                using ExcelDocument loaded = ExcelDocument.Load(xlsOutputPath);
                ExcelSheet loadedSheet = loaded.Sheets.Single();

                Assert.True(loaded.SourceFormat == ExcelFileFormat.Xls);
                Assert.Equal(11d, loadedSheet.DefaultColumnWidth);
                Assert.Equal(18.5d, loadedSheet.DefaultRowHeight);
                Assert.True(loadedSheet.DefaultRowsHidden);

                ExcelMergedRangeSnapshot merge = Assert.Single(loadedSheet.GetMergedRanges());
                Assert.Equal("A1:C1", merge.A1Range);
                Assert.Equal(1, merge.StartRow);
                Assert.Equal(3, merge.EndColumn);

                ExcelColumnSnapshot widthColumn = Assert.Single(loadedSheet.GetColumnDefinitions(), column => column.StartIndex == 2 && column.EndIndex == 2);
                Assert.Equal(12.5d, widthColumn.Width);
                Assert.Equal((byte?)2, widthColumn.OutlineLevel);
                Assert.True(widthColumn.Collapsed);

                ExcelColumnSnapshot hiddenColumn = Assert.Single(loadedSheet.GetColumnDefinitions(), column => column.StartIndex == 3 && column.EndIndex == 3);
                Assert.True(hiddenColumn.Hidden);

                ExcelRowSnapshot heightRow = Assert.Single(loadedSheet.GetRowDefinitions(), row => row.Index == 2);
                Assert.Equal(21d, heightRow.Height);
                Assert.True(heightRow.CustomHeight);
                Assert.Equal((byte?)1, heightRow.OutlineLevel);
                Assert.True(heightRow.Collapsed);

                ExcelRowSnapshot hiddenRow = Assert.Single(loadedSheet.GetRowDefinitions(), row => row.Index == 3);
                Assert.True(hiddenRow.Hidden);

                ExcelWorksheetViewInfo view = loadedSheet.GetViewInfo();
                Assert.True(view.HasPane);
                Assert.Equal(2, view.FrozenRowCount);
                Assert.Equal(1, view.FrozenColumnCount);
                Assert.False(view.ShowGridlines);
                Assert.False(loadedSheet.RowColumnHeadingsVisible);
                Assert.False(loadedSheet.ZeroValuesVisible);
                Assert.True(view.RightToLeft);
                Assert.Equal(125U, view.ZoomScale);
                Assert.Equal(90U, view.ZoomScaleNormal);
                Assert.Equal(125U, loadedSheet.GetZoomScale());
                Assert.Equal("pageBreakPreview", view.View);
                DocumentFormat.OpenXml.Spreadsheet.SheetView loadedSheetView = loadedSheet.WorksheetPart.Worksheet
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetView>()!;
                DocumentFormat.OpenXml.Spreadsheet.Pane loadedPane = loadedSheetView.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Pane>()!;
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.PaneStateValues.FrozenSplit, loadedPane.State!.Value);
                Assert.True(loadedSheetView.ShowFormulas!.Value);
                Assert.False(loadedSheetView.DefaultGridColor!.Value);
                Assert.Equal(22U, loadedSheetView.ColorId!.Value);
                Assert.False(loadedSheetView.ShowOutlineSymbols!.Value);
                Assert.True(loadedSheetView.TabSelected!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.SheetViewValues.PageBreakPreview, loadedSheetView.View!.Value);
                Assert.Equal(90U, loadedSheetView.ZoomScaleNormal!.Value);
                Assert.Equal("D6", loadedSheetView.TopLeftCell!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesPageLayoutView() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("PageLayout");
                    sheet.CellValue(1, 1, "Page layout view");
                    sheet.SetViewOptions(zoomScale: 115, zoomScaleNormal: 100, view: ExcelWorksheetViewKind.PageLayout);
                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.RecordType == 0x088b);
                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.True(legacySheet.PageLayoutView);
                Assert.Equal(115U, legacySheet.PageLayoutZoomScale);
                LegacyXlsSheetFutureMetadataRecord pageLayoutMetadata = Assert.Single(
                    legacySheet.FutureMetadataRecords,
                    record => record.Kind == LegacyXlsWorkbookMetadataKind.PageLayoutView);
                Assert.True(pageLayoutMetadata.HasMatchingFutureRecordHeader);
                Assert.Equal((ushort)0x0005, pageLayoutMetadata.HeaderFlags);
                Assert.Equal(4, pageLayoutMetadata.BodyByteCount);

                ExcelSheet loadedSheet = result.Document.Sheets.Single();
                ExcelWorksheetViewInfo view = loadedSheet.GetViewInfo();
                Assert.Equal("pageLayout", view.View);
                Assert.Equal(115U, view.ZoomScale);
                Assert.Equal(100U, view.ZoomScaleNormal);

                using ExcelDocument normalLoaded = ExcelDocument.Load(xlsOutputPath);
                ExcelWorksheetViewInfo normalView = normalLoaded.Sheets.Single().GetViewInfo();
                Assert.Equal("pageLayout", normalView.View);
                Assert.Equal(115U, normalView.ZoomScale);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetTabColor() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("TabColor");
                    sheet.CellValue(1, 1, "Tab color");
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                    DocumentFormat.OpenXml.Spreadsheet.SheetProperties sheetProperties =
                        worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>()
                        ?? worksheet.InsertAt(new DocumentFormat.OpenXml.Spreadsheet.SheetProperties(), 0);
                    sheetProperties.TabColor = new DocumentFormat.OpenXml.Spreadsheet.TabColor {
                        Rgb = "FF336699"
                    };

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.True(legacySheet.TabColorIndex.HasValue);
                LegacyXlsSheetFutureMetadataRecord sheetExtensionMetadata = Assert.Single(
                    legacySheet.FutureMetadataRecords,
                    record => record.Kind == LegacyXlsWorkbookMetadataKind.SheetExtension);
                Assert.True(sheetExtensionMetadata.HasMatchingFutureRecordHeader);
                Assert.Equal(8, sheetExtensionMetadata.BodyByteCount);

                DocumentFormat.OpenXml.Spreadsheet.SheetProperties projectedProperties = result.Document.Sheets.Single()
                    .WorksheetPart
                    .Worksheet
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>()!;
                Assert.Equal("FF336699", projectedProperties.TabColor!.Rgb!.Value);

                using ExcelDocument normalLoaded = ExcelDocument.Load(xlsOutputPath);
                DocumentFormat.OpenXml.Spreadsheet.SheetProperties normalProperties = normalLoaded.Sheets.Single()
                    .WorksheetPart
                    .Worksheet
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>()!;
                Assert.Equal("FF336699", normalProperties.TabColor!.Rgb!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSplitPanes() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Split");
                    sheet.CellValue(1, 1, "Split pane");

                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                    DocumentFormat.OpenXml.Spreadsheet.SheetViews sheetViews = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>() ?? worksheet.InsertAt(new DocumentFormat.OpenXml.Spreadsheet.SheetViews(), 0);
                    DocumentFormat.OpenXml.Spreadsheet.SheetView sheetView = sheetViews.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetView>() ?? sheetViews.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetView { WorkbookViewId = 0U });
                    sheetView.RemoveAllChildren<DocumentFormat.OpenXml.Spreadsheet.Pane>();
                    sheetView.PrependChild(new DocumentFormat.OpenXml.Spreadsheet.Pane {
                        State = DocumentFormat.OpenXml.Spreadsheet.PaneStateValues.Split,
                        HorizontalSplit = 1200D,
                        VerticalSplit = 900D,
                        TopLeftCell = "C5",
                        ActivePane = DocumentFormat.OpenXml.Spreadsheet.PaneValues.BottomRight
                    });
                    worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.NotNull(legacySheet.SplitPane);
                Assert.Equal((ushort)1200, legacySheet.SplitPane!.HorizontalSplit);
                Assert.Equal((ushort)900, legacySheet.SplitPane.VerticalSplit);
                Assert.Equal((ushort)4, legacySheet.SplitPane.TopRow);
                Assert.Equal((ushort)2, legacySheet.SplitPane.LeftColumn);
                Assert.Equal((byte)0, legacySheet.SplitPane.ActivePane);

                DocumentFormat.OpenXml.Spreadsheet.Pane projectedPane = result.Document.Sheets.Single().WorksheetPart.Worksheet
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetView>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Pane>()!;
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.PaneStateValues.Split, projectedPane.State!.Value);
                Assert.Equal(1200D, projectedPane.HorizontalSplit!.Value);
                Assert.Equal(900D, projectedPane.VerticalSplit!.Value);
                Assert.Equal("C5", projectedPane.TopLeftCell!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.PaneValues.BottomRight, projectedPane.ActivePane!.Value);

                using ExcelDocument normalLoaded = ExcelDocument.Load(xlsOutputPath);
                DocumentFormat.OpenXml.Spreadsheet.Pane normalPane = normalLoaded.Sheets.Single().WorksheetPart.Worksheet
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetView>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Pane>()!;
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.PaneStateValues.Split, normalPane.State!.Value);
                Assert.Equal("C5", normalPane.TopLeftCell!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookAndWorksheetCodeNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CodeNames");
                    sheet.CellValue(1, 1, "Code name metadata");

                    DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties workbookProperties =
                        document.WorkbookRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties>()
                        ?? new DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties();
                    if (workbookProperties.Parent == null) {
                        DocumentFormat.OpenXml.Spreadsheet.Sheets? sheets = document.WorkbookRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                        if (sheets != null) {
                            document.WorkbookRoot.InsertBefore(workbookProperties, sheets);
                        } else {
                            document.WorkbookRoot.Append(workbookProperties);
                        }
                    }

                    workbookProperties.CodeName = "ThisWorkbook";

                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                    DocumentFormat.OpenXml.Spreadsheet.SheetProperties sheetProperties =
                        worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>()
                        ?? new DocumentFormat.OpenXml.Spreadsheet.SheetProperties();
                    if (sheetProperties.Parent == null) {
                        worksheet.InsertAt(sheetProperties, 0);
                    }

                    sheetProperties.CodeName = "MetadataSheet";
                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.Equal("ThisWorkbook", result.Workbook.CodeName);
                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal("MetadataSheet", legacySheet.CodeName);
                Assert.Equal(1, result.ImportReport.WorkbookCodeNames["ThisWorkbook"]);
                Assert.Equal(1, result.ImportReport.WorksheetCodeNames["MetadataSheet"]);

                DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties projectedWorkbookProperties =
                    result.Document.WorkbookRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties>()!;
                Assert.Equal("ThisWorkbook", projectedWorkbookProperties.CodeName!.Value);
                DocumentFormat.OpenXml.Spreadsheet.SheetProperties projectedSheetProperties =
                    result.Document.WorkbookPartRoot.WorksheetParts.Single().Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>()!;
                Assert.Equal("MetadataSheet", projectedSheetProperties.CodeName!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetTabIds() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet summary = document.AddWorksheet("Summary");
                    ExcelSheet data = document.AddWorksheet("Data");
                    summary.CellValue(1, 1, "Tab id source");
                    data.CellValue(1, 1, 42d);

                    DocumentFormat.OpenXml.Spreadsheet.Sheet[] sheets =
                        document.WorkbookRoot.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToArray();
                    sheets[0].SheetId = 7U;
                    sheets[1].SheetId = 9U;

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.NotNull(result.Workbook.SheetTabIds);
                Assert.Equal(new ushort[] { 7, 9 }, result.Workbook.SheetTabIds!.TabIds);

                uint[] projectedSheetIds = result.Document.WorkbookRoot.Sheets!
                    .Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                    .Select(sheet => sheet.SheetId!.Value)
                    .ToArray();
                Assert.Equal(new uint[] { 7U, 9U }, projectedSheetIds);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCalculationSettings() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Calc");
                    sheet.CellValue(1, 1, "Calculation settings");
                    document.WorkbookRoot.Append(new DocumentFormat.OpenXml.Spreadsheet.CalculationProperties {
                        CalculationMode = DocumentFormat.OpenXml.Spreadsheet.CalculateModeValues.AutoNoTable,
                        IterateCount = 42U,
                        FullPrecision = false,
                        ReferenceMode = DocumentFormat.OpenXml.Spreadsheet.ReferenceModeValues.R1C1,
                        IterateDelta = 0.005d,
                        Iterate = true,
                        CalculationOnSave = true
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.Equal(7, result.Workbook.CalculationSettings.Records.Count);
                Assert.Equal(LegacyXlsCalculationMode.AutomaticExceptTables, result.Workbook.CalculationSettings.Mode);
                Assert.Equal((short)42, result.Workbook.CalculationSettings.IterationCount);
                Assert.False(result.Workbook.CalculationSettings.FullPrecision!.Value);
                Assert.False(result.Workbook.CalculationSettings.A1ReferenceMode!.Value);
                Assert.Equal(0.005d, result.Workbook.CalculationSettings.Delta!.Value);
                Assert.True(result.Workbook.CalculationSettings.IterationEnabled!.Value);
                Assert.True(result.Workbook.CalculationSettings.RecalculateBeforeSave!.Value);

                DocumentFormat.OpenXml.Spreadsheet.CalculationProperties projectedProperties = result.Document.WorkbookRoot
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CalculationProperties>()!;
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CalculateModeValues.AutoNoTable, projectedProperties.CalculationMode!.Value);
                Assert.Equal(42U, projectedProperties.IterateCount!.Value);
                Assert.False(projectedProperties.FullPrecision!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.ReferenceModeValues.R1C1, projectedProperties.ReferenceMode!.Value);
                Assert.Equal(0.005d, projectedProperties.IterateDelta!.Value);
                Assert.True(projectedProperties.Iterate!.Value);
                Assert.True(projectedProperties.CalculationOnSave!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetFullCalculationOnLoad() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("SheetCalc");
                    sheet.CellValue(1, 1, "Sheet recalc");
                    sheet.WorksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.SheetCalculationProperties {
                        FullCalculationOnLoad = true
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                LegacyXlsWorksheet worksheet = result.Workbook.Worksheets[0];
                Assert.True(worksheet.FullCalculationOnLoad.GetValueOrDefault());
                Assert.Contains(worksheet.MetadataRecords, record => record.Kind == LegacyXlsWorksheetMetadataKind.Uncalced);

                DocumentFormat.OpenXml.Spreadsheet.SheetCalculationProperties projectedProperties = result.Document.Sheets[0]
                    .WorksheetPart
                    .Worksheet!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetCalculationProperties>()!;
                Assert.NotNull(projectedProperties);
                Assert.True(projectedProperties.FullCalculationOnLoad!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetPhoneticSettings() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Phonetic");
                    sheet.CellValue(1, 1, "Worksheet phonetic defaults");
                    sheet.WorksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.PhoneticProperties {
                        FontId = 0U,
                        Type = DocumentFormat.OpenXml.Spreadsheet.PhoneticValues.Hiragana,
                        Alignment = DocumentFormat.OpenXml.Spreadsheet.PhoneticAlignmentValues.Center
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = result.Workbook.Worksheets[0];
                LegacyXlsPhoneticSettings settings = Assert.IsType<LegacyXlsPhoneticSettings>(worksheet.PhoneticSettings);
                Assert.Equal((ushort)0, settings.FontId);
                Assert.Equal(LegacyXlsPhoneticType.Hiragana, settings.Type);
                Assert.Equal(LegacyXlsPhoneticAlignment.Center, settings.Alignment);
                Assert.Empty(settings.Ranges);
                Assert.Contains(worksheet.MetadataRecords, record => record.Kind == LegacyXlsWorksheetMetadataKind.PhoneticSettings);

                DocumentFormat.OpenXml.Spreadsheet.PhoneticProperties projectedProperties = result.Document.Sheets[0]
                    .WorksheetPart
                    .Worksheet!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PhoneticProperties>()!;
                Assert.NotNull(projectedProperties);
                Assert.Equal(0U, projectedProperties.FontId!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.PhoneticValues.Hiragana, projectedProperties.Type!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.PhoneticAlignmentValues.Center, projectedProperties.Alignment!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookOptions() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Options");
                    sheet.CellValue(1, 1, "Workbook options");

                    DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties workbookProperties =
                        document.WorkbookRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties>()
                        ?? new DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties();
                    if (workbookProperties.Parent == null) {
                        DocumentFormat.OpenXml.Spreadsheet.Sheets? sheets = document.WorkbookRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                        if (sheets != null) {
                            document.WorkbookRoot.InsertBefore(workbookProperties, sheets);
                        } else {
                            document.WorkbookRoot.Append(workbookProperties);
                        }
                    }

                    workbookProperties.BackupFile = true;
                    workbookProperties.SaveExternalLinkValues = false;
                    workbookProperties.ShowObjects = DocumentFormat.OpenXml.Spreadsheet.ObjectDisplayValues.None;
                    workbookProperties.ShowBorderUnselectedTables = false;
                    workbookProperties.RefreshAllConnections = true;

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.True(result.Workbook.SaveBackup!.Value);
                Assert.True(result.Workbook.DoNotSaveExternalLinkValues!.Value);
                Assert.Equal((ushort)2, result.Workbook.HiddenObjectsMode!.Value);
                Assert.True(result.Workbook.HideBordersForInactiveTables!.Value);
                Assert.True(result.Workbook.HasRefreshAllMarker);

                DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties projectedProperties = result.Document.WorkbookRoot
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties>()!;
                Assert.True(projectedProperties.BackupFile!.Value);
                Assert.False(projectedProperties.SaveExternalLinkValues!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.ObjectDisplayValues.None, projectedProperties.ShowObjects!.Value);
                Assert.False(projectedProperties.ShowBorderUnselectedTables!.Value);
                Assert.True(projectedProperties.RefreshAllConnections!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesPageSetupAndPrintMetadata() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Print");
                    sheet.CellValue(1, 1, "Printable sheet");
                    sheet.CellValue(20, 5, 42d);
                    sheet.SetMargins(0.25d, 0.35d, 0.45d, 0.55d, 0.2d, 0.25d);
                    sheet.SetOrientation(ExcelPageOrientation.Landscape);
                    sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 0U, scale: 90U, pageOrder: ExcelPageOrder.OverThenDown);
                    sheet.SetPrintOptions(printGridLines: true, printHeadings: true, horizontalCentered: true, verticalCentered: false);
                    sheet.SetHeaderFooter("Left header", "Page &P", "Right header", "Left footer", "Generated", "Page &P of &N");
                    sheet.AddManualRowPageBreak(10);
                    sheet.AddManualColumnPageBreak(3);

                    document.Save(xlsOutputPath);
                }

                using ExcelDocument loaded = ExcelDocument.Load(xlsOutputPath);
                ExcelSheet loadedSheet = loaded.Sheets.Single();

                Assert.True(loaded.SourceFormat == ExcelFileFormat.Xls);

                ExcelSheetPageSetup pageSetup = loadedSheet.GetPageSetup();
                Assert.Equal(ExcelPageOrientation.Landscape, pageSetup.Orientation);
                Assert.Equal(1U, pageSetup.FitToWidth);
                Assert.Equal(0U, pageSetup.FitToHeight);
                Assert.Equal(90U, pageSetup.Scale);
                Assert.Equal(ExcelPageOrder.OverThenDown, pageSetup.PageOrder);
                Assert.NotNull(pageSetup.Margins);
                Assert.Equal(0.25d, pageSetup.Margins!.Left, 6);
                Assert.Equal(0.35d, pageSetup.Margins.Right, 6);
                Assert.Equal(0.45d, pageSetup.Margins.Top, 6);
                Assert.Equal(0.55d, pageSetup.Margins.Bottom, 6);
                Assert.Equal(0.2d, pageSetup.Margins.Header, 6);
                Assert.Equal(0.25d, pageSetup.Margins.Footer, 6);

                ExcelSheetPrintOptions printOptions = loadedSheet.GetPrintOptions();
                Assert.True(printOptions.PrintGridLines);
                Assert.True(printOptions.PrintHeadings);
                Assert.True(printOptions.HorizontalCentered);
                Assert.False(printOptions.VerticalCentered);

                ExcelSheet.HeaderFooterSnapshot headerFooter = loadedSheet.GetHeaderFooter();
                Assert.Equal("Left header", headerFooter.HeaderLeft);
                Assert.Equal("Page &P", headerFooter.HeaderCenter);
                Assert.Equal("Right header", headerFooter.HeaderRight);
                Assert.Equal("Left footer", headerFooter.FooterLeft);
                Assert.Equal("Generated", headerFooter.FooterCenter);
                Assert.Equal("Page &P of &N", headerFooter.FooterRight);

                Assert.Equal(new[] { 10 }, loadedSheet.GetManualRowPageBreaks());
                Assert.Equal(new[] { 3 }, loadedSheet.GetManualColumnPageBreaks());
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksManualRowPageBreaksOutsideBiff8LimitsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("manual row page breaks outside BIFF8 worksheet limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Out-of-range row break");
                sheet.WorksheetPart.Worksheet.AppendChild(new OpenXmlRowBreaks(
                    new OpenXmlBreak {
                        Id = 65536U,
                        Min = 0U,
                        Max = 255U,
                        ManualPageBreak = true
                    }) {
                    Count = 1U,
                    ManualBreakCount = 1U
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksManualColumnPageBreaksOutsideBiff8LimitsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("manual column page breaks outside BIFF8 worksheet limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Out-of-range column break");
                sheet.WorksheetPart.Worksheet.AppendChild(new OpenXmlColumnBreaks(
                    new OpenXmlBreak {
                        Id = 257U,
                        Min = 0U,
                        Max = 65535U,
                        ManualPageBreak = true
                    }) {
                    Count = 1U,
                    ManualBreakCount = 1U
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksTerminalManualColumnPageBreakBeforeWriting() {
            AssertNativeXlsSaveNotSupported("manual column page breaks outside BIFF8 worksheet limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Terminal column break");
                sheet.WorksheetPart.Worksheet.AppendChild(new OpenXmlColumnBreaks(
                    new OpenXmlBreak {
                        Id = 256U,
                        Min = 0U,
                        Max = 65535U,
                        ManualPageBreak = true
                    }) {
                    Count = 1U,
                    ManualBreakCount = 1U
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetOptionsAndGridSetMetadata() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Options");
                    sheet.CellValue(1, 1, "Worksheet options");

                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
                    DocumentFormat.OpenXml.Spreadsheet.SheetProperties? sheetProperties = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>();
                    if (sheetProperties == null) {
                        sheetProperties = new DocumentFormat.OpenXml.Spreadsheet.SheetProperties();
                        worksheet.InsertAt(sheetProperties, 0);
                    }

                    sheetProperties.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.OutlineProperties {
                        ApplyStyles = true,
                        SummaryBelow = false,
                        SummaryRight = true
                    });
                    sheetProperties.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties {
                        FitToPage = true
                    });
                    worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.PrintOptions {
                        GridLines = true,
                        GridLinesSet = false
                    });
                    worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using (LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath)) {
                    result.EnsureNoImportErrors();
                    Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                    LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                    Assert.True(legacySheet.ApplyOutlineStyles);
                    Assert.False(legacySheet.SummaryRowsBelow);
                    Assert.True(legacySheet.SummaryColumnsRightWhenLeftToRight);
                    Assert.True(legacySheet.PageSetup!.FitToPage);
                    Assert.True(legacySheet.PageSetup.PrintGridLines);
                    Assert.False(legacySheet.GridSet);
                    Assert.Contains(legacySheet.MetadataRecords, record => record.Kind == LegacyXlsWorksheetMetadataKind.SheetOptions);
                    Assert.Contains(legacySheet.MetadataRecords, record => record.Kind == LegacyXlsWorksheetMetadataKind.GridSet);
                }

                using ExcelDocument loaded = ExcelDocument.Load(xlsOutputPath);
                ExcelSheet loadedSheet = loaded.Sheets.Single();
                DocumentFormat.OpenXml.Spreadsheet.Worksheet loadedWorksheet = loadedSheet.WorksheetPart.Worksheet!;
                DocumentFormat.OpenXml.Spreadsheet.OutlineProperties loadedOutline = loadedWorksheet
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.OutlineProperties>()!;
                DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties loadedPageSetupProperties = loadedWorksheet
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties>()!;
                DocumentFormat.OpenXml.Spreadsheet.PrintOptions loadedPrintOptions = loadedWorksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PrintOptions>()!;

                Assert.True(loaded.SourceFormat == ExcelFileFormat.Xls);
                Assert.True(loadedOutline.ApplyStyles!.Value);
                Assert.False(loadedOutline.SummaryBelow!.Value);
                Assert.True(loadedOutline.SummaryRight!.Value);
                Assert.True(loadedPageSetupProperties.FitToPage!.Value);
                Assert.True(loadedPrintOptions.GridLines!.Value);
                Assert.False(loadedPrintOptions.GridLinesSet!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesHeaderFooterVariants() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Variants");
                    sheet.CellValue(1, 1, "Header/footer variants");
                    sheet.SetHeaderFooter(
                        headerCenter: "Odd header",
                        footerRight: "Odd footer",
                        alignWithMargins: false,
                        scaleWithDoc: false);
                    sheet.SetFirstPageHeaderFooter(headerLeft: "First left", footerCenter: "First footer");
                    sheet.SetEvenPageHeaderFooter(headerRight: "Even right", footerLeft: "Even footer");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.NotNull(legacySheet.PageSetup);
                Assert.Equal("&COdd header", legacySheet.PageSetup!.HeaderText);
                Assert.Equal("&ROdd footer", legacySheet.PageSetup.FooterText);
                Assert.Equal("&LFirst left", legacySheet.PageSetup.FirstHeaderText);
                Assert.Equal("&CFirst footer", legacySheet.PageSetup.FirstFooterText);
                Assert.Equal("&REven right", legacySheet.PageSetup.EvenHeaderText);
                Assert.Equal("&LEven footer", legacySheet.PageSetup.EvenFooterText);
                Assert.True(legacySheet.PageSetup.DifferentFirstHeaderFooter);
                Assert.True(legacySheet.PageSetup.DifferentOddEvenHeaderFooter);
                Assert.False(legacySheet.PageSetup.ScaleHeaderFooterWithDocument);
                Assert.False(legacySheet.PageSetup.AlignHeaderFooterWithMargins);

                using ExcelDocument loaded = ExcelDocument.Load(xlsOutputPath);
                ExcelSheet.HeaderFooterSnapshot headerFooter = loaded.Sheets.Single().GetHeaderFooter();
                Assert.Equal("Odd header", headerFooter.HeaderCenter);
                Assert.Equal("Odd footer", headerFooter.FooterRight);
                Assert.Equal("First left", headerFooter.FirstHeaderLeft);
                Assert.Equal("First footer", headerFooter.FirstFooterCenter);
                Assert.Equal("Even right", headerFooter.EvenHeaderRight);
                Assert.Equal("Even footer", headerFooter.EvenFooterLeft);
                Assert.True(headerFooter.DifferentFirstPage);
                Assert.True(headerFooter.DifferentOddEven);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetVisibility() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet visible = document.AddWorksheet("Visible");
                    visible.CellValue(1, 1, "Visible");

                    ExcelSheet hidden = document.AddWorksheet("Hidden");
                    hidden.CellValue(1, 1, "Hidden");
                    hidden.SetHidden(true);

                    ExcelSheet veryHidden = document.AddWorksheet("VeryHidden");
                    veryHidden.CellValue(1, 1, "VeryHidden");
                    veryHidden.SetVeryHidden(true);

                    document.Save(xlsOutputPath);
                }

                using ExcelDocument loaded = ExcelDocument.Load(xlsOutputPath);

                ExcelSheet visibleSheet = loaded.Sheets.Single(sheet => sheet.Name == "Visible");
                ExcelSheet hiddenSheet = loaded.Sheets.Single(sheet => sheet.Name == "Hidden");
                ExcelSheet veryHiddenSheet = loaded.Sheets.Single(sheet => sheet.Name == "VeryHidden");

                Assert.False(visibleSheet.Hidden);
                Assert.False(visibleSheet.VeryHidden);
                Assert.True(hiddenSheet.Hidden);
                Assert.False(hiddenSheet.VeryHidden);
                Assert.True(veryHiddenSheet.Hidden);
                Assert.True(veryHiddenSheet.VeryHidden);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesActiveWorkbookWindow() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    document.AddWorksheet("First").CellValue(1, 1, "First");
                    document.AddWorksheet("Second").CellValue(1, 1, "Second");
                    document.AddWorksheet("Third").CellValue(1, 1, "Third");
                    document.SetActiveWorksheet("Third");
                    DocumentFormat.OpenXml.Spreadsheet.WorkbookView workbookView = document.WorkbookRoot
                        .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.BookViews>()!
                        .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookView>()!;
                    workbookView.XWindow = 123;
                    workbookView.YWindow = 234;
                    workbookView.WindowWidth = 6000U;
                    workbookView.WindowHeight = 5000U;
                    workbookView.Visibility = DocumentFormat.OpenXml.Spreadsheet.VisibilityValues.Hidden;
                    workbookView.Minimized = true;
                    workbookView.ShowHorizontalScroll = false;
                    workbookView.ShowVerticalScroll = false;
                    workbookView.ShowSheetTabs = false;
                    workbookView.TabRatio = 725U;

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbookWindow window = Assert.Single(result.Workbook.Windows);
                Assert.Equal((short)123, window.HorizontalPositionTwips);
                Assert.Equal((short)234, window.VerticalPositionTwips);
                Assert.Equal((short)6000, window.WidthTwips);
                Assert.Equal((short)5000, window.HeightTwips);
                Assert.True(window.Hidden);
                Assert.True(window.Minimized);
                Assert.False(window.VeryHidden);
                Assert.Equal((ushort)2, window.ActiveSheetIndex);
                Assert.Equal((ushort)2, window.FirstVisibleSheetTabIndex);
                Assert.Equal((ushort)1, window.SelectedSheetTabCount);
                Assert.False(window.HorizontalScrollBarVisible);
                Assert.False(window.VerticalScrollBarVisible);
                Assert.False(window.SheetTabsVisible);
                Assert.Equal((ushort)725, window.SheetTabRatio);

                DocumentFormat.OpenXml.Spreadsheet.WorkbookView projectedView = result.Document
                    .WorkbookRoot
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.BookViews>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookView>()!;
                Assert.Equal(123, projectedView.XWindow!.Value);
                Assert.Equal(234, projectedView.YWindow!.Value);
                Assert.Equal(6000U, projectedView.WindowWidth!.Value);
                Assert.Equal(5000U, projectedView.WindowHeight!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.VisibilityValues.Hidden, projectedView.Visibility!.Value);
                Assert.True(projectedView.Minimized!.Value);
                Assert.Equal(2U, projectedView.ActiveTab!.Value);
                Assert.Equal(2U, projectedView.FirstSheet!.Value);
                Assert.False(projectedView.ShowHorizontalScroll!.Value);
                Assert.False(projectedView.ShowVerticalScroll!.Value);
                Assert.False(projectedView.ShowSheetTabs!.Value);
                Assert.Equal(725U, projectedView.TabRatio!.Value);

                bool[] selectedTabs = result.Document.Sheets
                    .Select(sheet => sheet.WorksheetPart.Worksheet
                        .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>()!
                        .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetView>()!
                        .TabSelected!.Value)
                    .ToArray();
                Assert.Equal(new[] { false, false, true }, selectedTabs);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesNumberFormats() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Formats");
                    sheet.CellValue(1, 1, 44562d);
                    sheet.CellAt(1, 1).SetNumberFormat("yyyy-mm-dd");
                    sheet.CellValue(2, 1, new DateTime(2026, 1, 2));

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);

                LegacyXlsCell customDateCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat customDateFormat = workbook.CellFormats[customDateCell.StyleIndex];
                Assert.Equal("yyyy-mm-dd", customDateFormat.NumberFormatCode);
                Assert.False(customDateFormat.IsBuiltInNumberFormat);
                Assert.True(customDateFormat.IsDateLike);
                Assert.Contains(workbook.NumberFormats, format => format.FormatId == customDateFormat.NumberFormatId && format.FormatCode == "yyyy-mm-dd");

                LegacyXlsCell builtInDateCell = Assert.Single(worksheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
                LegacyXlsCellFormat builtInDateFormat = workbook.CellFormats[builtInDateCell.StyleIndex];
                Assert.Equal((ushort)14, builtInDateFormat.NumberFormatId);
                Assert.True(builtInDateFormat.IsBuiltInNumberFormat);
                Assert.True(builtInDateFormat.IsDateLike);
                Assert.DoesNotContain(workbook.NumberFormats, format => format.FormatId == 14);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesBasicCellFontStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Fonts");
                    sheet.CellValue(1, 1, 12.345d);
                    sheet.CellAt(1, 1)
                        .SetNumberFormat("0.00")
                        .SetBold()
                        .SetItalic()
                        .SetUnderline()
                        .SetFontName("Consolas")
                        .SetFontColor("#123456");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell styledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[styledCell.StyleIndex];
                Assert.True(cellFormat.ApplyFont);
                Assert.Equal("0.00", cellFormat.NumberFormatCode);

                LegacyXlsFont font = GetLegacyFont(workbook, cellFormat.FontIndex);
                Assert.Equal("Consolas", font.Name);
                Assert.True(font.Bold);
                Assert.True(font.Italic);
                Assert.True(font.Underline);
                Assert.True(workbook.TryResolveColor(font.ColorIndex, out string? fontColor));
                Assert.Equal("FF123456", fontColor);

                ExcelCellStyleSnapshot projectedStyle = result.Document.Sheets[0].GetCellStyle(1, 1);
                Assert.True(projectedStyle.Bold);
                Assert.True(projectedStyle.Italic);
                Assert.True(projectedStyle.Underline);
                Assert.Equal("Consolas", projectedStyle.FontName);
                Assert.Equal("FF123456", projectedStyle.FontColorArgb);
                Assert.Equal("0.00", projectedStyle.NumberFormatCode);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCellFontFamilyAndCharset() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("FontBytes");
                    sheet.CellValue(1, 1, "Font bytes");
                    sheet.CellAt(1, 1).SetFontName("Arial");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var openXmlFont = new DocumentFormat.OpenXml.Spreadsheet.Font(
                        new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = "Arial" },
                        new DocumentFormat.OpenXml.Spreadsheet.FontFamilyNumbering { Val = 2 });
                    var charset = new DocumentFormat.OpenXml.OpenXmlUnknownElement(string.Empty, "charset", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    charset.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("val", string.Empty, "238"));
                    openXmlFont.AppendChild(charset);
                    stylesheet.Fonts!.Append(openXmlFont);
                    stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                    var format = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    format.FontId = stylesheet.Fonts.Count!.Value - 1U;
                    format.ApplyFont = true;
                    stylesheet.CellFormats.Append(format);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell styledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[styledCell.StyleIndex];
                LegacyXlsFont font = GetLegacyFont(workbook, cellFormat.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal((byte)2, font.Family);
                Assert.Equal((byte)238, font.CharacterSet);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCellFontVerticalTextAlignment() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("FontEscapement");
                    sheet.CellValue(1, 1, "Superscript");
                    sheet.CellAt(1, 1).SetFontName("Arial");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var openXmlFont = new DocumentFormat.OpenXml.Spreadsheet.Font(
                        new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = "Arial" },
                        new DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment { Val = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentRunValues.Superscript });
                    stylesheet.Fonts!.Append(openXmlFont);
                    stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                    var format = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    format.FontId = stylesheet.Fonts.Count!.Value - 1U;
                    format.ApplyFont = true;
                    stylesheet.CellFormats.Append(format);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell styledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[styledCell.StyleIndex];
                LegacyXlsFont font = GetLegacyFont(workbook, cellFormat.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal(LegacyXlsFontEscapement.Superscript, font.Escapement);

                DocumentFormat.OpenXml.Spreadsheet.Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                DocumentFormat.OpenXml.Spreadsheet.CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                DocumentFormat.OpenXml.Spreadsheet.Font projectedFont = projectedStylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)projectedFormat.FontId!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentRunValues.Superscript, projectedFont.VerticalTextAlignment!.Val!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCellFontOptionFlags() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("FontOptionFlags");
                    sheet.CellValue(1, 1, "Option flags");
                    sheet.CellAt(1, 1).SetFontName("Arial");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var openXmlFont = new DocumentFormat.OpenXml.Spreadsheet.Font(
                        new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = "Arial" },
                        new DocumentFormat.OpenXml.Spreadsheet.Outline(),
                        new DocumentFormat.OpenXml.Spreadsheet.Shadow(),
                        new DocumentFormat.OpenXml.Spreadsheet.Condense(),
                        new DocumentFormat.OpenXml.Spreadsheet.Extend());
                    stylesheet.Fonts!.Append(openXmlFont);
                    stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                    var format = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    format.FontId = stylesheet.Fonts.Count!.Value - 1U;
                    format.ApplyFont = true;
                    stylesheet.CellFormats.Append(format);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell styledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[styledCell.StyleIndex];
                LegacyXlsFont font = GetLegacyFont(workbook, cellFormat.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.True(font.Outline);
                Assert.True(font.Shadow);
                Assert.True(font.Condense);
                Assert.True(font.Extend);

                DocumentFormat.OpenXml.Spreadsheet.Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                DocumentFormat.OpenXml.Spreadsheet.CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                DocumentFormat.OpenXml.Spreadsheet.Font projectedFont = projectedStylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)projectedFormat.FontId!.Value);
                Assert.NotNull(projectedFont.Outline);
                Assert.NotNull(projectedFont.Shadow);
                Assert.NotNull(projectedFont.Condense);
                Assert.NotNull(projectedFont.Extend);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCellFontUnderlineStyle() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("UnderlineStyle");
                    sheet.CellValue(1, 1, "Double underline");
                    sheet.CellAt(1, 1).SetFontName("Arial");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var openXmlFont = new DocumentFormat.OpenXml.Spreadsheet.Font(
                        new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = "Arial" },
                        new DocumentFormat.OpenXml.Spreadsheet.Underline { Val = DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.Double });
                    stylesheet.Fonts!.Append(openXmlFont);
                    stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                    var format = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    format.FontId = stylesheet.Fonts.Count!.Value - 1U;
                    format.ApplyFont = true;
                    stylesheet.CellFormats.Append(format);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell styledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[styledCell.StyleIndex];
                LegacyXlsFont font = GetLegacyFont(workbook, cellFormat.FontIndex);
                Assert.Equal("Arial", font.Name);
                Assert.Equal((byte)0x02, font.UnderlineStyle);

                DocumentFormat.OpenXml.Spreadsheet.Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                DocumentFormat.OpenXml.Spreadsheet.CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                DocumentFormat.OpenXml.Spreadsheet.Font projectedFont = projectedStylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)projectedFormat.FontId!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.Double, projectedFont.Underline!.Val!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesBasicCellFillStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Fills");
                    sheet.CellValue(1, 1, "Filled");
                    sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell filledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[filledCell.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? foregroundColor));
                Assert.Equal("FFABCDEF", foregroundColor);
                Assert.True(workbook.TryResolveColor(cellFormat.FillBackgroundColorIndex, out string? backgroundColor));
                Assert.Equal("FFABCDEF", backgroundColor);

                ExcelCellStyleSnapshot projectedStyle = result.Document.Sheets[0].GetCellStyle(1, 1);
                Assert.Equal("FFABCDEF", projectedStyle.FillColorArgb);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesPatternedCellFillStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("PatternFills");
                    sheet.CellValue(1, 1, "Patterned");
                    sheet.CellAt(1, 1).SetFillColor("#EEEEEE");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var patternedFill = new DocumentFormat.OpenXml.Spreadsheet.Fill(new DocumentFormat.OpenXml.Spreadsheet.PatternFill {
                        PatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.DarkGrid,
                        ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = "FF123456" },
                        BackgroundColor = new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor { Rgb = "FFABCDEF" }
                    });
                    stylesheet.Fills!.Append(patternedFill);
                    stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
                    var patternedFormat = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    patternedFormat.FillId = stylesheet.Fills.Count!.Value - 1;
                    patternedFormat.ApplyFill = true;
                    stylesheet.CellFormats.Append(patternedFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell filledCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[filledCell.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)9, cellFormat.FillPattern);
                Assert.True(workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? foregroundColor));
                Assert.Equal("FF123456", foregroundColor);
                Assert.True(workbook.TryResolveColor(cellFormat.FillBackgroundColorIndex, out string? backgroundColor));
                Assert.Equal("FFABCDEF", backgroundColor);

                DocumentFormat.OpenXml.Spreadsheet.Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                DocumentFormat.OpenXml.Spreadsheet.Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                DocumentFormat.OpenXml.Spreadsheet.Fill projectedFill = projectedStylesheet.Fills!.Elements<DocumentFormat.OpenXml.Spreadsheet.Fill>().ElementAt((int)projectedFormat.FillId!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.PatternValues.DarkGrid, projectedFill.PatternFill!.PatternType!.Value);
                Assert.Equal("FF123456", projectedFill.PatternFill.ForegroundColor!.Rgb!.Value);
                Assert.Equal("FFABCDEF", projectedFill.PatternFill.BackgroundColor!.Rgb!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesBasicCellAlignmentStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Alignment");
                    sheet.CellValue(1, 1, "Wrapped text");
                    sheet.CellAlign(1, 1, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    sheet.CellVerticalAlign(1, 1, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Bottom);
                    sheet.CellWrapText(1, 1);
                    sheet.CellValue(2, 1, "Top aligned text");
                    sheet.CellVerticalAlign(2, 1, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell alignedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[alignedCell.StyleIndex];
                Assert.True(cellFormat.ApplyAlignment);
                Assert.Equal((byte)2, cellFormat.HorizontalAlignment);
                Assert.Equal((byte)2, cellFormat.VerticalAlignment);
                Assert.True(cellFormat.WrapText);
                LegacyXlsCell topAlignedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
                LegacyXlsCellFormat topAlignedCellFormat = workbook.CellFormats[topAlignedCell.StyleIndex];
                Assert.True(topAlignedCellFormat.ApplyAlignment);
                Assert.Equal((byte)0, topAlignedCellFormat.VerticalAlignment);

                ExcelCellStyleSnapshot projectedStyle = result.Document.Sheets[0].GetCellStyle(1, 1);
                Assert.Equal("center", projectedStyle.HorizontalAlignment);
                Assert.Equal("bottom", projectedStyle.VerticalAlignment);
                Assert.True(projectedStyle.WrapText);
                ExcelCellStyleSnapshot topProjectedStyle = result.Document.Sheets[0].GetCellStyle(2, 1);
                Assert.Equal("top", topProjectedStyle.VerticalAlignment);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExtendedCellAlignmentStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Alignment");
                    sheet.CellValue(1, 1, "Extended alignment");
                    sheet.CellAlign(1, 1, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right);

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var alignedFormat = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    alignedFormat.Alignment = new DocumentFormat.OpenXml.Spreadsheet.Alignment {
                        Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right,
                        Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
                        TextRotation = 45U,
                        Indent = 3U,
                        ShrinkToFit = true,
                        ReadingOrder = 2U
                    };
                    alignedFormat.ApplyAlignment = true;
                    stylesheet.CellFormats.Append(alignedFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell alignedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[alignedCell.StyleIndex];
                Assert.True(cellFormat.ApplyAlignment);
                Assert.Equal((byte)3, cellFormat.HorizontalAlignment);
                Assert.Equal((byte)1, cellFormat.VerticalAlignment);
                Assert.Equal((byte)45, cellFormat.TextRotation);
                Assert.Equal((byte)3, cellFormat.Indent);
                Assert.True(cellFormat.ShrinkToFit);
                Assert.Equal((byte)2, cellFormat.ReadingOrder);

                DocumentFormat.OpenXml.Spreadsheet.Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                DocumentFormat.OpenXml.Spreadsheet.Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                Assert.True(projectedFormat.ApplyAlignment!.Value);
                Assert.NotNull(projectedFormat.Alignment);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right, projectedFormat.Alignment!.Horizontal!.Value);
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center, projectedFormat.Alignment.Vertical!.Value);
                Assert.Equal(45U, projectedFormat.Alignment.TextRotation!.Value);
                Assert.Equal(3U, projectedFormat.Alignment.Indent!.Value);
                Assert.True(projectedFormat.Alignment.ShrinkToFit!.Value);
                Assert.Equal(2U, projectedFormat.Alignment.ReadingOrder!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesBasicCellBorderStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Borders");
                    sheet.CellValue(1, 1, "Bordered");
                    sheet.CellAt(1, 1)
                        .SetFillColor("#ABCDEF")
                        .SetBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, "#654321")
                        .SetDiagonalBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Dotted, "#123456", diagonalUp: true, diagonalDown: true);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell borderedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[borderedCell.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? fillColor));
                Assert.Equal("FFABCDEF", fillColor);
                Assert.True(cellFormat.ApplyBorder);
                Assert.NotNull(cellFormat.Border);
                LegacyXlsBorder border = cellFormat.Border!;
                Assert.Equal((byte)1, border.LeftStyle);
                Assert.Equal((byte)1, border.RightStyle);
                Assert.Equal((byte)1, border.TopStyle);
                Assert.Equal((byte)1, border.BottomStyle);
                Assert.Equal((byte)4, border.DiagonalStyle);
                Assert.True(border.DiagonalUp);
                Assert.True(border.DiagonalDown);
                Assert.True(workbook.TryResolveColor(border.LeftColorIndex, out string? leftColor));
                Assert.Equal("FF654321", leftColor);
                Assert.True(workbook.TryResolveColor(border.RightColorIndex, out string? rightColor));
                Assert.Equal("FF654321", rightColor);
                Assert.True(workbook.TryResolveColor(border.TopColorIndex, out string? topColor));
                Assert.Equal("FF654321", topColor);
                Assert.True(workbook.TryResolveColor(border.BottomColorIndex, out string? bottomColor));
                Assert.Equal("FF654321", bottomColor);
                Assert.True(workbook.TryResolveColor(border.DiagonalColorIndex, out string? diagonalColor));
                Assert.Equal("FF123456", diagonalColor);

                ExcelCellStyleSnapshot projectedStyle = result.Document.Sheets[0].GetCellStyle(1, 1);
                Assert.Equal("FFABCDEF", projectedStyle.FillColorArgb);
                Assert.NotNull(projectedStyle.Border);
                Assert.Equal("thin", projectedStyle.Border!.Left!.Style);
                Assert.Equal("FF654321", projectedStyle.Border.Left.ColorArgb);
                Assert.Equal("thin", projectedStyle.Border.Right!.Style);
                Assert.Equal("FF654321", projectedStyle.Border.Right.ColorArgb);
                Assert.Equal("thin", projectedStyle.Border.Top!.Style);
                Assert.Equal("FF654321", projectedStyle.Border.Top.ColorArgb);
                Assert.Equal("thin", projectedStyle.Border.Bottom!.Style);
                Assert.Equal("FF654321", projectedStyle.Border.Bottom.ColorArgb);
                Assert.Equal("dotted", projectedStyle.Border.Diagonal!.Style);
                Assert.Equal("FF123456", projectedStyle.Border.Diagonal.ColorArgb);
                Assert.True(projectedStyle.Border.DiagonalUp);
                Assert.True(projectedStyle.Border.DiagonalDown);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesAutomaticBorderColorWhenOpenXmlBorderOmitsColor() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("AutoBorder");
                    sheet.CellValue(1, 1, "Automatic border");
                    sheet.CellAt(1, 1).SetBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell borderedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsBorder? border = workbook.CellFormats[borderedCell.StyleIndex].Border;
                Assert.NotNull(border);
                Assert.Equal((byte)1, border.LeftStyle);
                Assert.Equal((byte)1, border.RightStyle);
                Assert.Equal((byte)1, border.TopStyle);
                Assert.Equal((byte)1, border.BottomStyle);
                Assert.Equal((ushort)0x0040, border.LeftColorIndex);
                Assert.Equal((ushort)0x0040, border.RightColorIndex);
                Assert.Equal((ushort)0x0040, border.TopColorIndex);
                Assert.Equal((ushort)0x0040, border.BottomColorIndex);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesCellProtectionStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Protection");
                    sheet.CellValue(1, 1, "Unlocked hidden formula");
                    sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var protectedFormat = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    protectedFormat.Protection = new DocumentFormat.OpenXml.Spreadsheet.Protection {
                        Locked = false,
                        Hidden = true
                    };
                    protectedFormat.ApplyProtection = true;
                    stylesheet.CellFormats.Append(protectedFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell protectedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[protectedCell.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(cellFormat.ApplyProtection);
                Assert.False(cellFormat.Locked);
                Assert.True(cellFormat.FormulaHidden);

                DocumentFormat.OpenXml.Spreadsheet.Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                DocumentFormat.OpenXml.Spreadsheet.Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                Assert.True(projectedFormat.ApplyProtection!.Value);
                Assert.NotNull(projectedFormat.Protection);
                Assert.False(projectedFormat.Protection!.Locked!.Value);
                Assert.True(projectedFormat.Protection.Hidden!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesQuotePrefixStyles() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("QuotePrefix");
                    sheet.CellValue(1, 1, "00123");
                    sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var quotePrefixFormat = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    quotePrefixFormat.QuotePrefix = true;
                    stylesheet.CellFormats.Append(quotePrefixFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell quotePrefixedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[quotePrefixedCell.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(cellFormat.QuotePrefix);
                Assert.False(cellFormat.ApplyProtection);

                DocumentFormat.OpenXml.Spreadsheet.Cell projectedCell = result.Document.Sheets[0].WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                DocumentFormat.OpenXml.Spreadsheet.Stylesheet projectedStylesheet = result.Document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.CellFormat projectedFormat = projectedStylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)projectedCell.StyleIndex!.Value);
                Assert.True(projectedFormat.QuotePrefix!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksGradientFillBeforeWriting() {
            AssertNativeXlsSaveNotSupported("gradient fills", (document, sheet) => {
                sheet.CellValue(1, 1, "Gradient");
                sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                var gradientFormat = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                var gradientFill = new DocumentFormat.OpenXml.Spreadsheet.Fill(
                    new DocumentFormat.OpenXml.Spreadsheet.GradientFill(
                        new DocumentFormat.OpenXml.Spreadsheet.GradientStop(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFFF0000" }) { Position = 0D },
                        new DocumentFormat.OpenXml.Spreadsheet.GradientStop(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF00FF00" }) { Position = 1D }) {
                        Degree = 45D
                    });

                stylesheet.Fills!.Append(gradientFill);
                stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
                gradientFormat.FillId = stylesheet.Fills.Count!.Value - 1;
                gradientFormat.ApplyFill = true;
                stylesheet.CellFormats.Append(gradientFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesThemeOrTintFillColors() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ThemeColor");
                    sheet.CellValue(1, 1, "Theme color");
                    sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                    DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                    var themedFormat = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                    var themedFill = new DocumentFormat.OpenXml.Spreadsheet.Fill(new DocumentFormat.OpenXml.Spreadsheet.PatternFill {
                        PatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid,
                        ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor {
                            Theme = 1U,
                            Tint = 0.4D
                        },
                        BackgroundColor = new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor {
                            Indexed = 64U
                        }
                    });

                    stylesheet.Fills!.Append(themedFill);
                    stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
                    themedFormat.FillId = stylesheet.Fills.Count!.Value - 1;
                    themedFormat.ApplyFill = true;
                    stylesheet.CellFormats.Append(themedFormat);
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                    cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                    stylesheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                LegacyXlsCell themedCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCellFormat cellFormat = workbook.CellFormats[themedCell.StyleIndex];
                Assert.True(cellFormat.ApplyFill);
                Assert.Equal((byte)1, cellFormat.FillPattern);
                Assert.True(workbook.TryResolveColor(cellFormat.FillForegroundColorIndex, out string? foregroundColor));
                Assert.Equal("FF666666", foregroundColor);

                ExcelCellStyleSnapshot projectedStyle = result.Document.Sheets[0].GetCellStyle(1, 1);
                Assert.Equal("FF666666", projectedStyle.FillColorArgb);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksStyleExtensionPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell-format style extension payloads", (document, sheet) => {
                sheet.CellValue(1, 1, "Style extension");
                sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                DocumentFormat.OpenXml.Spreadsheet.Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                DocumentFormat.OpenXml.Spreadsheet.Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                DocumentFormat.OpenXml.Spreadsheet.CellFormat baseFormat = stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt((int)baseStyleIndex);
                var extensionFormat = (DocumentFormat.OpenXml.Spreadsheet.CellFormat)baseFormat.CloneNode(true);
                extensionFormat.Append(new DocumentFormat.OpenXml.Spreadsheet.ExtensionList(new DocumentFormat.OpenXml.Spreadsheet.Extension {
                    Uri = "{00000000-0000-0000-0000-000000000001}"
                }));

                stylesheet.CellFormats.Append(extensionFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSupportedNumericFormulas() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            double dateSerial = new DateTime(2026, 1, 2).ToOADate();
            string dateSerialText = dateSerial.ToString(CultureInfo.InvariantCulture);

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Formulas");
                    sheet.CellValue(1, 1, 2d);
                    sheet.CellValue(2, 1, 3d);
                    sheet.CellValue(1, 3, 2d);
                    sheet.CellValue(1, 4, 20d);
                    sheet.CellValue(2, 3, 3d);
                    sheet.CellValue(2, 4, 30d);
                    sheet.CellValue(4, 3, 2d);
                    sheet.CellValue(4, 4, 3d);
                    sheet.CellValue(5, 3, 200d);
                    sheet.CellValue(5, 4, 300d);
                    sheet.CellFormula(3, 1, "SUM(A1:A2)");
                    sheet.CellFormula(4, 1, "A1+A2");
                    sheet.CellFormula(5, 1, "$A$1+2");
                    sheet.CellValue(6, 1, 7d);
                    sheet.CellFormula(6, 1, "SUM(A1,A2,2)");
                    sheet.CellValue(7, 1, 2.5d);
                    sheet.CellFormula(7, 1, "AVERAGE(A1,A2)");
                    sheet.CellValue(8, 1, 2d);
                    sheet.CellFormula(8, 1, "MIN(A1,A2)");
                    sheet.CellValue(9, 1, 3d);
                    sheet.CellFormula(9, 1, "MAX(A1,A2)");
                    sheet.CellValue(10, 1, 3d);
                    sheet.CellFormula(10, 1, "COUNT(A1,A2,2)");
                    sheet.CellValue(11, 1, 3d);
                    sheet.CellFormula(11, 1, "ABS(-3)");
                    sheet.CellValue(12, 1, 2d);
                    sheet.CellFormula(12, 1, "INT(2.9)");
                    sheet.CellValue(13, 1, 3d);
                    sheet.CellFormula(13, 1, "ROUND(2.6,0)");
                    sheet.CellValue(14, 1, 4d);
                    sheet.CellFormula(14, 1, "SQRT(16)");
                    sheet.CellValue(15, 1, Math.PI);
                    sheet.CellFormula(15, 1, "PI()");
                    sheet.CellValue(16, 1, 0.02d);
                    sheet.CellFormula(16, 1, "A1%");
                    sheet.CellValue(17, 1, 2.03d);
                    sheet.CellFormula(17, 1, "A1+A2%");
                    sheet.CellValue(18, 1, -2d);
                    sheet.CellFormula(18, 1, "-A1");
                    sheet.CellValue(19, 1, -1d);
                    sheet.CellFormula(19, 1, "A1+-A2");
                    sheet.CellValue(20, 1, 1d);
                    sheet.CellFormula(20, 1, "SUM(-A1,+A2)");
                    sheet.CellValue(21, 1, 5d);
                    sheet.CellFormula(21, 1, "(A1+A2)");
                    sheet.CellValue(22, 1, 10d);
                    sheet.CellFormula(22, 1, "A1*(A2+A1)");
                    sheet.CellValue(23, 1, -5d);
                    sheet.CellFormula(23, 1, "-(A1+A2)");
                    sheet.CellValue(24, 1, 0.05d);
                    sheet.CellFormula(24, 1, "(A1+A2)%");
                    sheet.CellValue(25, 1, 0d);
                    sheet.CellFormula(25, 1, "SIN(0)");
                    sheet.CellValue(26, 1, 1d);
                    sheet.CellFormula(26, 1, "COS(0)");
                    sheet.CellValue(27, 1, 0d);
                    sheet.CellFormula(27, 1, "TAN(0)");
                    sheet.CellValue(28, 1, Math.E);
                    sheet.CellFormula(28, 1, "EXP(1)");
                    sheet.CellValue(29, 1, Math.Log(2d));
                    sheet.CellFormula(29, 1, "LN(2)");
                    sheet.CellValue(30, 1, 2d);
                    sheet.CellFormula(30, 1, "LOG10(100)");
                    sheet.CellValue(31, 1, -1d);
                    sheet.CellFormula(31, 1, "SIGN(-2)");
                    sheet.CellValue(32, 1, 1d);
                    sheet.CellFormula(32, 1, "MOD(7,3)");
                    sheet.CellValue(33, 1, 3d);
                    sheet.CellFormula(33, 1, "ROUNDUP(2.1,0)");
                    sheet.CellValue(34, 1, 2d);
                    sheet.CellFormula(34, 1, "ROUNDDOWN(2.9,0)");
                    sheet.CellValue(35, 1, 8d);
                    sheet.CellFormula(35, 1, "POWER(2,3)");
                    sheet.CellValue(36, 1, 3d);
                    sheet.CellFormula(36, 1, "COUNTA(A1,A2,\"x\")");
                    sheet.CellValue(37, 1, 24d);
                    sheet.CellFormula(37, 1, "PRODUCT(A1,A2,4)");
                    sheet.CellValue(38, 1, 3d);
                    sheet.CellFormula(38, 1, "MEDIAN(A1,A2,4)");
                    sheet.CellValue(39, 1, 13d);
                    sheet.CellFormula(39, 1, "SUMPRODUCT(A1:A2,A1:A2)");
                    sheet.CellValue(40, 1, dateSerial);
                    sheet.CellFormula(40, 1, "DATE(2026,1,2)");
                    sheet.CellValue(41, 1, TimeSpan.FromHours(12).TotalDays);
                    sheet.CellFormula(41, 1, "TIME(12,0,0)");
                    sheet.CellValue(42, 1, 2026d);
                    sheet.CellFormula(42, 1, $"YEAR({dateSerialText})");
                    sheet.CellValue(43, 1, 1d);
                    sheet.CellFormula(43, 1, $"MONTH({dateSerialText})");
                    sheet.CellValue(44, 1, 2d);
                    sheet.CellFormula(44, 1, $"DAY({dateSerialText})");
                    sheet.CellValue(45, 1, 12d);
                    sheet.CellFormula(45, 1, "HOUR(0.5)");
                    sheet.CellValue(46, 1, 0d);
                    sheet.CellFormula(46, 1, "MINUTE(0.5)");
                    sheet.CellValue(47, 1, 0d);
                    sheet.CellFormula(47, 1, "SECOND(0.5)");
                    sheet.CellValue(48, 1, 1d);
                    sheet.CellFormula(48, 1, "ROW(A1)");
                    sheet.CellValue(49, 1, 2d);
                    sheet.CellFormula(49, 1, "COLUMN(B1)");
                    sheet.CellValue(50, 1, Math.Atan(1d));
                    sheet.CellFormula(50, 1, "ATAN(1)");
                    sheet.CellValue(51, 1, 2d);
                    sheet.CellFormula(51, 1, "ROWS(A1:A2)");
                    sheet.CellValue(52, 1, 2d);
                    sheet.CellFormula(52, 1, "COLUMNS(A1:B1)");
                    sheet.CellValue(53, 1, Math.Sqrt(0.5d));
                    sheet.CellFormula(53, 1, "STDEV(A1:A2)");
                    sheet.CellValue(54, 1, 3d);
                    sheet.CellFormula(54, 1, "LARGE(A1:A2,1)");
                    sheet.CellValue(55, 1, 2d);
                    sheet.CellFormula(55, 1, "COUNTBLANK(B1:B2)");
                    sheet.CellValue(56, 1, 1d);
                    sheet.CellFormula(56, 1, "COUNTIF(A1:A2,\">2\")");
                    sheet.CellValue(57, 1, 1d);
                    sheet.CellFormula(57, 1, "RSQ(A1:A2,A1:A2)");
                    sheet.CellValue(58, 1, dateSerial);
                    sheet.CellFormula(58, 1, "DATEVALUE(\"1/2/2026\")");
                    sheet.CellValue(59, 1, 0.5d);
                    sheet.CellFormula(59, 1, "VAR(A1:A2)");
                    sheet.CellValue(60, 1, 2d / 1.1d + 3d / (1.1d * 1.1d));
                    sheet.CellFormula(60, 1, "NPV(0.1,A1:A2)");
                    sheet.CellValue(61, 1, 30d);
                    sheet.CellFormula(61, 1, "INDEX(C1:D2,2,2)");
                    sheet.CellValue(62, 1, 2d);
                    sheet.CellFormula(62, 1, "MATCH(3,A1:A2,0)");
                    sheet.CellValue(63, 1, 30d);
                    sheet.CellFormula(63, 1, "VLOOKUP(3,C1:D2,2,FALSE)");
                    sheet.CellValue(64, 1, 300d);
                    sheet.CellFormula(64, 1, "HLOOKUP(3,C4:D5,2,FALSE)");
                    sheet.CellValue(65, 1, 5d);
                    sheet.CellFormula(65, 1, "SUBTOTAL(9,A1:A2)");
                    sheet.CellValue(66, 1, 3d);
                    sheet.CellFormula(66, 1, "SUMIF(A1:A2,\">2\")");
                    sheet.CellValue(67, 1, 20d);
                    sheet.CellFormula(67, 1, "IF(A1>2,10,20)");
                    sheet.CellValue(68, 1, 5d);
                    sheet.CellFormula(68, 1, "IF(A2>=3,A1+A2,0)");
                    sheet.CellValue(69, 1, 5d);
                    sheet.CellFormula(69, 1, "ROUND(A1+A2,0)");
                    sheet.CellValue(70, 1, 9d);
                    sheet.CellFormula(70, 1, "SUM(A1+A2,4)");
                    sheet.CellValue(71, 1, 20d);
                    sheet.CellFormula(71, 1, "CHOOSE(A1,10,20,30)");
                    sheet.CellValue(72, 1, 5d);
                    sheet.CellFormula(72, 1, "CHOOSE(1,A1+A2,10)");
                    sheet.CellValue(73, 1, 3d);
                    sheet.CellFormula(73, 1, "OFFSET(A1,1,0)");
                    sheet.CellValue(74, 1, 5d);
                    sheet.CellFormula(74, 1, "SUM(OFFSET(A1,0,0,2,1))");
                    sheet.CellValue(75, 1, 8d);
                    sheet.CellFormula(75, 1, "A1^A2");
                    sheet.CellValue(76, 1, 25d);
                    sheet.CellFormula(76, 1, "(A1+A2)^2");
                    sheet.CellValue(77, 1, 10d);
                    sheet.CellFormula(77, 1, "IF(A1>2,10,)");
                    sheet.CellValue(78, 1, 2d);
                    sheet.CellFormula(78, 1, "COUNT(A1,,A2)");
                    document.Calculation.EvaluateFormulasBeforeSave = true;

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);

                LegacyXlsCell sumFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 3 && cell.Column == 1);
                Assert.True(sumFormula.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Number, sumFormula.Kind);
                Assert.Equal(5d, Assert.IsType<double>(sumFormula.Value));
                Assert.Equal("SUM(A1:A2)", sumFormula.FormulaText);

                LegacyXlsCell additionFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 4 && cell.Column == 1);
                Assert.True(additionFormula.IsFormula);
                Assert.Equal(5d, Assert.IsType<double>(additionFormula.Value));
                Assert.Equal("A1+A2", additionFormula.FormulaText);

                LegacyXlsCell absoluteFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 5 && cell.Column == 1);
                Assert.True(absoluteFormula.IsFormula);
                Assert.Equal(4d, Assert.IsType<double>(absoluteFormula.Value));
                Assert.Equal("$A$1+2", absoluteFormula.FormulaText);

                LegacyXlsCell multiArgumentSum = Assert.Single(worksheet.Cells, cell => cell.Row == 6 && cell.Column == 1);
                Assert.True(multiArgumentSum.IsFormula);
                Assert.Equal(7d, Assert.IsType<double>(multiArgumentSum.Value));
                Assert.Equal("SUM(A1,A2,2)", multiArgumentSum.FormulaText);

                LegacyXlsCell averageFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 7 && cell.Column == 1);
                Assert.True(averageFormula.IsFormula);
                Assert.Equal(2.5d, Assert.IsType<double>(averageFormula.Value));
                Assert.Equal("AVERAGE(A1,A2)", averageFormula.FormulaText);

                LegacyXlsCell minFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 8 && cell.Column == 1);
                Assert.True(minFormula.IsFormula);
                Assert.Equal(2d, Assert.IsType<double>(minFormula.Value));
                Assert.Equal("MIN(A1,A2)", minFormula.FormulaText);

                LegacyXlsCell maxFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 9 && cell.Column == 1);
                Assert.True(maxFormula.IsFormula);
                Assert.Equal(3d, Assert.IsType<double>(maxFormula.Value));
                Assert.Equal("MAX(A1,A2)", maxFormula.FormulaText);

                LegacyXlsCell countFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 10 && cell.Column == 1);
                Assert.True(countFormula.IsFormula);
                Assert.Equal(3d, Assert.IsType<double>(countFormula.Value));
                Assert.Equal("COUNT(A1,A2,2)", countFormula.FormulaText);

                LegacyXlsCell absFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 11 && cell.Column == 1);
                Assert.True(absFormula.IsFormula);
                Assert.Equal(3d, Assert.IsType<double>(absFormula.Value));
                Assert.Equal("ABS(-3)", absFormula.FormulaText);

                LegacyXlsCell intFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 12 && cell.Column == 1);
                Assert.True(intFormula.IsFormula);
                Assert.Equal(2d, Assert.IsType<double>(intFormula.Value));
                Assert.Equal("INT(2.9)", intFormula.FormulaText);

                LegacyXlsCell roundFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 13 && cell.Column == 1);
                Assert.True(roundFormula.IsFormula);
                Assert.Equal(3d, Assert.IsType<double>(roundFormula.Value));
                Assert.Equal("ROUND(2.6,0)", roundFormula.FormulaText);

                LegacyXlsCell sqrtFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 14 && cell.Column == 1);
                Assert.True(sqrtFormula.IsFormula);
                Assert.Equal(4d, Assert.IsType<double>(sqrtFormula.Value));
                Assert.Equal("SQRT(16)", sqrtFormula.FormulaText);

                LegacyXlsCell piFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 15 && cell.Column == 1);
                Assert.True(piFormula.IsFormula);
                Assert.Equal(Math.PI, Assert.IsType<double>(piFormula.Value), precision: 12);
                Assert.Equal("PI()", piFormula.FormulaText);

                LegacyXlsCell percentFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 16 && cell.Column == 1);
                Assert.True(percentFormula.IsFormula);
                Assert.Equal(0.02d, Assert.IsType<double>(percentFormula.Value), precision: 12);
                Assert.Equal("A1%", percentFormula.FormulaText);

                LegacyXlsCell binaryPercentFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 17 && cell.Column == 1);
                Assert.True(binaryPercentFormula.IsFormula);
                Assert.Equal(2.03d, Assert.IsType<double>(binaryPercentFormula.Value), precision: 12);
                Assert.Equal("A1+A2%", binaryPercentFormula.FormulaText);

                LegacyXlsCell unaryMinusFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 18 && cell.Column == 1);
                Assert.True(unaryMinusFormula.IsFormula);
                Assert.Equal(-2d, Assert.IsType<double>(unaryMinusFormula.Value));
                Assert.Equal("-A1", unaryMinusFormula.FormulaText);

                LegacyXlsCell binaryUnaryFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 19 && cell.Column == 1);
                Assert.True(binaryUnaryFormula.IsFormula);
                Assert.Equal(-1d, Assert.IsType<double>(binaryUnaryFormula.Value));
                Assert.Equal("A1+-A2", binaryUnaryFormula.FormulaText);

                LegacyXlsCell aggregateUnaryFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 20 && cell.Column == 1);
                Assert.True(aggregateUnaryFormula.IsFormula);
                Assert.Equal(1d, Assert.IsType<double>(aggregateUnaryFormula.Value));
                Assert.Equal("SUM(-A1,+A2)", aggregateUnaryFormula.FormulaText);

                LegacyXlsCell parenthesizedFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 21 && cell.Column == 1);
                Assert.True(parenthesizedFormula.IsFormula);
                Assert.Equal(5d, Assert.IsType<double>(parenthesizedFormula.Value));
                Assert.Equal("(A1+A2)", parenthesizedFormula.FormulaText);

                LegacyXlsCell parenthesizedTermFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 22 && cell.Column == 1);
                Assert.True(parenthesizedTermFormula.IsFormula);
                Assert.Equal(10d, Assert.IsType<double>(parenthesizedTermFormula.Value));
                Assert.Equal("A1*(A2+A1)", parenthesizedTermFormula.FormulaText);

                LegacyXlsCell unaryParenthesizedFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 23 && cell.Column == 1);
                Assert.True(unaryParenthesizedFormula.IsFormula);
                Assert.Equal(-5d, Assert.IsType<double>(unaryParenthesizedFormula.Value));
                Assert.Equal("-(A1+A2)", unaryParenthesizedFormula.FormulaText);

                LegacyXlsCell parenthesizedPercentFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 24 && cell.Column == 1);
                Assert.True(parenthesizedPercentFormula.IsFormula);
                Assert.Equal(0.05d, Assert.IsType<double>(parenthesizedPercentFormula.Value), precision: 12);
                Assert.Equal("(A1+A2)%", parenthesizedPercentFormula.FormulaText);

                AssertNumericFormula(worksheet, 25, 0d, "SIN(0)");
                AssertNumericFormula(worksheet, 26, 1d, "COS(0)");
                AssertNumericFormula(worksheet, 27, 0d, "TAN(0)");
                AssertNumericFormula(worksheet, 28, Math.E, "EXP(1)", precision: 12);
                AssertNumericFormula(worksheet, 29, Math.Log(2d), "LN(2)", precision: 12);
                AssertNumericFormula(worksheet, 30, 2d, "LOG10(100)");
                AssertNumericFormula(worksheet, 31, -1d, "SIGN(-2)");
                AssertNumericFormula(worksheet, 32, 1d, "MOD(7,3)");
                AssertNumericFormula(worksheet, 33, 3d, "ROUNDUP(2.1,0)");
                AssertNumericFormula(worksheet, 34, 2d, "ROUNDDOWN(2.9,0)");
                AssertNumericFormula(worksheet, 35, 8d, "POWER(2,3)");
                AssertNumericFormula(worksheet, 36, 3d, "COUNTA(A1,A2,\"x\")");
                AssertNumericFormula(worksheet, 37, 24d, "PRODUCT(A1,A2,4)");
                AssertNumericFormula(worksheet, 38, 3d, "MEDIAN(A1,A2,4)");
                AssertNumericFormula(worksheet, 39, 13d, "SUMPRODUCT(A1:A2,A1:A2)");
                AssertNumericFormula(worksheet, 40, dateSerial, "DATE(2026,1,2)");
                AssertNumericFormula(worksheet, 41, TimeSpan.FromHours(12).TotalDays, "TIME(12,0,0)", precision: 12);
                AssertNumericFormula(worksheet, 42, 2026d, $"YEAR({dateSerialText})");
                AssertNumericFormula(worksheet, 43, 1d, $"MONTH({dateSerialText})");
                AssertNumericFormula(worksheet, 44, 2d, $"DAY({dateSerialText})");
                AssertNumericFormula(worksheet, 45, 12d, "HOUR(0.5)");
                AssertNumericFormula(worksheet, 46, 0d, "MINUTE(0.5)");
                AssertNumericFormula(worksheet, 47, 0d, "SECOND(0.5)");
                AssertNumericFormula(worksheet, 48, 1d, "ROW(A1)");
                AssertNumericFormula(worksheet, 49, 2d, "COLUMN(B1)");
                AssertNumericFormula(worksheet, 50, Math.Atan(1d), "ATAN(1)", precision: 12);
                AssertNumericFormula(worksheet, 51, 2d, "ROWS(A1:A2)");
                AssertNumericFormula(worksheet, 52, 2d, "COLUMNS(A1:B1)");
                AssertNumericFormula(worksheet, 53, Math.Sqrt(0.5d), "STDEV(A1:A2)", precision: 12);
                AssertNumericFormula(worksheet, 54, 3d, "LARGE(A1:A2,1)");
                AssertNumericFormula(worksheet, 55, 2d, "COUNTBLANK(B1:B2)");
                AssertNumericFormula(worksheet, 56, 1d, "COUNTIF(A1:A2,\">2\")");
                AssertNumericFormula(worksheet, 57, 1d, "RSQ(A1:A2,A1:A2)");
                AssertNumericFormula(worksheet, 58, dateSerial, "DATEVALUE(\"1/2/2026\")");
                AssertNumericFormula(worksheet, 59, 0.5d, "VAR(A1:A2)");
                AssertNumericFormula(worksheet, 60, 2d / 1.1d + 3d / (1.1d * 1.1d), "NPV(0.1,A1:A2)", precision: 12);
                AssertNumericFormula(worksheet, 61, 30d, "INDEX(C1:D2,2,2)");
                AssertNumericFormula(worksheet, 62, 2d, "MATCH(3,A1:A2,0)");
                AssertNumericFormula(worksheet, 63, 30d, "VLOOKUP(3,C1:D2,2,FALSE)");
                AssertNumericFormula(worksheet, 64, 300d, "HLOOKUP(3,C4:D5,2,FALSE)");
                AssertNumericFormula(worksheet, 65, 5d, "SUBTOTAL(9,A1:A2)");
                AssertNumericFormula(worksheet, 66, 3d, "SUMIF(A1:A2,\">2\")");
                AssertNumericFormula(worksheet, 67, 20d, "IF(A1>2,10,20)");
                AssertNumericFormula(worksheet, 68, 5d, "IF(A2>=3,A1+A2,0)");
                AssertNumericFormula(worksheet, 69, 5d, "ROUND(A1+A2,0)");
                AssertNumericFormula(worksheet, 70, 9d, "SUM(A1+A2,4)");
                AssertNumericFormula(worksheet, 71, 20d, "CHOOSE(A1,10,20,30)");
                AssertNumericFormula(worksheet, 72, 5d, "CHOOSE(1,A1+A2,10)");
                AssertNumericFormula(worksheet, 73, 3d, "OFFSET(A1,1,0)");
                AssertNumericFormula(worksheet, 74, 5d, "SUM(OFFSET(A1,0,0,2,1))");
                AssertNumericFormula(worksheet, 75, 8d, "A1^A2");
                AssertNumericFormula(worksheet, 76, 25d, "(A1+A2)^2");
                AssertNumericFormula(worksheet, 77, 10d, "IF(A1>2,10,)");
                AssertNumericFormula(worksheet, 78, 2d, "COUNT(A1,,A2)");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesFormulasReferencingDefinedNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("NamedFormulas");
                    sheet.CellValue(1, 1, 2d);
                    sheet.CellValue(2, 1, 3d);
                    sheet.CellValue(1, 2, 9d);

                    document.SetNamedRange("SalesBlock", "'NamedFormulas'!A1:A2", save: false);
                    document.SetNamedRange("LocalBase", "B1", sheet, save: false);

                    sheet.CellValue(3, 1, 5d);
                    sheet.CellFormula(3, 1, "SUM(SalesBlock)");
                    sheet.CellValue(4, 1, 10d);
                    sheet.CellFormula(4, 1, "LocalBase+1");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell globalNameFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 3 && cell.Column == 1);
                Assert.True(globalNameFormula.IsFormula);
                Assert.Equal(5d, Assert.IsType<double>(globalNameFormula.Value));
                Assert.Equal("SUM(SalesBlock)", globalNameFormula.FormulaText);

                LegacyXlsCell localNameFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 4 && cell.Column == 1);
                Assert.True(localNameFormula.IsFormula);
                Assert.Equal(10d, Assert.IsType<double>(localNameFormula.Value));
                Assert.Equal("LocalBase+1", localNameFormula.FormulaText);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookInternalSheetQualifiedFormulas() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet data = document.AddWorksheet("Data Sheet");
                    data.CellValue(1, 1, 2d);
                    data.CellValue(2, 1, 3d);

                    ExcelSheet region1 = document.AddWorksheet("Region 1");
                    region1.CellValue(1, 1, 2d);
                    region1.CellValue(2, 1, 4d);

                    ExcelSheet region2 = document.AddWorksheet("Region 2");
                    region2.CellValue(1, 1, 3d);
                    region2.CellValue(2, 1, 6d);

                    ExcelSheet calc = document.AddWorksheet("Calc");
                    calc.CellValue(1, 1, 5d);
                    calc.CellFormula(1, 1, "'Data Sheet'!A1+'Data Sheet'!A2");
                    calc.CellValue(2, 1, 5d);
                    calc.CellFormula(2, 1, "SUM('Data Sheet'!A1:A2)");
                    calc.CellValue(3, 1, 2d);
                    calc.CellFormula(3, 1, "'Data Sheet'!$A$1");
                    calc.CellValue(4, 1, 5d);
                    calc.CellFormula(4, 1, "SUM('Region 1:Region 2'!A1)");
                    calc.CellValue(5, 1, 15d);
                    calc.CellFormula(5, 1, "SUM('Region 1:Region 2'!A1:A2)");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets, sheet => sheet.Name == "Calc");
                AssertNumericFormula(worksheet, 1, 5d, "'Data Sheet'!A1+'Data Sheet'!A2");
                AssertNumericFormula(worksheet, 2, 5d, "SUM('Data Sheet'!A1:A2)");
                AssertNumericFormula(worksheet, 3, 2d, "'Data Sheet'!$A$1");
                AssertNumericFormula(worksheet, 4, 5d, "SUM('Region 1:Region 2'!A1)");
                AssertNumericFormula(worksheet, 5, 15d, "SUM('Region 1:Region 2'!A1:A2)");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSupportedBooleanAndTextFormulas() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("FormulaKinds");
                    sheet.CellValue(1, 1, 2d);
                    sheet.CellValue(2, 1, 3d);
                    sheet.CellValue(3, 1, true);
                    sheet.CellFormula(3, 1, "A2>A1");
                    sheet.CellValue(4, 1, "AlphaBeta");
                    sheet.CellFormula(4, 1, "\"Alpha\"&\"Beta\"");
                    sheet.CellValue(5, 1, true);
                    sheet.CellFormula(5, 1, "TRUE()");
                    sheet.CellValue(6, 1, false);
                    sheet.CellFormula(6, 1, "FALSE()");
                    sheet.CellValue(7, 1, false);
                    sheet.CellFormula(7, 1, "AND(TRUE,FALSE)");
                    sheet.CellValue(8, 1, true);
                    sheet.CellFormula(8, 1, "OR(FALSE,TRUE)");
                    sheet.CellValue(9, 1, true);
                    sheet.CellFormula(9, 1, "NOT(FALSE)");
                    sheet.CellValue(10, 1, "AlphaBeta");
                    sheet.CellFormula(10, 1, "CONCATENATE(\"Alpha\",\"Beta\")");
                    sheet.CellValue(11, 1, "Al");
                    sheet.CellFormula(11, 1, "LEFT(\"Alpha\",2)");
                    sheet.CellValue(12, 1, "ta");
                    sheet.CellFormula(12, 1, "RIGHT(\"Beta\",2)");
                    sheet.CellValue(13, 1, "lph");
                    sheet.CellFormula(13, 1, "MID(\"Alpha\",2,3)");
                    sheet.CellValue(14, 1, "HaHa");
                    sheet.CellFormula(14, 1, "REPT(\"Ha\",2)");
                    sheet.CellValue(15, 1, "2.5");
                    sheet.CellFormula(15, 1, "TEXT(2.5,\"0.0\")");
                    sheet.CellValue(16, 1, 5d);
                    sheet.CellFormula(16, 1, "LEN(\"Alpha\")");
                    sheet.CellValue(17, 1, 42d);
                    sheet.CellFormula(17, 1, "VALUE(\"42\")");
                    sheet.CellValue(18, 1, "alpha");
                    sheet.CellFormula(18, 1, "LOWER(\"Alpha\")");
                    sheet.CellValue(19, 1, "ALPHA");
                    sheet.CellFormula(19, 1, "UPPER(\"Alpha\")");
                    sheet.CellValue(20, 1, "Alpha Beta");
                    sheet.CellFormula(20, 1, "PROPER(\"alpha beta\")");
                    sheet.CellValue(21, 1, "Alpha Beta");
                    sheet.CellFormula(21, 1, "TRIM(\" Alpha  Beta \")");
                    sheet.CellValue(22, 1, "AlZZa");
                    sheet.CellFormula(22, 1, "REPLACE(\"Alpha\",3,2,\"ZZ\")");
                    sheet.CellValue(23, 1, 3d);
                    sheet.CellFormula(23, 1, "SEARCH(\"ph\",\"Alpha\")");
                    sheet.CellValue(24, 1, 3d);
                    sheet.CellFormula(24, 1, "FIND(\"ph\",\"Alpha\")");
                    sheet.CellValue(25, 1, "AlXXa");
                    sheet.CellFormula(25, 1, "SUBSTITUTE(\"Alpha\",\"ph\",\"XX\")");
                    sheet.CellValue(26, 1, true);
                    sheet.CellFormula(26, 1, "ISBLANK(B1)");
                    sheet.CellValue(27, 1, true);
                    sheet.CellFormula(27, 1, "ISNUMBER(A1)");
                    sheet.CellValue(28, 1, true);
                    sheet.CellFormula(28, 1, "ISTEXT(A10)");
                    sheet.CellValue(29, 1, true);
                    sheet.CellFormula(29, 1, "ISLOGICAL(A3)");
                    sheet.CellValue(30, 1, true);
                    sheet.CellFormula(30, 1, "ISNONTEXT(A1)");
                    sheet.CellValue(31, 1, true);
                    sheet.CellFormula(31, 1, "ISREF(A1)");
                    sheet.CellValue(32, 1, true);
                    sheet.CellFormula(32, 1, "ISNA(#N/A)");
                    sheet.CellValue(33, 1, true);
                    sheet.CellFormula(33, 1, "ISERR(#VALUE!)");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);

                LegacyXlsCell booleanFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 3 && cell.Column == 1);
                Assert.True(booleanFormula.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Boolean, booleanFormula.Kind);
                Assert.True(Assert.IsType<bool>(booleanFormula.Value));
                Assert.Equal("A2>A1", booleanFormula.FormulaText);

                LegacyXlsCell textFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 4 && cell.Column == 1);
                Assert.True(textFormula.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Text, textFormula.Kind);
                Assert.Equal("AlphaBeta", Assert.IsType<string>(textFormula.Value));
                Assert.Equal("\"Alpha\"&\"Beta\"", textFormula.FormulaText);

                LegacyXlsCell trueFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 5 && cell.Column == 1);
                Assert.True(trueFormula.IsFormula);
                Assert.True(Assert.IsType<bool>(trueFormula.Value));
                Assert.Equal("TRUE()", trueFormula.FormulaText);

                LegacyXlsCell falseFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 6 && cell.Column == 1);
                Assert.True(falseFormula.IsFormula);
                Assert.False(Assert.IsType<bool>(falseFormula.Value));
                Assert.Equal("FALSE()", falseFormula.FormulaText);

                LegacyXlsCell andFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 7 && cell.Column == 1);
                Assert.True(andFormula.IsFormula);
                Assert.False(Assert.IsType<bool>(andFormula.Value));
                Assert.Equal("AND(TRUE,FALSE)", andFormula.FormulaText);

                LegacyXlsCell orFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 8 && cell.Column == 1);
                Assert.True(orFormula.IsFormula);
                Assert.True(Assert.IsType<bool>(orFormula.Value));
                Assert.Equal("OR(FALSE,TRUE)", orFormula.FormulaText);

                LegacyXlsCell notFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 9 && cell.Column == 1);
                Assert.True(notFormula.IsFormula);
                Assert.True(Assert.IsType<bool>(notFormula.Value));
                Assert.Equal("NOT(FALSE)", notFormula.FormulaText);

                AssertTextFormula(worksheet, 10, "AlphaBeta", "CONCATENATE(\"Alpha\",\"Beta\")");
                AssertTextFormula(worksheet, 11, "Al", "LEFT(\"Alpha\",2)");
                AssertTextFormula(worksheet, 12, "ta", "RIGHT(\"Beta\",2)");
                AssertTextFormula(worksheet, 13, "lph", "MID(\"Alpha\",2,3)");
                AssertTextFormula(worksheet, 14, "HaHa", "REPT(\"Ha\",2)");
                AssertTextFormula(worksheet, 15, "2.5", "TEXT(2.5,\"0.0\")");
                AssertNumericFormula(worksheet, 16, 5d, "LEN(\"Alpha\")");
                AssertNumericFormula(worksheet, 17, 42d, "VALUE(\"42\")");
                AssertTextFormula(worksheet, 18, "alpha", "LOWER(\"Alpha\")");
                AssertTextFormula(worksheet, 19, "ALPHA", "UPPER(\"Alpha\")");
                AssertTextFormula(worksheet, 20, "Alpha Beta", "PROPER(\"alpha beta\")");
                AssertTextFormula(worksheet, 21, "Alpha Beta", "TRIM(\" Alpha  Beta \")");
                AssertTextFormula(worksheet, 22, "AlZZa", "REPLACE(\"Alpha\",3,2,\"ZZ\")");
                AssertNumericFormula(worksheet, 23, 3d, "SEARCH(\"ph\",\"Alpha\")");
                AssertNumericFormula(worksheet, 24, 3d, "FIND(\"ph\",\"Alpha\")");
                AssertTextFormula(worksheet, 25, "AlXXa", "SUBSTITUTE(\"Alpha\",\"ph\",\"XX\")");
                AssertBooleanFormula(worksheet, 26, expectedValue: true, "ISBLANK(B1)");
                AssertBooleanFormula(worksheet, 27, expectedValue: true, "ISNUMBER(A1)");
                AssertBooleanFormula(worksheet, 28, expectedValue: true, "ISTEXT(A10)");
                AssertBooleanFormula(worksheet, 29, expectedValue: true, "ISLOGICAL(A3)");
                AssertBooleanFormula(worksheet, 30, expectedValue: true, "ISNONTEXT(A1)");
                AssertBooleanFormula(worksheet, 31, expectedValue: true, "ISREF(A1)");
                AssertBooleanFormula(worksheet, 32, expectedValue: true, "ISNA(#N/A)");
                AssertBooleanFormula(worksheet, 33, expectedValue: true, "ISERR(#VALUE!)");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSupportedErrorCellsAndFormulaErrors() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Errors");
                    sheet.CellValue(1, 1, 1d);

                    sheet.CellValue(2, 1, "#N/A");
                    SetOpenXmlErrorCell(sheet, "A2", "#N/A");

                    sheet.CellValue(3, 1, "#DIV/0!");
                    sheet.CellFormula(3, 1, "A1/0");
                    SetOpenXmlErrorCell(sheet, "A3", "#DIV/0!");
                    sheet.CellValue(4, 1, "#N/A");
                    sheet.CellFormula(4, 1, "NA()");
                    SetOpenXmlErrorCell(sheet, "A4", "#N/A");
                    sheet.CellValue(5, 1, true);
                    sheet.CellFormula(5, 1, "ISERROR(#N/A)");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);

                LegacyXlsCell scalarError = Assert.Single(worksheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
                Assert.False(scalarError.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Error, scalarError.Kind);
                Assert.Equal("#N/A", Assert.IsType<string>(scalarError.Value));

                LegacyXlsCell formulaError = Assert.Single(worksheet.Cells, cell => cell.Row == 3 && cell.Column == 1);
                Assert.True(formulaError.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Error, formulaError.Kind);
                Assert.Equal("#DIV/0!", Assert.IsType<string>(formulaError.Value));
                Assert.Equal("A1/0", formulaError.FormulaText);

                LegacyXlsCell naFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 4 && cell.Column == 1);
                Assert.True(naFormula.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Error, naFormula.Kind);
                Assert.Equal("#N/A", Assert.IsType<string>(naFormula.Value));
                Assert.Equal("NA()", naFormula.FormulaText);

                LegacyXlsCell isErrorFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 5 && cell.Column == 1);
                Assert.True(isErrorFormula.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Boolean, isErrorFormula.Kind);
                Assert.True(Assert.IsType<bool>(isErrorFormula.Value));
                Assert.Equal("ISERROR(#N/A)", isErrorFormula.FormulaText);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesVolatileFormulasWithCachedResults() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            double todaySerial = new DateTime(2026, 1, 2).ToOADate();
            double nowSerial = todaySerial + 0.5d;

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Volatile");
                    sheet.CellValue(1, 1, 0.42d);
                    sheet.CellFormula(1, 1, "RAND()");
                    sheet.CellValue(2, 1, todaySerial);
                    sheet.CellFormula(2, 1, "TODAY()");
                    sheet.CellValue(3, 1, nowSerial);
                    sheet.CellFormula(3, 1, "NOW()");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);

                AssertNumericFormula(worksheet, 1, 0.42d, "RAND()", precision: 12);
                AssertNumericFormula(worksheet, 2, todaySerial, "TODAY()");
                AssertNumericFormula(worksheet, 3, nowSerial, "NOW()");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesDateSystemAndExplicitOpenXmlDateCells() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            DateTime date = new DateTime(2024, 2, 3);

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    document.DateSystem = ExcelDateSystem.NineteenFour;
                    ExcelSheet sheet = document.AddWorksheet("Dates");
                    sheet.CellValue(1, 1, date.ToString("O", CultureInfo.InvariantCulture));
                    sheet.CellAt(1, 1).SetNumberFormat("yyyy-mm-dd");
                    SetOpenXmlDateCell(sheet, "A1", date.ToString("O", CultureInfo.InvariantCulture));

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.True(result.Workbook.Uses1904DateSystem);
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell dateCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);

                Assert.False(dateCell.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Number, dateCell.Kind);
                double serial = Assert.IsType<double>(dateCell.Value);
                Assert.Equal(date.ToOADate() - 1462d, serial, 6);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public async Task LegacyXls_NativeSave_WritesLegacyXlsStreamsWhenRequested() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using ExcelDocument document = ExcelDocument.Create(openXmlPath);
                ExcelSheet sheet = document.AddWorksheet("Streamed");
                sheet.CellValue(1, 1, "Sync");
                sheet.CellValue(2, 1, 42d);
                var options = new ExcelSaveOptions();

                using var syncStream = new MemoryStream();
                document.Save(syncStream, ExcelFileFormat.Xls, options);
                AssertLegacyXlsStreamCell(syncStream, 1, 1, "Sync");

                sheet.CellValue(1, 1, "Async");
                using var asyncStream = new MemoryStream();
                await document.SaveAsync(asyncStream, ExcelFileFormat.Xls, options);
                AssertLegacyXlsStreamCell(asyncStream, 1, 1, "Async");
            } finally {
                TryDelete(openXmlPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesDocumentProperties() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            DateTime created = new DateTime(2026, 6, 26, 8, 15, 0, DateTimeKind.Utc);
            DateTime lastPrinted = new DateTime(2026, 6, 26, 9, 45, 0, DateTimeKind.Utc);
            DateTime reviewedAt = new DateTime(2026, 6, 26, 10, 30, 0, DateTimeKind.Utc);
            byte[] binaryPayload = { 0x00, 0x01, 0x42, 0x80, 0xff };

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Metadata");
                    sheet.CellValue(1, 1, "Document metadata");
                    document.BuiltinDocumentProperties.Title = "Native Metadata Workbook";
                    document.BuiltinDocumentProperties.Subject = "Native XLS metadata parity";
                    document.BuiltinDocumentProperties.Creator = "OfficeIMO Native Writer";
                    document.BuiltinDocumentProperties.Keywords = "xls, native, metadata";
                    document.BuiltinDocumentProperties.Description = "Native XLS SummaryInformation comments";
                    document.BuiltinDocumentProperties.Category = "Native Category";
                    document.BuiltinDocumentProperties.LastModifiedBy = "Native Reviewer";
                    document.BuiltinDocumentProperties.Revision = "12";
                    document.BuiltinDocumentProperties.Created = created;
                    document.BuiltinDocumentProperties.LastPrinted = lastPrinted;
                    document.ApplicationProperties.Company = "EvotecIT";
                    document.ApplicationProperties.Manager = "Workbook Manager";
                    document.SetCustomDocumentProperty("ReleaseStatus", "Ready");
                    document.SetCustomDocumentProperty("Ticket", 1998);
                    document.SetCustomDocumentProperty("Score", 98.5d);
                    document.SetCustomDocumentProperty("Reviewed", true);
                    document.SetCustomDocumentProperty("ReviewedAt", reviewedAt);
                    document.SetCustomDocumentProperty("BinaryPayload", binaryPayload);
                    document.SetCustomDocumentProperty("SignedByte", new ExcelCustomProperty((sbyte)-12, ExcelCustomPropertyType.NumberInteger));
                    document.SetCustomDocumentProperty("SignedShort", new ExcelCustomProperty((short)-32000, ExcelCustomPropertyType.NumberInteger));
                    document.SetCustomDocumentProperty("UnsignedByte", new ExcelCustomProperty((byte)250, ExcelCustomPropertyType.NumberInteger));
                    document.SetCustomDocumentProperty("UnsignedShort", new ExcelCustomProperty((ushort)65000, ExcelCustomPropertyType.NumberInteger));
                    document.SetCustomDocumentProperty("UnsignedInt32", new ExcelCustomProperty(4000000000U, ExcelCustomPropertyType.NumberInteger));
                    document.SetCustomDocumentProperty("UnsignedInt64", new ExcelCustomProperty(ulong.MaxValue, ExcelCustomPropertyType.NumberInteger));

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                ExcelDocument loaded = result.Document;
                AssertCompoundRootTreeContains(
                    xlsOutputPath,
                    "\u0005SummaryInformation",
                    "\u0005DocumentSummaryInformation",
                    "Workbook");
                Assert.Equal("Native Metadata Workbook", loaded.BuiltinDocumentProperties.Title);
                Assert.Equal("Native XLS metadata parity", loaded.BuiltinDocumentProperties.Subject);
                Assert.Equal("OfficeIMO Native Writer", loaded.BuiltinDocumentProperties.Creator);
                Assert.Equal("xls, native, metadata", loaded.BuiltinDocumentProperties.Keywords);
                Assert.Equal("Native XLS SummaryInformation comments", loaded.BuiltinDocumentProperties.Description);
                Assert.Equal("Native Category", loaded.BuiltinDocumentProperties.Category);
                Assert.Equal("Native Reviewer", loaded.BuiltinDocumentProperties.LastModifiedBy);
                Assert.Equal("12", loaded.BuiltinDocumentProperties.Revision);
                AssertSameInstant(created, loaded.BuiltinDocumentProperties.Created);
                AssertSameInstant(lastPrinted, loaded.BuiltinDocumentProperties.LastPrinted);
                Assert.Equal("EvotecIT", loaded.ApplicationProperties.Company);
                Assert.Equal("Workbook Manager", loaded.ApplicationProperties.Manager);
                Assert.Equal("Ready", loaded.CustomDocumentProperties["ReleaseStatus"].Text);
                Assert.Equal(1998, loaded.CustomDocumentProperties["Ticket"].NumberInteger);
                Assert.Equal(98.5d, loaded.CustomDocumentProperties["Score"].NumberDouble);
                Assert.True(loaded.CustomDocumentProperties["Reviewed"].Bool);
                AssertSameInstant(reviewedAt, loaded.CustomDocumentProperties["ReviewedAt"].Date);
                Assert.Equal(ExcelCustomPropertyType.Binary, loaded.CustomDocumentProperties["BinaryPayload"].PropertyType);
                Assert.Equal(binaryPayload, loaded.CustomDocumentProperties["BinaryPayload"].Binary);
                Assert.Equal(-12, loaded.CustomDocumentProperties["SignedByte"].Value);
                Assert.Equal(-32000, loaded.CustomDocumentProperties["SignedShort"].Value);
                Assert.Equal((byte)250, loaded.CustomDocumentProperties["UnsignedByte"].Value);
                Assert.Equal((ushort)65000, loaded.CustomDocumentProperties["UnsignedShort"].Value);
                Assert.Equal(4000000000U, loaded.CustomDocumentProperties["UnsignedInt32"].Value);
                Assert.Equal(ulong.MaxValue, loaded.CustomDocumentProperties["UnsignedInt64"].Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookAndWorksheetProtection() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Protected");
                    sheet.CellValue(1, 1, "Protected content");
                    document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                        ProtectStructure = true,
                        ProtectWindows = true,
                        LegacyPasswordHash = "CAFE"
                    });
                    sheet.Protect(new ExcelSheetProtectionOptions {
                        LegacyPasswordHash = "BEEF",
                        ProtectObjects = true,
                        ProtectScenarios = false
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorkbook workbook = result.Workbook;
                Assert.NotNull(workbook.Protection);
                Assert.True(workbook.Protection!.IsProtected);
                Assert.Equal("CAFE", workbook.Protection.LegacyPasswordHash);
                Assert.True(workbook.WindowsLocked);

                LegacyXlsWorksheet worksheet = Assert.Single(workbook.Worksheets);
                Assert.NotNull(worksheet.Protection);
                Assert.True(worksheet.Protection!.IsProtected);
                Assert.Equal("BEEF", worksheet.Protection.LegacyPasswordHash);
                Assert.True(worksheet.Protection.ProtectObjects);
                Assert.False(worksheet.Protection.ProtectScenarios);

                ExcelDocument loaded = result.Document;
                Assert.True(loaded.IsWorkbookProtected);
                Assert.True(loaded.Sheets.Single().IsProtected);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookProtectionPasswordWithoutLocks() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("PasswordOnly");
                    sheet.CellValue(1, 1, "Password-only workbook protection metadata");
                    document.WorkbookRoot.Append(new DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection {
                        WorkbookPassword = "CAFE"
                    });
                    document.WorkbookRoot.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorkbook workbook = result.Workbook;
                Assert.NotNull(workbook.Protection);
                Assert.False(workbook.Protection!.IsProtected);
                Assert.Equal("CAFE", workbook.Protection.LegacyPasswordHash);
                Assert.Null(workbook.WindowsLocked);

                DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection projectedProtection = result.Document.WorkbookRoot
                    .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection>()!;
                Assert.Equal("CAFE", projectedProtection.WorkbookPassword!.Value);
                Assert.Null(projectedProtection.LockStructure);
                Assert.Null(projectedProtection.LockWindows);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetProtectionPermissionExceptions() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Protected");
                    sheet.CellValue(1, 1, "Protected table editing");
                    sheet.ProtectTableEditing();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                Assert.NotNull(worksheet.Protection);
                Assert.True(worksheet.Protection!.IsProtected);
                Assert.NotNull(worksheet.Protection.Permissions);
                Assert.True(worksheet.Protection.Permissions!.AllowSelectLockedCells);
                Assert.True(worksheet.Protection.Permissions.AllowSelectUnlockedCells);
                Assert.True(worksheet.Protection.Permissions.AllowInsertRows);
                Assert.True(worksheet.Protection.Permissions.AllowSort);
                Assert.True(worksheet.Protection.Permissions.AllowAutoFilter);
                Assert.False(worksheet.Protection.Permissions.AllowFormatCells);

                ExcelWorksheetProtectionSnapshot projectedProtection = Assert.Single(result.Document.CreateInspectionSnapshot().Worksheets).Protection!;
                Assert.True(projectedProtection.AllowSelectLockedCells);
                Assert.True(projectedProtection.AllowSelectUnlockedCells);
                Assert.True(projectedProtection.AllowInsertRows);
                Assert.True(projectedProtection.AllowSort);
                Assert.True(projectedProtection.AllowAutoFilter);
                Assert.False(projectedProtection.AllowFormatCells);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookRevisionProtection() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Revision");
                    sheet.CellValue(1, 1, "Revision protection");
                    document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                        ProtectStructure = true,
                        LegacyPasswordHash = "CAFE"
                    });

                    DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection protection = document.WorkbookRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection>()!;
                    protection.LockRevision = true;
                    protection.RevisionsPassword = "BEEF";
                    document.WorkbookRoot.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorkbook workbook = result.Workbook;
                Assert.NotNull(workbook.Protection);
                Assert.True(workbook.Protection!.IsProtected);
                Assert.Equal("CAFE", workbook.Protection.LegacyPasswordHash);
                Assert.True(workbook.RevisionTrackingLocked);
                Assert.Equal((ushort)0xBEEF, workbook.RevisionTrackingPasswordHash);
                Assert.Contains(workbook.MetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.RevisionProtection);
                Assert.Contains(workbook.MetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.RevisionProtectionPassword);

                DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection projectedProtection = result.Document.WorkbookRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection>()!;
                Assert.True(projectedProtection.LockStructure!.Value);
                Assert.Equal("CAFE", projectedProtection.WorkbookPassword!.Value);
                Assert.True(projectedProtection.LockRevision!.Value);
                Assert.Equal("BEEF", projectedProtection.RevisionsPassword!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookWriteReservation() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Reservation");
                    sheet.CellValue(1, 1, "Write reservation");
                    document.SetWriteReservation(new ExcelWorkbookWriteReservationOptions {
                        ReadOnlyRecommended = true,
                        UserName = "Native Writer",
                        LegacyPasswordHash = "CAFE"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                Assert.NotNull(result.Workbook.WriteReservation);
                Assert.True(result.Workbook.WriteReservation!.ReadOnlyRecommended);
                Assert.Equal("CAFE", result.Workbook.WriteReservation.LegacyPasswordHash);
                Assert.Equal("Native Writer", result.Workbook.WriteReservation.UserName);

                ExcelWorkbookWriteReservationInfo reservation = result.Document.GetWriteReservation();
                Assert.True(reservation.Exists);
                Assert.True(reservation.ReadOnlyRecommended);
                Assert.Equal("Native Writer", reservation.UserName);
                Assert.Equal("CAFE", reservation.LegacyPasswordHash);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWriteAccessUserNameWithoutReservationPassword() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Reservation");
                    sheet.CellValue(1, 1, "Write access");
                    document.SetWriteReservation(new ExcelWorkbookWriteReservationOptions {
                        ReadOnlyRecommended = true,
                        UserName = "Reservation Owner"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                Assert.NotNull(result.Workbook.WriteReservation);
                Assert.True(result.Workbook.WriteReservation!.ReadOnlyRecommended);
                Assert.Null(result.Workbook.WriteReservation.LegacyPasswordHash);
                Assert.Null(result.Workbook.WriteReservation.UserName);
                Assert.Equal("Reservation Owner", result.Workbook.LastWriteUserName);
                Assert.Contains(result.Workbook.MetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.FileSharing);
                Assert.Contains(result.Workbook.MetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.WriteAccess);

                ExcelWorkbookWriteReservationInfo reservation = result.Document.GetWriteReservation();
                Assert.True(reservation.Exists);
                Assert.True(reservation.ReadOnlyRecommended);
                Assert.Equal("Reservation Owner", reservation.UserName);
                Assert.Null(reservation.LegacyPasswordHash);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalAndInternalHyperlinks() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet links = document.AddWorksheet("Links");
                    ExcelSheet target = document.AddWorksheet("Target");
                    target.CellValue(2, 2, "Target cell");
                    links.SetHyperlink(1, 1, "https://evotec.xyz/xls", "Evotec", style: false);
                    links.SetInternalLink(2, 1, target, "B2", "Jump", style: false);
                    links.CellValue(3, 1, "Spec");
                    AddExternalHyperlink(links, "A3", "../docs/spec.pdf", UriKind.Relative);
                    links.CellValue(4, 1, "Unicode Spec");
                    AddExternalHyperlink(links, "A4", "../docs/zażółć.xlsx", UriKind.Relative);
                    links.SetHyperlink(5, 1, "mailto:support@officeimo.net", "Support", style: false);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorksheet worksheet = result.Workbook.Worksheets[0];
                Assert.Equal(5, worksheet.Hyperlinks.Count);
                LegacyXlsHyperlink externalLink = Assert.Single(worksheet.Hyperlinks, link => link.StartRow == 1 && link.StartColumn == 1);
                LegacyXlsHyperlink internalLink = Assert.Single(worksheet.Hyperlinks, link => link.StartRow == 2 && link.StartColumn == 1);
                LegacyXlsHyperlink relativeLink = Assert.Single(worksheet.Hyperlinks, link => link.StartRow == 3 && link.StartColumn == 1);
                LegacyXlsHyperlink unicodeRelativeLink = Assert.Single(worksheet.Hyperlinks, link => link.StartRow == 4 && link.StartColumn == 1);
                LegacyXlsHyperlink mailtoLink = Assert.Single(worksheet.Hyperlinks, link => link.StartRow == 5 && link.StartColumn == 1);
                Assert.True(externalLink.IsExternal);
                Assert.Equal("https://evotec.xyz/xls", externalLink.Target);
                Assert.False(internalLink.IsExternal);
                Assert.Equal("'Target'!B2", internalLink.Target);
                Assert.True(relativeLink.IsExternal);
                Assert.Equal("../docs/spec.pdf", relativeLink.Target);
                Assert.True(unicodeRelativeLink.IsExternal);
                Assert.Equal("../docs/zażółć.xlsx", unicodeRelativeLink.Target);
                Assert.True(mailtoLink.IsExternal);
                Assert.Equal("mailto:support@officeimo.net", mailtoLink.Target);

                LegacyXlsCell externalCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                LegacyXlsCell internalCell = Assert.Single(worksheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
                LegacyXlsCell relativeCell = Assert.Single(worksheet.Cells, cell => cell.Row == 3 && cell.Column == 1);
                LegacyXlsCell unicodeRelativeCell = Assert.Single(worksheet.Cells, cell => cell.Row == 4 && cell.Column == 1);
                LegacyXlsCell mailtoCell = Assert.Single(worksheet.Cells, cell => cell.Row == 5 && cell.Column == 1);
                Assert.Equal("Evotec", Assert.IsType<string>(externalCell.Value));
                Assert.Equal("Jump", Assert.IsType<string>(internalCell.Value));
                Assert.Equal("Spec", Assert.IsType<string>(relativeCell.Value));
                Assert.Equal("Unicode Spec", Assert.IsType<string>(unicodeRelativeCell.Value));
                Assert.Equal("Support", Assert.IsType<string>(mailtoCell.Value));

                IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> projectedLinks = result.Document.Sheets[0].GetHyperlinks();
                Assert.True(projectedLinks["A1"].IsExternal);
                Assert.Equal("https://evotec.xyz/xls", projectedLinks["A1"].Target);
                Assert.False(projectedLinks["A2"].IsExternal);
                Assert.Equal("'Target'!B2", projectedLinks["A2"].Target);
                Assert.True(projectedLinks["A3"].IsExternal);
                Assert.Equal("../docs/spec.pdf", projectedLinks["A3"].Target);
                Assert.True(projectedLinks["A4"].IsExternal);
                Assert.Equal("../docs/zażółć.xlsx", projectedLinks["A4"].Target);
                Assert.True(projectedLinks["A5"].IsExternal);
                Assert.Equal("mailto:support@officeimo.net", projectedLinks["A5"].Target);

                byte[] expectedHLinkClsid = {
                    0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11,
                    0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b
                };
                Assert.All(GetBiffRecordPayloads(xlsOutputPath, 0x01b8), payload => {
                    Assert.True(payload.Length >= 24);
                    Assert.Equal(expectedHLinkClsid, payload.Skip(8).Take(16).ToArray());
                });
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesHyperlinkTooltips() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet links = document.AddWorksheet("Links");
                    ExcelSheet target = document.AddWorksheet("Target");
                    links.SetHyperlink(1, 1, "https://officeimo.net/legacy-xls", "OfficeIMO", style: false, tooltip: "Open OfficeIMO XLS docs");
                    links.SetInternalLink(2, 1, target, "A1", "Jump", style: false, tooltip: "Jump inside workbook");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorksheet worksheet = result.Workbook.Worksheets[0];
                Assert.Equal(2, worksheet.Hyperlinks.Count);
                LegacyXlsHyperlink externalLink = Assert.Single(worksheet.Hyperlinks, link => link.StartRow == 1 && link.StartColumn == 1);
                LegacyXlsHyperlink internalLink = Assert.Single(worksheet.Hyperlinks, link => link.StartRow == 2 && link.StartColumn == 1);
                Assert.Equal("https://officeimo.net/legacy-xls", externalLink.Target);
                Assert.Equal("Open OfficeIMO XLS docs", externalLink.Tooltip);
                Assert.Equal("'Target'!A1", internalLink.Target);
                Assert.Equal("Jump inside workbook", internalLink.Tooltip);

                IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> projectedLinks = result.Document.Sheets[0].GetHyperlinks();
                Assert.Equal("Open OfficeIMO XLS docs", projectedLinks["A1"].Tooltip);
                Assert.Equal("Jump inside workbook", projectedLinks["A2"].Tooltip);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetAutoFilterRangeAndCriteria() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Filtered");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(1, 3, "Notes");
                    sheet.CellValue(2, 1, "Open");
                    sheet.CellValue(2, 2, 12d);
                    sheet.CellValue(2, 3, "Ready");
                    sheet.CellValue(3, 1, "Pending");
                    sheet.CellValue(3, 2, 18d);
                    sheet.CellValue(3, 3, "Review");
                    sheet.CellValue(4, 1, "Closed");
                    sheet.CellValue(4, 2, 25d);
                    sheet.CellValue(4, 3, "Done");
                    sheet.AutoFilterAdd("A1:C4");
                    sheet.AutoFilterByHeadersEquals(("Status", new[] { "Open", "Pending" }));
                    sheet.AutoFilterByHeaderBetween("Amount", 10d, 20d);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal((ushort)3, legacySheet.AutoFilterDropDownCount);
                Assert.Equal(2, legacySheet.AutoFilterCriteria.Count);
                Assert.Contains(result.Workbook.DefinedNames, name =>
                    name.Name == "_FilterDatabase"
                    && name.Reference == "'Filtered'!$A$1:$C$4"
                    && name.LocalSheetIndex == 0
                    && name.Hidden
                    && name.BuiltIn);

                LegacyXlsAutoFilterCriteria statusCriteria = Assert.Single(legacySheet.AutoFilterCriteria, criteria => criteria.ColumnId == 0U);
                Assert.Equal(LegacyXlsAutoFilterKind.Custom, statusCriteria.Kind);
                Assert.Equal(LegacyXlsAutoFilterJoinOperator.Or, statusCriteria.JoinOperator);
                Assert.Equal(new[] { "Open", "Pending" }, statusCriteria.Conditions.Select(condition => condition.Value).ToArray());
                Assert.All(statusCriteria.Conditions, condition => Assert.Equal(LegacyXlsAutoFilterOperator.Equal, condition.Operator));

                LegacyXlsAutoFilterCriteria amountCriteria = Assert.Single(legacySheet.AutoFilterCriteria, criteria => criteria.ColumnId == 1U);
                Assert.Equal(LegacyXlsAutoFilterKind.Custom, amountCriteria.Kind);
                Assert.Equal(LegacyXlsAutoFilterJoinOperator.And, amountCriteria.JoinOperator);
                Assert.Equal(LegacyXlsAutoFilterOperator.GreaterThanOrEqual, amountCriteria.Conditions[0].Operator);
                Assert.Equal("10", amountCriteria.Conditions[0].Value);
                Assert.Equal(LegacyXlsAutoFilterOperator.LessThanOrEqual, amountCriteria.Conditions[1].Operator);
                Assert.Equal("20", amountCriteria.Conditions[1].Value);

                OpenXmlWorksheetPart worksheetPart = result.Document.Sheets[0].WorksheetPart;
                OpenXmlAutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<OpenXmlAutoFilter>());
                Assert.Equal("A1:C4", autoFilter.Reference!.Value);
                List<OpenXmlFilterColumn> filterColumns = autoFilter.Elements<OpenXmlFilterColumn>().OrderBy(column => column.ColumnId?.Value ?? 0U).ToList();
                Assert.Equal(2, filterColumns.Count);
                Assert.Equal(new[] { "Open", "Pending" }, filterColumns[0].GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Filters>()!.Elements<OpenXmlFilter>().Select(filter => filter.Val!.Value).ToArray());
                OpenXmlCustomFilter lower = filterColumns[1].GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CustomFilters>()!.Elements<OpenXmlCustomFilter>().First();
                OpenXmlCustomFilter upper = filterColumns[1].GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CustomFilters>()!.Elements<OpenXmlCustomFilter>().Last();
                Assert.Equal(OpenXmlFilterOperatorValues.GreaterThanOrEqual, lower.Operator!.Value);
                Assert.Equal("10", lower.Val!.Value);
                Assert.Equal(OpenXmlFilterOperatorValues.LessThanOrEqual, upper.Operator!.Value);
                Assert.Equal("20", upper.Val!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetDataValidations() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Validation");
                    sheet.CellValue(1, 2, "Status");
                    sheet.CellValue(1, 3, "Quantity");
                    sheet.CellValue(1, 4, "Discount");
                    sheet.CellValue(1, 5, "Start");

                    sheet.ValidationList("B2:B5", new[] { "Open", "Closed", "Pending" });
                    sheet.SetDataValidationMessages("B2:B5", new ExcelDataValidationMessageOptions {
                        PromptTitle = "Status",
                        Prompt = "Pick a status.",
                        ErrorTitle = "Invalid status",
                        Error = "Use one of the listed statuses.",
                        ErrorStyle = OpenXmlDataValidationErrorStyleValues.Warning,
                        SuppressDropDown = true
                    });
                    sheet.ValidationWholeNumber("C2:C5", OpenXmlDataValidationOperatorValues.Between, 1, 10, errorTitle: "Invalid quantity", errorMessage: "Use 1-10.");
                    sheet.ValidationDecimal("D2:D5", OpenXmlDataValidationOperatorValues.GreaterThan, 5.5d, allowBlank: false, errorTitle: "Invalid discount", errorMessage: "Use a value greater than 5.5.");
                    sheet.ValidationTime("E2:E5", OpenXmlDataValidationOperatorValues.Equal, TimeSpan.FromHours(9), allowBlank: true, errorTitle: "Invalid time", errorMessage: "Use 09:00.");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsDataValidationCollectionRecord collection = Assert.Single(legacySheet.DataValidationCollections);
                Assert.Equal(4U, collection.DeclaredValidationCount);
                Assert.Equal(4, legacySheet.DataValidations.Count);

                LegacyXlsDataValidation listValidation = Assert.Single(legacySheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.List);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, listValidation.ListItems.ToArray());
                Assert.Equal("B2:B5", Assert.Single(listValidation.Ranges));
                Assert.True(listValidation.AllowBlank);
                Assert.True(listValidation.ShowInputMessage);
                Assert.True(listValidation.ShowErrorMessage);
                Assert.True(listValidation.SuppressDropDown);
                Assert.Equal(LegacyXlsDataValidationErrorStyle.Warning, listValidation.ErrorStyle);
                Assert.Equal("Status", listValidation.PromptTitle);
                Assert.Equal("Pick a status.", listValidation.Prompt);
                Assert.Equal("Invalid status", listValidation.ErrorTitle);

                LegacyXlsDataValidation wholeValidation = Assert.Single(legacySheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.WholeNumber);
                Assert.Equal(LegacyXlsDataValidationOperator.Between, wholeValidation.Operator);
                Assert.Equal("1", wholeValidation.Formula1);
                Assert.Equal("10", wholeValidation.Formula2);
                Assert.Equal("C2:C5", Assert.Single(wholeValidation.Ranges));

                LegacyXlsDataValidation decimalValidation = Assert.Single(legacySheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.Decimal);
                Assert.Equal(LegacyXlsDataValidationOperator.GreaterThan, decimalValidation.Operator);
                Assert.Equal("5.5", decimalValidation.Formula1);
                Assert.False(decimalValidation.AllowBlank);

                LegacyXlsDataValidation timeValidation = Assert.Single(legacySheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.Time);
                Assert.Equal(LegacyXlsDataValidationOperator.Equal, timeValidation.Operator);
                Assert.Equal(TimeSpan.FromHours(9).TotalDays.ToString(CultureInfo.InvariantCulture), timeValidation.Formula1);

                ExcelSheet projectedSheet = result.Document.Sheets[0];
                ExcelDataValidationInfo projectedList = Assert.Single(projectedSheet.GetDataValidations("B2:B5"));
                Assert.Equal("list", projectedList.Type);
                Assert.Equal("\"Open,Closed,Pending\"", projectedList.Formula1);
                Assert.True(projectedList.SuppressDropDown);
                Assert.Equal("warning", projectedList.ErrorStyle);

                ExcelDataValidationInfo projectedWhole = Assert.Single(projectedSheet.GetDataValidations("C2:C5"));
                Assert.Equal("whole", projectedWhole.Type);
                Assert.Equal("between", projectedWhole.Operator);
                Assert.Equal("1", projectedWhole.Formula1);
                Assert.Equal("10", projectedWhole.Formula2);

                ExcelDataValidationInfo projectedDecimal = Assert.Single(projectedSheet.GetDataValidations("D2:D5"));
                Assert.Equal("decimal", projectedDecimal.Type);
                Assert.Equal("greaterThan", projectedDecimal.Operator);
                Assert.Equal("5.5", projectedDecimal.Formula1);

                ExcelDataValidationInfo projectedTime = Assert.Single(projectedSheet.GetDataValidations("E2:E5"));
                Assert.Equal("time", projectedTime.Type);
                Assert.Equal("equal", projectedTime.Operator);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_RebasesDataValidationFormulaReferencesToSqrefAnchor() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Validation");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 1, 1d);
                    sheet.ValidationCustomFormula("A2:A5", "A2>0");

                    document.Save(xlsOutputPath);
                }

                byte[] payload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x01be));
                byte[] formulaTokens = ReadDataValidationFormulaTokens(payload, formulaIndex: 0);
                Assert.Contains((byte)0x4c, formulaTokens);
                Assert.DoesNotContain((byte)0x44, formulaTokens);

                int refOffset = Array.IndexOf(formulaTokens, (byte)0x4c);
                Assert.True(refOffset >= 0);
                Assert.Equal((ushort)0, ReadUInt16(formulaTokens, refOffset + 1));
                Assert.Equal((ushort)0xc000, ReadUInt16(formulaTokens, refOffset + 3));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetConditionalFormatting() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Conditions");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.CellValue(2, 2, 8d);
                    sheet.CellValue(3, 2, 12d);
                    sheet.CellValue(1, 3, "Band");
                    sheet.CellValue(2, 3, 3d);
                    sheet.CellValue(3, 3, 7d);
                    sheet.CellValue(1, 4, "Formula");
                    sheet.CellValue(2, 4, 1d);
                    sheet.CellValue(3, 4, -1d);

                    sheet.AddConditionalRule("B2:B5", OpenXmlConditionalFormattingOperatorValues.GreaterThan, "10");
                    sheet.AddConditionalRule("C2:C5", OpenXmlConditionalFormattingOperatorValues.Between, "1", "5");
                    sheet.AddConditionalFormulaRule("D2:D5", "D2>0");

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal(3, legacySheet.ConditionalFormattings.Count);

                LegacyXlsConditionalFormatting greaterThanRule = Assert.Single(legacySheet.ConditionalFormattings, rule =>
                    rule.Type == LegacyXlsConditionalFormattingType.CellIs
                    && rule.Operator == LegacyXlsConditionalFormattingOperator.GreaterThan);
                Assert.Equal("B2:B5", Assert.Single(greaterThanRule.Ranges));
                Assert.Equal("10", greaterThanRule.Formula1);
                Assert.Null(greaterThanRule.Formula2);

                LegacyXlsConditionalFormatting betweenRule = Assert.Single(legacySheet.ConditionalFormattings, rule =>
                    rule.Type == LegacyXlsConditionalFormattingType.CellIs
                    && rule.Operator == LegacyXlsConditionalFormattingOperator.Between);
                Assert.Equal("C2:C5", Assert.Single(betweenRule.Ranges));
                Assert.Equal("1", betweenRule.Formula1);
                Assert.Equal("5", betweenRule.Formula2);

                LegacyXlsConditionalFormatting expressionRule = Assert.Single(legacySheet.ConditionalFormattings, rule =>
                    rule.Type == LegacyXlsConditionalFormattingType.Formula);
                Assert.Equal("D2:D5", Assert.Single(expressionRule.Ranges));
                Assert.Equal("D2>0", expressionRule.Formula1);

                ExcelSheet projectedSheet = result.Document.Sheets[0];
                ExcelConditionalFormattingInfo projectedGreaterThan = Assert.Single(projectedSheet.GetConditionalFormattingRules("B2:B5"));
                Assert.Equal("CellIs", projectedGreaterThan.Type);
                Assert.Equal(nameof(OpenXmlConditionalFormattingOperatorValues.GreaterThan), projectedGreaterThan.Operator);
                Assert.Equal(new[] { "10" }, projectedGreaterThan.Formulas);

                ExcelConditionalFormattingInfo projectedBetween = Assert.Single(projectedSheet.GetConditionalFormattingRules("C2:C5"));
                Assert.Equal("CellIs", projectedBetween.Type);
                Assert.Equal(nameof(OpenXmlConditionalFormattingOperatorValues.Between), projectedBetween.Operator);
                Assert.Equal(new[] { "1", "5" }, projectedBetween.Formulas);

                ExcelConditionalFormattingInfo projectedExpression = Assert.Single(projectedSheet.GetConditionalFormattingRules("D2:D5"));
                Assert.Equal("Expression", projectedExpression.Type);
                Assert.Equal(new[] { "D2>0" }, projectedExpression.Formulas);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_RebasesConditionalFormattingFormulaReferencesToSqrefAnchor() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Conditions");
                    sheet.CellValue(2, 1, 1d);
                    sheet.AddConditionalFormulaRule("A2:A5", "A2>0");

                    document.Save(xlsOutputPath);
                }

                byte[] payload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x01b1));
                byte[] formulaTokens = ReadConditionalFormattingFormulaTokens(payload, formulaIndex: 0);
                Assert.Contains((byte)0x4c, formulaTokens);
                Assert.DoesNotContain((byte)0x44, formulaTokens);

                int refOffset = Array.IndexOf(formulaTokens, (byte)0x4c);
                Assert.True(refOffset >= 0);
                Assert.Equal((ushort)0, ReadUInt16(formulaTokens, refOffset + 1));
                Assert.Equal((ushort)0xc000, ReadUInt16(formulaTokens, refOffset + 3));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetComments() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Comments");
                    sheet.CellValue(1, 1, "Review");
                    sheet.CellValue(2, 2, "Unicode");
                    sheet.SetComment(1, 1, "Review this cell", author: "Alice", initials: "AK");
                    sheet.SetComment(2, 2, "Zażółć note", author: "Żaneta");

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal(2, legacySheet.Comments.Count);
                byte[] drawingGroupPayload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x00eb));
                Assert.True(drawingGroupPayload.Length > 8);
                IReadOnlyList<byte[]> drawingPayloads = GetBiffRecordPayloads(xlsOutputPath, 0x00ec);
                Assert.Equal(2, drawingPayloads.Count);
                Assert.All(drawingPayloads, payload => AssertCommentDrawingInfo(payload, expectedShapeCount: 2, expectedLastShapeId: 1026));
                AssertBiffRecordOccursBefore(xlsOutputPath, 0x00eb, 0x0085);

                LegacyXlsComment firstLegacyComment = Assert.Single(legacySheet.Comments, comment => comment.Row == 1 && comment.Column == 1);
                Assert.Equal("Review this cell", firstLegacyComment.Text);
                Assert.Equal("Alice (AK)", firstLegacyComment.Author);
                Assert.True(firstLegacyComment.HasAnchor);

                LegacyXlsComment secondLegacyComment = Assert.Single(legacySheet.Comments, comment => comment.Row == 2 && comment.Column == 2);
                Assert.Equal("Zażółć note", secondLegacyComment.Text);
                Assert.Equal("Żaneta", secondLegacyComment.Author);
                Assert.True(secondLegacyComment.HasAnchor);

                ExcelSheet projectedSheet = result.Document.Sheets[0];
                ExcelCommentInfo firstProjectedComment = Assert.Single(projectedSheet.GetComments(), comment => comment.CellReference == "A1");
                Assert.Equal("Review this cell", firstProjectedComment.Text);
                Assert.Equal("Alice (AK)", firstProjectedComment.Author);

                ExcelCommentInfo secondProjectedComment = Assert.Single(projectedSheet.GetComments(), comment => comment.CellReference == "B2");
                Assert.Equal("Zażółć note", secondProjectedComment.Text);
                Assert.Equal("Żaneta", secondProjectedComment.Author);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesRichTextWorksheetComments() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Rich Comments");
                    sheet.CellValue(1, 1, "Review");
                    sheet.SetCommentRichText(
                        1,
                        1,
                        new[] {
                            new ExcelRichTextRun("Important ") {
                                Bold = true,
                                FontColor = "#123456",
                                FontName = "Consolas",
                                FontSize = 13D
                            },
                            new ExcelRichTextRun("note") {
                                Italic = true,
                                Underline = true,
                                FontName = "Arial",
                                FontSize = 11D
                            }
                        },
                        author: "Alice",
                        initials: "AK");

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsComment legacyComment = Assert.Single(legacySheet.Comments);
                Assert.Equal("Important note", legacyComment.Text);
                Assert.Equal("Alice (AK)", legacyComment.Author);
                Assert.Equal(2, legacyComment.FormattingRuns.Count);

                ExcelCommentInfo projectedComment = Assert.Single(result.Document.Sheets[0].GetComments());
                Assert.Equal("Important note", projectedComment.Text);
                Assert.Equal("Alice (AK)", projectedComment.Author);
                Assert.Equal(2, projectedComment.RichTextRuns.Count);

                ExcelRichTextRun firstRun = projectedComment.RichTextRuns[0];
                Assert.Equal("Important ", firstRun.Text);
                Assert.True(firstRun.Bold);
                Assert.False(firstRun.Italic);
                Assert.Equal("FF123456", firstRun.FontColor);
                Assert.Equal("Consolas", firstRun.FontName);
                Assert.Equal(13D, firstRun.FontSize);

                ExcelRichTextRun secondRun = projectedComment.RichTextRuns[1];
                Assert.Equal("note", secondRun.Text);
                Assert.False(secondRun.Bold);
                Assert.True(secondRun.Italic);
                Assert.True(secondRun.Underline);
                Assert.Equal("Arial", secondRun.FontName);
                Assert.Equal(11D, secondRun.FontSize);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesDefinedNamesAndPrintNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet names = document.AddWorksheet("Names");
                    ExcelSheet scoped = document.AddWorksheet("Scoped");
                    names.CellValue(1, 1, "North");
                    names.CellValue(2, 2, 125d);
                    scoped.CellValue(3, 3, "Local");

                    document.SetNamedRange("SalesBlock", "'Names'!A1:B2", save: false);
                    document.SetNamedRange("LocalCell", "C3", scoped, save: false, hidden: true);
                    document.SetPrintArea(names, "A1:C5", save: false);
                    document.SetPrintTitles(names, firstRow: 1, lastRow: 2, firstCol: 1, lastCol: 2, save: false);

                    document.Save(xlsOutputPath);
                }

                LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                Assert.Equal(4, result.Workbook.DefinedNames.Count);

                LegacyXlsDefinedName globalName = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "SalesBlock");
                Assert.Null(globalName.LocalSheetIndex);
                Assert.False(globalName.Hidden);
                Assert.False(globalName.BuiltIn);
                Assert.Equal("'Names'!$A$1:$B$2", globalName.Reference);

                LegacyXlsDefinedName localName = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "LocalCell");
                Assert.Equal(1, localName.LocalSheetIndex);
                Assert.True(localName.Hidden);
                Assert.False(localName.BuiltIn);
                Assert.Equal("'Scoped'!$C$3", localName.Reference);

                LegacyXlsDefinedName printArea = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "_xlnm.Print_Area");
                Assert.Equal(0, printArea.LocalSheetIndex);
                Assert.True(printArea.BuiltIn);
                Assert.Equal("'Names'!$A$1:$C$5", printArea.Reference);

                LegacyXlsDefinedName printTitles = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "_xlnm.Print_Titles");
                Assert.Equal(0, printTitles.LocalSheetIndex);
                Assert.True(printTitles.BuiltIn);
                Assert.Equal("'Names'!$1:$2,'Names'!$A:$B", printTitles.Reference);

                Assert.Equal("'Names'!$A$1:$B$2", result.Document.GetNamedRange("SalesBlock"));
                Assert.Equal("$C$3", result.Document.GetNamedRange("LocalCell", result.Document.Sheets[1]));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksFormulaWithoutNumericCachedResult() {
            AssertNativeXlsSaveNotSupported("cached result", (document, sheet) => {
                sheet.CellValue(1, 1, 2d);
                sheet.CellValue(2, 1, 3d);
                sheet.CellFormula(3, 1, "A1+A2");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalWorkbookFormulaReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalRefs");
                    sheet.CellValue(1, 1, 2d);
                    sheet.CellFormula(1, 1, "'[Budget.xls]Other'!$A$1");
                    sheet.CellValue(2, 1, 5d);
                    sheet.CellFormula(2, 1, "SUM('[Budget.xls]Other'!$A$1:$B$2)");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Budget.xls", externalReference.Target);
                Assert.Equal(new[] { "Other" }, externalReference.SheetNames);

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertNumericFormula(worksheet, 1, 2d, "'[Budget.xls]Other'!$A$1");
                AssertNumericFormula(worksheet, 2, 5d, "SUM('[Budget.xls]Other'!$A$1:$B$2)");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalDefinedNameFormulaReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalNames");
                    sheet.CellValue(1, 1, 0.25d);
                    sheet.CellFormula(1, 1, "[Budget.xls]TaxRate");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Budget.xls", externalReference.Target);
                Assert.Equal(new[] { "Sheet1" }, externalReference.SheetNames);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("TaxRate", externalName.Name);
                Assert.Equal(LegacyXlsExternalNameBodyKind.ExternalDefinedName, externalName.BodyKind);

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertNumericFormula(worksheet, 1, 0.25d, "'Budget.xls'!TaxRate");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.TokenName == "PtgNameX"
                    && token.OperandKind == "ExternalName");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSelfSupBookForInternal3DReferencesWithoutExternalLinks() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet first = document.AddWorksheet("First");
                    ExcelSheet second = document.AddWorksheet("Second");
                    second.CellValue(1, 1, 42d);
                    first.CellValue(1, 1, 42d);
                    first.CellFormula(1, 1, "Second!A1");

                    document.Save(xlsOutputPath);
                }

                byte[] supBookPayload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x01ae));
                Assert.Equal(4, supBookPayload.Length);
                Assert.Equal((ushort)2, ReadUInt16(supBookPayload, 0));
                Assert.Equal((ushort)0x0401, ReadUInt16(supBookPayload, 2));

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference selfReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.Self);
                Assert.Equal(2, selfReference.SheetCount);
                Assert.DoesNotContain(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets, sheet => sheet.Name == "First");
                AssertNumericFormula(worksheet, 1, 42d, "'Second'!A1");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_FormulaEncoder_UsesUnionBeforeIntersectionForReferenceOperators() {
            Assert.True(
                LegacyXlsFormulaEncoder.TryEncode("A1,B1 C1", LegacyXlsFormulaNameIndex.Empty, formulaSheetIndex: -1, out byte[] tokens, out string? reason),
                reason);

            Assert.Equal((byte)0x10, tokens[tokens.Length - 1]);
            Assert.Contains((byte)0x0f, tokens);
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksFutureFunctionAliasesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("future-function aliases", (document, sheet) => {
                sheet.CellValue(1, 1, "North");
                sheet.CellValue(2, 1, "South");
                sheet.CellValue(1, 2, 10d);
                sheet.CellValue(2, 2, 20d);
                sheet.CellFormula(1, 3, "_xlfn.XLOOKUP(A1,A1:A2,B1:B2)");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedFormulaTokenPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("formula token payload lengths outside BIFF8 limits", (document, sheet) => {
                string longLiteral = "\"" + new string('A', 255) + "\"";
                string formula = string.Join("&", Enumerable.Repeat(longLiteral, 260));

                sheet.CellValue(1, 1, "cached text");
                OpenXmlCell cell = sheet.WorksheetPart.Worksheet!.Descendants<OpenXmlCell>().Single(item => item.CellReference?.Value == "A1");
                cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula(formula);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSingleCellArrayFormula() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Array Formula");
                    sheet.CellValue(1, 1, 2d);
                    sheet.CellValue(2, 1, 3d);
                    sheet.CellValue(3, 1, 5d);
                    sheet.SetArrayFormula("A3:A3", "A1+A2");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertNumericFormula(worksheet, 3, 5d, "A1+A2");

                LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(worksheet.ArrayFormulaRecords);
                Assert.Equal("A3", arrayFormula.Range);
                Assert.Equal(1, arrayFormula.DeclaredCellCount);
                Assert.Equal(1, arrayFormula.MatchedFormulaCellCount);
                Assert.True(arrayFormula.FormulaTextProjected);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesMultiCellArrayFormulaWithCachedResults() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Array Formula");
                    sheet.CellValue(1, 1, 2d);
                    sheet.CellValue(2, 1, 3d);
                    sheet.CellValue(1, 2, 5d);
                    sheet.SetArrayFormula("B1:B2", "SUM(A1:A2)");
                    sheet.CellValue(2, 2, 5d);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertArrayFormulaCell(worksheet, 1, 2, 5d, "SUM(A1:A2)");
                AssertArrayFormulaCell(worksheet, 2, 2, 5d, "SUM(A1:A2)");

                LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(worksheet.ArrayFormulaRecords);
                Assert.Equal("B1:B2", arrayFormula.Range);
                Assert.Equal(2, arrayFormula.DeclaredCellCount);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksMultiCellArrayFormulaWithoutCachedResultsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cached results for every cell", (document, sheet) => {
                sheet.CellValue(1, 1, 2d);
                sheet.CellValue(2, 1, 3d);
                sheet.SetArrayFormula("A3:A4", "SUM(A1:A2)");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSharedFormulaRecords() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Shared Formula");
                    sheet.CellValue(1, 1, 10d);
                    sheet.CellValue(2, 1, 20d);
                    sheet.CellValue(3, 1, 30d);
                    sheet.CellValue(1, 2, 15d);
                    sheet.CellValue(2, 2, 25d);
                    sheet.CellValue(3, 2, 35d);
                    sheet.CellFormula(1, 2, "A1+5");
                    sheet.CellFormula(2, 2, "A2+5");
                    sheet.CellFormula(3, 2, "A3+5");

                    OpenXmlCell sharedRoot = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                        .Single(cell => string.Equals(cell.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase));
                    sharedRoot.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                    sharedRoot.CellFormula.SharedIndex = 0U;
                    sharedRoot.CellFormula.Reference = "B1:B3";

                    foreach (string reference in new[] { "B2", "B3" }) {
                        OpenXmlCell sharedFollower = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                            .Single(cell => string.Equals(cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
                        sharedFollower.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                        sharedFollower.CellFormula.SharedIndex = 0U;
                        sharedFollower.CellFormula.Text = string.Empty;
                    }

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertSharedFormulaCell(worksheet, 1, 15d, "A1+5");
                AssertSharedFormulaCell(worksheet, 2, 25d, "A2+5");
                AssertSharedFormulaCell(worksheet, 3, 35d, "A3+5");
                Assert.Equal(3, result.ImportReport.FormulaTokensByContext["SharedFormula"]);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_RebasesSharedFormula3DReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet data = document.AddWorksheet("Data");
                    data.CellValue(1, 1, 10d);
                    data.CellValue(2, 1, 20d);

                    ExcelSheet sheet = document.AddWorksheet("Shared Formula");
                    sheet.CellValue(1, 2, 15d);
                    sheet.CellValue(2, 2, 25d);
                    sheet.CellFormula(1, 2, "Data!A1+5");
                    sheet.CellFormula(2, 2, "Data!A2+5");

                    OpenXmlCell sharedRoot = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                        .Single(cell => string.Equals(cell.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase));
                    sharedRoot.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                    sharedRoot.CellFormula.SharedIndex = 0U;
                    sharedRoot.CellFormula.Reference = "B1:B2";

                    OpenXmlCell sharedFollower = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                        .Single(cell => string.Equals(cell.CellReference?.Value, "B2", StringComparison.OrdinalIgnoreCase));
                    sharedFollower.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                    sharedFollower.CellFormula.SharedIndex = 0U;
                    sharedFollower.CellFormula.Text = string.Empty;

                    document.Save(xlsOutputPath);
                }

                byte[] sharedFormulaPayload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x04bc));
                byte[] formulaTokens = new byte[ReadUInt16(sharedFormulaPayload, 8)];
                Buffer.BlockCopy(sharedFormulaPayload, 10, formulaTokens, 0, formulaTokens.Length);

                int ref3dOffset = Array.IndexOf(formulaTokens, (byte)0x5a);
                Assert.True(ref3dOffset >= 0);
                Assert.Equal((ushort)0, ReadUInt16(formulaTokens, ref3dOffset + 3));
                Assert.Equal((ushort)0xffff, ReadUInt16(formulaTokens, ref3dOffset + 5));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSharedFormulaUseCountForWholeRange() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Shared Formula");
                    sheet.CellValue(1, 1, 10d);
                    sheet.CellValue(2, 1, 20d);

                    int[,] sharedCells = { { 1, 2 }, { 1, 3 }, { 2, 2 }, { 2, 3 } };
                    for (int i = 0; i < sharedCells.GetLength(0); i++) {
                        sheet.CellValue(sharedCells[i, 0], sharedCells[i, 1], 15d);
                        sheet.CellFormula(sharedCells[i, 0], sharedCells[i, 1], "A1+5");
                    }

                    OpenXmlCell sharedRoot = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                        .Single(cell => string.Equals(cell.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase));
                    sharedRoot.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                    sharedRoot.CellFormula.SharedIndex = 0U;
                    sharedRoot.CellFormula.Reference = "B1:C2";

                    foreach (string reference in new[] { "C1", "B2", "C2" }) {
                        OpenXmlCell sharedFollower = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                            .Single(cell => string.Equals(cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
                        sharedFollower.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                        sharedFollower.CellFormula.SharedIndex = 0U;
                        sharedFollower.CellFormula.Text = string.Empty;
                    }

                    document.Save(xlsOutputPath);
                }

                byte[] sharedFormulaPayload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x04bc));
                Assert.Equal(4, sharedFormulaPayload[7]);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSharedFormulaRangesAboveBiffUseLimitBeforeWriting() {
            AssertNativeXlsSaveNotSupported("shared formula ranges with more than 255 cells", (document, sheet) => {
                for (int row = 1; row <= 256; row++) {
                    sheet.CellValue(row, 1, row);
                    sheet.CellValue(row, 2, row + 5);
                    sheet.CellFormula(row, 2, $"A{row}+5");
                }

                OpenXmlCell sharedRoot = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                    .Single(cell => string.Equals(cell.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase));
                sharedRoot.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                sharedRoot.CellFormula.SharedIndex = 0U;
                sharedRoot.CellFormula.Reference = "B1:B256";

                foreach (OpenXmlCell sharedFollower in sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                    .Where(cell => !string.Equals(cell.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase)
                        && cell.CellReference?.Value?.StartsWith("B", StringComparison.OrdinalIgnoreCase) == true)) {
                    sharedFollower.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                    sharedFollower.CellFormula.SharedIndex = 0U;
                    sharedFollower.CellFormula.Text = string.Empty;
                }

                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSharedFormulaWithoutDefinitionBeforeWriting() {
            AssertNativeXlsSaveNotSupported("shared formula definition", (document, sheet) => {
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(1, 2, 15d);
                sheet.CellFormula(1, 2, "A1+5");

                OpenXmlCell sharedFollower = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                    .Single(cell => string.Equals(cell.CellReference?.Value, "B1", StringComparison.OrdinalIgnoreCase));
                sharedFollower.CellFormula!.FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
                sharedFollower.CellFormula.SharedIndex = 0U;
                sharedFollower.CellFormula.Text = string.Empty;
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedFeatureFamiliesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter equality lists with more than two values", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(3, 1, "Beta");
                sheet.CellValue(4, 1, "Gamma");
                sheet.AddAutoFilter("A1:A2");
                sheet.AutoFilterByHeaderEquals("Name", new[] { "Alpha", "Beta", "Gamma" });
            });

            AssertNativeXlsSaveNotSupported("conditional formatting visual or extension payloads", (document, sheet) => {
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);
                sheet.AddConditionalColorScale("A1:A3", "FFFF0000", "FF00FF00");
            });

            AssertNativeXlsSaveNotSupported("unsupported external hyperlink targets", (document, sheet) => {
                sheet.CellValue(1, 1, "Script");
                AddExternalHyperlink(sheet, "A1", "javascript:alert(1)", UriKind.Absolute);
            });

            AssertNativeXlsSaveNotSupported("defined-name formulas outside the supported native XLS formula subset", (document, sheet) => {
                sheet.CellValue(1, 1, "Named");
                document.WorkbookRoot.DefinedNames ??= new DocumentFormat.OpenXml.Spreadsheet.DefinedNames();
                document.WorkbookRoot.DefinedNames.Append(new DocumentFormat.OpenXml.Spreadsheet.DefinedName {
                    Name = "FormulaName",
                    Text = "SUM(SalesTable[Amount])"
                });
            });

            AssertNativeXlsSaveNotSupported("tables", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AddTable("A1:A2", hasHeader: true, name: "UnsupportedTable", TableStyle.TableStyleMedium2);
            });
        }

        private static void AssertLegacyXlsStreamCell(MemoryStream stream, int row, int column, string expectedValue) {
            stream.Position = 0;
            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(stream);
            LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == column);
            Assert.Equal(expectedValue, Assert.IsType<string>(cell.Value));
        }

        private static LegacyXlsFont GetLegacyFont(LegacyXlsWorkbook workbook, ushort fontIndex) {
            int index = fontIndex < 4 ? fontIndex : fontIndex > 4 ? fontIndex - 1 : -1;
            Assert.InRange(index, 0, workbook.Fonts.Count - 1);
            return workbook.Fonts[index];
        }

        private static void AssertNumericFormula(LegacyXlsWorksheet worksheet, int row, double expectedValue, string expectedFormulaText, int? precision = null) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == 1);
            Assert.True(cell.IsFormula);
            double actualValue = Assert.IsType<double>(cell.Value);
            if (precision.HasValue) {
                Assert.Equal(expectedValue, actualValue, precision.Value);
            } else {
                Assert.Equal(expectedValue, actualValue);
            }

            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static void AssertTextFormula(LegacyXlsWorksheet worksheet, int row, string expectedValue, string expectedFormulaText) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == 1);
            Assert.True(cell.IsFormula);
            Assert.Equal(LegacyXlsCellValueKind.Text, cell.Kind);
            Assert.Equal(expectedValue, Assert.IsType<string>(cell.Value));
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static void AssertBooleanFormula(LegacyXlsWorksheet worksheet, int row, bool expectedValue, string expectedFormulaText) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == 1);
            Assert.True(cell.IsFormula);
            Assert.Equal(LegacyXlsCellValueKind.Boolean, cell.Kind);
            Assert.Equal(expectedValue, Assert.IsType<bool>(cell.Value));
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static void AssertSharedFormulaCell(LegacyXlsWorksheet worksheet, int row, double expectedValue, string expectedFormulaText) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == 2);
            Assert.True(cell.IsFormula);
            Assert.Equal(expectedValue, Assert.IsType<double>(cell.Value));
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static void AssertArrayFormulaCell(LegacyXlsWorksheet worksheet, int row, int column, double expectedValue, string expectedFormulaText) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == column);
            Assert.True(cell.IsFormula);
            Assert.Equal(expectedValue, Assert.IsType<double>(cell.Value));
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static byte[] ReadDataValidationFormulaTokens(byte[] payload, int formulaIndex) {
            Assert.True(formulaIndex == 0 || formulaIndex == 1);
            int offset = 4;
            for (int i = 0; i < 4; i++) {
                ushort characterCount = ReadUInt16(payload, offset);
                byte flags = payload[offset + 2];
                offset += 3 + (((flags & 0x01) != 0) ? characterCount * 2 : characterCount);
            }

            byte[] firstFormula = ReadDataValidationFormulaTokens(payload, ref offset);
            if (formulaIndex == 0) {
                return firstFormula;
            }

            return ReadDataValidationFormulaTokens(payload, ref offset);
        }

        private static byte[] ReadDataValidationFormulaTokens(byte[] payload, ref int offset) {
            ushort tokenLength = ReadUInt16(payload, offset);
            offset += 4;
            byte[] tokens = new byte[tokenLength];
            Buffer.BlockCopy(payload, offset, tokens, 0, tokens.Length);
            offset += tokenLength;
            return tokens;
        }

        private static byte[] ReadConditionalFormattingFormulaTokens(byte[] payload, int formulaIndex) {
            Assert.True(formulaIndex == 0 || formulaIndex == 1);
            ushort firstFormulaLength = ReadUInt16(payload, 2);
            ushort secondFormulaLength = ReadUInt16(payload, 4);
            int offset = 6;
            if (formulaIndex == 1) {
                offset += firstFormulaLength;
            }

            ushort tokenLength = formulaIndex == 0 ? firstFormulaLength : secondFormulaLength;
            byte[] tokens = new byte[tokenLength];
            Buffer.BlockCopy(payload, offset, tokens, 0, tokens.Length);
            return tokens;
        }

        private static IReadOnlyList<byte[]> GetBiffRecordPayloads(string xlsPath, ushort recordType) {
            byte[] fileBytes = File.ReadAllBytes(xlsPath);
            byte[] workbookStream = ReadCompoundStream(fileBytes, "Workbook");

            var payloads = new List<byte[]>();
            int offset = 0;
            while (offset + 4 <= workbookStream.Length) {
                ushort type = ReadUInt16(workbookStream, offset);
                ushort length = ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) {
                    break;
                }

                if (type == recordType) {
                    byte[] payload = new byte[length];
                    Buffer.BlockCopy(workbookStream, payloadOffset, payload, 0, length);
                    payloads.Add(payload);
                }

                offset = payloadOffset + length;
            }

            return payloads;
        }

        private static void AssertBiffRecordOccursBefore(string xlsPath, ushort beforeRecordType, ushort afterRecordType) {
            byte[] fileBytes = File.ReadAllBytes(xlsPath);
            byte[] workbookStream = ReadCompoundStream(fileBytes, "Workbook");

            int beforeOffset = -1;
            int afterOffset = -1;
            int offset = 0;
            while (offset + 4 <= workbookStream.Length) {
                ushort type = ReadUInt16(workbookStream, offset);
                ushort length = ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) {
                    break;
                }

                if (type == beforeRecordType && beforeOffset < 0) {
                    beforeOffset = offset;
                }

                if (type == afterRecordType && afterOffset < 0) {
                    afterOffset = offset;
                }

                offset = payloadOffset + length;
            }

            Assert.True(beforeOffset >= 0, $"BIFF record 0x{beforeRecordType:X4} was not found.");
            Assert.True(afterOffset >= 0, $"BIFF record 0x{afterRecordType:X4} was not found.");
            Assert.True(beforeOffset < afterOffset, $"BIFF record 0x{beforeRecordType:X4} should appear before 0x{afterRecordType:X4}.");
        }

        private static int GetBiffContentLength(byte[] workbookStream, int expectedEndOfFileRecords) {
            int endOfFileRecords = 0;
            int offset = 0;
            while (offset + 4 <= workbookStream.Length) {
                ushort type = ReadUInt16(workbookStream, offset);
                ushort length = ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                Assert.True(payloadOffset + length <= workbookStream.Length);

                offset = payloadOffset + length;
                if (type == 0x000a) {
                    endOfFileRecords++;
                    if (endOfFileRecords == expectedEndOfFileRecords) {
                        return offset;
                    }
                }
            }

            throw new InvalidOperationException("The expected BIFF EOF records were not found.");
        }

        private static void AssertCompoundRootTreeContains(string xlsPath, params string[] expectedNames) {
            byte[] fileBytes = File.ReadAllBytes(xlsPath);
            IReadOnlyList<CompoundDirectoryEntry> entries = ReadFirstCompoundDirectorySector(fileBytes);
            CompoundDirectoryEntry root = Assert.Single(entries, entry => entry.ObjectType == 5);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            TraverseCompoundDirectoryTree(entries, root.ChildId, names, new HashSet<uint>());

            foreach (string expectedName in expectedNames) {
                Assert.Contains(expectedName, names);
            }
        }

        private static IReadOnlyList<CompoundDirectoryEntry> ReadFirstCompoundDirectorySector(byte[] fileBytes) {
            int sectorSize = 1 << ReadUInt16(fileBytes, 30);
            uint directorySector = ReadUInt32(fileBytes, 48);
            int directoryOffset = checked(512 + ((int)directorySector * sectorSize));
            Assert.True(directoryOffset >= 512 && directoryOffset + sectorSize <= fileBytes.Length);

            var entries = new List<CompoundDirectoryEntry>();
            for (int offset = directoryOffset; offset + 128 <= directoryOffset + sectorSize; offset += 128) {
                ushort nameLength = ReadUInt16(fileBytes, offset + 64);
                byte objectType = fileBytes[offset + 66];
                string name = objectType == 0 || nameLength < 2 || nameLength > 64
                    ? string.Empty
                    : System.Text.Encoding.Unicode.GetString(fileBytes, offset, nameLength - 2);
                entries.Add(new CompoundDirectoryEntry(
                    unchecked((uint)entries.Count),
                    name,
                    objectType,
                    ReadUInt32(fileBytes, offset + 68),
                    ReadUInt32(fileBytes, offset + 72),
                    ReadUInt32(fileBytes, offset + 76),
                    ReadUInt64(fileBytes, offset + 120)));
            }

            return entries;
        }

        private static void TraverseCompoundDirectoryTree(
            IReadOnlyList<CompoundDirectoryEntry> entries,
            uint entryId,
            HashSet<string> names,
            HashSet<uint> visited) {
            const uint freeSector = 0xffffffff;
            const uint endOfChain = 0xfffffffe;
            if (entryId == freeSector || entryId == endOfChain || entryId >= entries.Count || !visited.Add(entryId)) {
                return;
            }

            CompoundDirectoryEntry entry = entries[(int)entryId];
            TraverseCompoundDirectoryTree(entries, entry.LeftSiblingId, names, visited);
            if (entry.ObjectType != 0 && !string.IsNullOrEmpty(entry.Name)) {
                names.Add(entry.Name);
            }

            TraverseCompoundDirectoryTree(entries, entry.RightSiblingId, names, visited);
        }

        private static void AssertCommentDrawingInfo(byte[] drawingPayload, uint expectedShapeCount, uint expectedLastShapeId) {
            Assert.True(drawingPayload.Length >= 24);
            Assert.Equal((ushort)0xf002, ReadUInt16(drawingPayload, 2));
            const int drawingInfoOffset = 8;
            Assert.Equal((ushort)0xf008, ReadUInt16(drawingPayload, drawingInfoOffset + 2));
            Assert.Equal(8U, ReadUInt32(drawingPayload, drawingInfoOffset + 4));
            Assert.Equal(expectedShapeCount, ReadUInt32(drawingPayload, drawingInfoOffset + 8));
            Assert.Equal(expectedLastShapeId, ReadUInt32(drawingPayload, drawingInfoOffset + 12));
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) {
            return unchecked((ushort)(bytes[offset] | (bytes[offset + 1] << 8)));
        }

        private static uint ReadUInt32(byte[] bytes, int offset) {
            return unchecked((uint)(bytes[offset]
                | (bytes[offset + 1] << 8)
                | (bytes[offset + 2] << 16)
                | (bytes[offset + 3] << 24)));
        }

        private static ulong ReadUInt64(byte[] bytes, int offset) {
            return ReadUInt32(bytes, offset) | ((ulong)ReadUInt32(bytes, offset + 4) << 32);
        }

        private readonly struct CompoundDirectoryEntry {
            internal CompoundDirectoryEntry(uint index, string name, byte objectType, uint leftSiblingId, uint rightSiblingId, uint childId, ulong streamSize) {
                Index = index;
                Name = name;
                ObjectType = objectType;
                LeftSiblingId = leftSiblingId;
                RightSiblingId = rightSiblingId;
                ChildId = childId;
                StreamSize = streamSize;
            }

            internal uint Index { get; }

            internal string Name { get; }

            internal byte ObjectType { get; }

            internal uint LeftSiblingId { get; }

            internal uint RightSiblingId { get; }

            internal uint ChildId { get; }

            internal ulong StreamSize { get; }
        }

        private static byte[] ReadCompoundStream(byte[] compoundBytes, string streamName) {
            Assert.True(
                OfficeCompoundFileReader.TryRead(compoundBytes, out OfficeCompoundFile? compoundFile, out string? error),
                error);
            Assert.True(compoundFile!.Streams.TryGetValue(streamName, out byte[]? stream), $"Compound stream '{streamName}' was not found.");
            return stream!;
        }

        private static void SetOpenXmlErrorCell(ExcelSheet sheet, string reference, string errorText) {
            OpenXmlCell cell = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                .Single(item => string.Equals(item.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
            cell.DataType = OpenXmlCellValues.Error;
            cell.CellValue = new OpenXmlCellValue(errorText);
        }

        private static void SetOpenXmlDateCell(ExcelSheet sheet, string reference, string dateText) {
            OpenXmlCell cell = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                .Single(item => string.Equals(item.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
            cell.DataType = OpenXmlCellValues.Date;
            cell.CellValue = new OpenXmlCellValue(dateText);
        }

        private static void AddExternalHyperlink(ExcelSheet sheet, string reference, string target, UriKind uriKind, string? tooltip = null) {
            DocumentFormat.OpenXml.Packaging.HyperlinkRelationship relationship = sheet.WorksheetPart.AddHyperlinkRelationship(new Uri(target, uriKind), true);
            DocumentFormat.OpenXml.Spreadsheet.Hyperlinks hyperlinks = sheet.WorksheetPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Hyperlinks>().FirstOrDefault()
                ?? sheet.WorksheetPart.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Hyperlinks());
            hyperlinks.Append(new DocumentFormat.OpenXml.Spreadsheet.Hyperlink {
                Reference = reference,
                Id = relationship.Id,
                Tooltip = string.IsNullOrWhiteSpace(tooltip) ? null : tooltip
            });
        }

        private static void AssertNativeXlsSaveNotSupported(string expectedMessagePart, Action<ExcelDocument, ExcelSheet> configure) {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using ExcelDocument document = ExcelDocument.Create(openXmlPath);
                ExcelSheet sheet = document.AddWorksheet("Unsupported");
                configure(document, sheet);

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(xlsOutputPath));
                Assert.True(
                    exception.Message.Contains(expectedMessagePart, StringComparison.OrdinalIgnoreCase),
                    exception.Message);
                Assert.False(File.Exists(xlsOutputPath));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        private static void AssertNativeXlsSignedSaveBlocked(Action<ExcelDocument, ExcelSheet> configure) {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using ExcelDocument document = ExcelDocument.Create(openXmlPath);
                ExcelSheet sheet = document.AddWorksheet("Signed");
                configure(document, sheet);

                ExcelSignedWorkbookMutationException exception = Assert.Throws<ExcelSignedWorkbookMutationException>(() =>
                    document.Save(xlsOutputPath));
                Assert.True(exception.SignatureInfo.HasSignatures);
                Assert.False(File.Exists(xlsOutputPath));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        private static string FormatUnsupportedFeatures(IEnumerable<LegacyXlsUnsupportedFeature> features) {
            return string.Join(
                Environment.NewLine,
                features.Select(feature =>
                    $"{feature.Kind}|{feature.Code}|{feature.Description}|Sheet:{feature.SheetName ?? string.Empty}|Record:{feature.RecordType?.ToString("X4", CultureInfo.InvariantCulture) ?? string.Empty}|Detail:{feature.DetailCode ?? string.Empty}"));
        }
    }
}
