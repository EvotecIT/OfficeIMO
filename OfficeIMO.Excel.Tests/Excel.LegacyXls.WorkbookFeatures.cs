using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_ImportsPhase4ExternalHyperlinks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4HyperlinkWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsHyperlink hyperlink = Assert.Single(sheet.Hyperlinks);
            Assert.Equal(1, hyperlink.StartRow);
            Assert.Equal(1, hyperlink.StartColumn);
            Assert.Equal(1, hyperlink.EndRow);
            Assert.Equal(1, hyperlink.EndColumn);
            Assert.Equal("https://officeimo.net/legacy-xls", hyperlink.Target);
            Assert.True(hyperlink.IsExternal);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? text));
            Assert.Equal("OfficeIMO", text);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Hyperlink projected = Assert.Single(worksheetPart.Worksheet.Descendants<Hyperlink>());
            Assert.Equal("A1", projected.Reference!.Value);
            HyperlinkRelationship relationship = Assert.Single(worksheetPart.HyperlinkRelationships);
            Assert.Equal(new Uri("https://officeimo.net/legacy-xls"), relationship.Uri);
            Assert.Equal(relationship.Id, projected.Id!.Value);
            Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference!.Value == "A1");
            Assert.NotNull(cell.StyleIndex);
        }

        [Fact]
        public void LegacyXls_Load_ImportsHyperlinkTooltips() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4HyperlinkTooltipWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            LegacyXlsHyperlink hyperlink = Assert.Single(Assert.Single(legacy.Worksheets).Hyperlinks);
            Assert.Equal("Open OfficeIMO XLS docs", hyperlink.Tooltip);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> projectedLinks = document.Sheets[0].GetHyperlinks();
            Assert.Equal("Open OfficeIMO XLS docs", projectedLinks["A1"].Tooltip);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Hyperlink projected = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Hyperlink>());
            Assert.Equal("Open OfficeIMO XLS docs", projected.Tooltip!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4InternalHyperlinks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4InternalHyperlinkWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            Assert.Equal(2, legacy.Worksheets.Count);
            LegacyXlsHyperlink hyperlink = Assert.Single(legacy.Worksheets[0].Hyperlinks);
            Assert.False(hyperlink.IsExternal);
            Assert.Equal("'Target'!B2", hyperlink.Target);
            Assert.Equal("Jump", hyperlink.DisplayText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? text));
            Assert.Equal("Jump", text);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First(part => part.Worksheet.Descendants<Hyperlink>().Any());
            Hyperlink projected = Assert.Single(worksheetPart.Worksheet.Descendants<Hyperlink>());
            Assert.Equal("A1", projected.Reference!.Value);
            Assert.Equal("'Target'!B2", projected.Location!.Value);
            Assert.Null(projected.Id);
            Assert.Empty(worksheetPart.HyperlinkRelationships);
            Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference!.Value == "A1");
            Assert.NotNull(cell.StyleIndex);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4FileHyperlinks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4FileHyperlinkWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            LegacyXlsHyperlink hyperlink = Assert.Single(Assert.Single(legacy.Worksheets).Hyperlinks);
            Assert.True(hyperlink.IsExternal);
            Assert.Equal(@"C:\Data\Budget.pdf", hyperlink.Target);
            Assert.Equal("Budget", hyperlink.DisplayText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Hyperlink projected = Assert.Single(worksheetPart.Worksheet.Descendants<Hyperlink>());
            Assert.Equal("A1", projected.Reference!.Value);
            HyperlinkRelationship relationship = Assert.Single(worksheetPart.HyperlinkRelationships);
            Assert.Equal(Uri.UriSchemeFile, relationship.Uri.Scheme);
            Assert.EndsWith("C:/Data/Budget.pdf", relationship.Uri.AbsolutePath, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(relationship.Id, projected.Id!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4UncFileHyperlinks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4UncFileHyperlinkWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            LegacyXlsHyperlink hyperlink = Assert.Single(Assert.Single(legacy.Worksheets).Hyperlinks);
            Assert.True(hyperlink.IsExternal);
            Assert.Equal(@"\\fileserver\share\Budget.pdf", hyperlink.Target);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            HyperlinkRelationship relationship = Assert.Single(worksheetPart.HyperlinkRelationships);
            Assert.Equal(Uri.UriSchemeFile, relationship.Uri.Scheme);
            Assert.Equal("fileserver", relationship.Uri.Host);
            Assert.Equal("/share/Budget.pdf", relationship.Uri.AbsolutePath);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4RelativeFileHyperlinks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4RelativeFileHyperlinkWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            LegacyXlsHyperlink hyperlink = Assert.Single(Assert.Single(legacy.Worksheets).Hyperlinks);
            Assert.True(hyperlink.IsExternal);
            Assert.Equal(@"..\Docs\Budget.pdf", hyperlink.Target);
            Assert.Equal("Relative Budget", hyperlink.DisplayText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Hyperlink projected = Assert.Single(worksheetPart.Worksheet.Descendants<Hyperlink>());
            Assert.Equal("A1", projected.Reference!.Value);
            HyperlinkRelationship relationship = Assert.Single(worksheetPart.HyperlinkRelationships);
            Assert.False(relationship.Uri.IsAbsoluteUri);
            Assert.Equal("../Docs/Budget.pdf", relationship.Uri.OriginalString);
            Assert.Equal(relationship.Id, projected.Id!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFileHyperlinkParentDirectoryCount() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4FileHyperlinkWithParentDirectoryCountWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsHyperlink hyperlink = Assert.Single(Assert.Single(legacy.Worksheets).Hyperlinks);
            Assert.True(hyperlink.IsExternal);
            Assert.Equal(@"..\Docs\Budget.pdf", hyperlink.Target);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            HyperlinkRelationship relationship = Assert.Single(worksheetPart.HyperlinkRelationships);
            Assert.False(relationship.Uri.IsAbsoluteUri);
            Assert.Equal("../Docs/Budget.pdf", relationship.Uri.OriginalString);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4CellComments() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4CommentWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-COMMENT-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsComment comment = Assert.Single(sheet.Comments);
            Assert.Equal(1, comment.Row);
            Assert.Equal(1, comment.Column);
            Assert.Equal(1, comment.ObjectId);
            Assert.Equal((ushort)0x0019, comment.ObjectType);
            Assert.Equal(LegacyXlsDrawingObjectType.Note, comment.ObjectTypeKind);
            Assert.Equal("Note", comment.ObjectTypeName);
            Assert.Equal((ushort)0x4011, comment.ObjectFlags);
            Assert.Contains("Locked", comment.ObjectFlagNames);
            Assert.Contains("Printable", comment.ObjectFlagNames);
            Assert.True(comment.IsObjectLocked);
            Assert.True(comment.IsObjectPrintable);
            Assert.Equal("Legacy Author", comment.Author);
            Assert.Equal("Imported legacy note", comment.Text);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.CommentsByObjectTypeName["Note"]);
            Assert.Equal(1, report.CommentsByObjectFlagName["Locked"]);
            Assert.Equal(1, report.CommentsByObjectFlagName["Printable"]);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelCommentInfo projectedComment = Assert.Single(document.Sheets[0].GetComments());
            Assert.Equal("A1", projectedComment.CellReference);
            Assert.Equal("Legacy Author", projectedComment.Author);
            Assert.Equal("Imported legacy note", projectedComment.Text);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetCommentsPart commentsPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().WorksheetCommentsPart!;
            Assert.NotNull(commentsPart);
            DocumentFormat.OpenXml.Spreadsheet.Comment openXmlComment = Assert.Single(commentsPart.Comments!.CommentList!.Elements<DocumentFormat.OpenXml.Spreadsheet.Comment>());
            Assert.Equal("A1", openXmlComment.Reference!.Value);
            Assert.Equal("Imported legacy note", openXmlComment.InnerText);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4WorksheetProtection() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4WorksheetProtectionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.NotNull(legacy.Protection);
            Assert.True(legacy.Protection!.IsProtected);
            Assert.Equal("CAFE", legacy.Protection.LegacyPasswordHash);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.NotNull(sheet.Protection);
            Assert.True(sheet.Protection!.IsProtected);
            Assert.Equal("BEEF", sheet.Protection.LegacyPasswordHash);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.IsWorkbookProtected);
            Assert.True(document.Sheets[0].IsProtected);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookProtection workbookProtection = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<WorkbookProtection>()!;
            Assert.True(workbookProtection.LockStructure!.Value);
            Assert.Equal("CAFE", workbookProtection.WorkbookPassword!.Value);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            SheetProtection protection = Assert.Single(worksheetPart.Worksheet.Elements<SheetProtection>());
            Assert.True(protection.Sheet!.Value);
            Assert.Equal("BEEF", protection.Password!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4PrintPageSetup() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4PrintPageSetupWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            Assert.NotNull(legacySheet.PageSetup);
            Assert.Equal(0.25d, legacySheet.PageSetup!.LeftMargin);
            Assert.Equal(0.35d, legacySheet.PageSetup.RightMargin);
            Assert.Equal(0.5d, legacySheet.PageSetup.TopMargin);
            Assert.Equal(0.6d, legacySheet.PageSetup.BottomMargin);
            Assert.Equal(0.4d, legacySheet.PageSetup.HeaderMargin);
            Assert.Equal(0.45d, legacySheet.PageSetup.FooterMargin);
            Assert.True(legacySheet.PageSetup.Landscape);
            Assert.Equal((ushort)125, legacySheet.PageSetup.Scale);
            Assert.Equal((ushort)1, legacySheet.PageSetup.FitToWidth);
            Assert.Equal((ushort)2, legacySheet.PageSetup.FitToHeight);
            Assert.Equal((ushort)2, legacySheet.PageSetup.PrintedSize);
            Assert.Contains(legacySheet.MetadataRecords, record => record.Kind == LegacyXlsWorksheetMetadataKind.PrinterSettings);
            Assert.Contains(legacySheet.MetadataRecords, record => record.Kind == LegacyXlsWorksheetMetadataKind.PrintSize);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == 0x0033);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == 0x004d);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelSheetPageSetup setup = document.Sheets[0].GetPageSetup();
            Assert.Equal(ExcelPageOrientation.Landscape, setup.Orientation);
            Assert.Equal((uint)125, setup.Scale);
            Assert.Equal((uint)1, setup.FitToWidth);
            Assert.Equal((uint)2, setup.FitToHeight);
            Assert.NotNull(setup.Margins);
            Assert.Equal(0.25d, setup.Margins!.Left);
            Assert.Equal(0.35d, setup.Margins.Right);
            Assert.Equal(0.5d, setup.Margins.Top);
            Assert.Equal(0.6d, setup.Margins.Bottom);
            Assert.Equal(0.4d, setup.Margins.Header);
            Assert.Equal(0.45d, setup.Margins.Footer);
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorksheetMetadataRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetMetadataWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.True(result.HasUnsupportedFeatures);
            LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
            Assert.Equal(6, legacySheet.MetadataRecords.Count);
            Assert.True(legacySheet.AutomaticPageBreaksVisible);
            Assert.True(legacySheet.ApplyOutlineStyles);
            Assert.True(legacySheet.SummaryRowsBelow);
            Assert.True(legacySheet.SummaryColumnsRightWhenLeftToRight);
            Assert.NotNull(legacySheet.PageSetup);
            Assert.True(legacySheet.PageSetup!.FitToPage);
            Assert.True(legacySheet.SynchronizedHorizontalScrolling);
            Assert.True(legacySheet.SynchronizedVerticalScrolling);
            Assert.True(legacySheet.TransitionFormulaEvaluation);
            Assert.True(legacySheet.TransitionFormulaEntry);
            Assert.Equal((byte)2, legacySheet.RowOutlineLevel);
            Assert.Equal((byte)3, legacySheet.ColumnOutlineLevel);
            Assert.True(legacySheet.GridSet);
            Assert.NotNull(legacySheet.RowBlockIndex);
            Assert.Equal(1, legacySheet.RowBlockIndex!.FirstRowIndex);
            Assert.Equal(5, legacySheet.RowBlockIndex.RowAfterLastIndex);
            Assert.Equal(1234U, legacySheet.RowBlockIndex.ReservedRecordOffset);
            Assert.Equal(2, legacySheet.RowBlockIndex.DbCellBlockCount);
            LegacyXlsSelection selection = Assert.Single(legacySheet.Selections);
            Assert.Equal(0, selection.Pane);
            Assert.Equal(3, selection.ActiveRow);
            Assert.Equal(2, selection.ActiveColumn);
            Assert.Equal((ushort)0, selection.ActiveRangeIndex);
            LegacyXlsSelectedRange selectedRange = Assert.Single(selection.SelectedRanges);
            Assert.Equal("B3:C4", selectedRange.Reference);
            LegacyXlsSortSettings sortSettings = Assert.IsType<LegacyXlsSortSettings>(legacySheet.SortSettings);
            Assert.True(sortSettings.SortLeftToRight);
            Assert.True(sortSettings.Key1Descending);
            Assert.False(sortSettings.Key2Descending);
            Assert.True(sortSettings.Key3Descending);
            Assert.True(sortSettings.CaseSensitive);
            Assert.Equal(3, sortSettings.CustomListIndex);
            Assert.True(sortSettings.UsePhoneticInformation);
            Assert.Equal("A2:A4", sortSettings.Key1);
            Assert.Equal("B2:B4", sortSettings.Key2);
            Assert.Equal("C2:C4", sortSettings.Key3);
            LegacyXlsUnsupportedFeature unsupportedSort = Assert.Single(result.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.WorksheetSort, unsupportedSort.Kind);
            Assert.Equal("XLS-BIFF-FEATURE-WORKSHEET-SORT-CUSTOM-LIST-UNSUPPORTED", unsupportedSort.Code);
            Assert.Equal((ushort)0x0090, unsupportedSort.RecordType);
            Assert.Equal("WorksheetSort:CustomListIndex", unsupportedSort.DetailCode);
            Assert.Equal(1, result.ImportReport.UnsupportedProjectionGapsByKind[LegacyXlsUnsupportedFeatureKind.WorksheetSort]);
            Assert.Equal(6, result.ImportReport.WorksheetMetadataRecordCount);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.SheetOptions]);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.OutlineLevels]);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.GridSet]);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.RowBlockIndex]);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.Selection]);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.Sort]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.RecordType is 0x0081 or 0x0080 or 0x0082 or 0x020b or 0x001d);

            using var output = new MemoryStream();
            result.Document.Save(output, new ExcelSaveOptions {
                LossPolicy = ExcelConversionLossPolicy.Allow
            });
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            PageSetupProperties pageSetupProperties = worksheetPart.Worksheet.GetFirstChild<SheetProperties>()!.GetFirstChild<PageSetupProperties>()!;
            Assert.True(pageSetupProperties.FitToPage!.Value);
            Selection projectedSelection = Assert.Single(worksheetPart.Worksheet.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!.Elements<Selection>());
            Assert.Null(projectedSelection.Pane);
            Assert.Equal("B3", projectedSelection.ActiveCell!.Value);
            Assert.Equal("B3:C4", projectedSelection.SequenceOfReferences!.InnerText);
            SortState projectedSort = Assert.IsType<SortState>(worksheetPart.Worksheet.GetFirstChild<SortState>());
            Assert.Equal("A2:C4", projectedSort.Reference!.Value);
            Assert.True(projectedSort.ColumnSort!.Value);
            Assert.True(projectedSort.CaseSensitive!.Value);
            Assert.Equal(SortMethodValues.PinYin, projectedSort.SortMethod!.Value);
            SortCondition[] projectedConditions = projectedSort.Elements<SortCondition>().ToArray();
            Assert.Equal(3, projectedConditions.Length);
            Assert.Equal("A2:A4", projectedConditions[0].Reference!.Value);
            Assert.True(projectedConditions[0].Descending!.Value);
            Assert.Equal("B2:B4", projectedConditions[1].Reference!.Value);
            Assert.Null(projectedConditions[1].Descending);
            Assert.Equal("C2:C4", projectedConditions[2].Reference!.Value);
            Assert.True(projectedConditions[2].Descending!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorkbookMetadataRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorkbookMetadataWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsWorkbook workbook = result.Workbook;
            Assert.Equal(25, workbook.MetadataRecords.Count);
            Assert.Equal(6, workbook.FutureMetadataRecords.Count);
            Assert.Equal((ushort)1200, workbook.CodePage.GetValueOrDefault());
            Assert.Equal("ThisWorkbook", workbook.CodeName);
            Assert.Equal((ushort)1200, workbook.UserInterfaceCodePage.GetValueOrDefault());
            Assert.Equal("OfficeIMO", workbook.LastWriteUserName);
            Assert.True(workbook.WindowsLocked.GetValueOrDefault());
            Assert.NotNull(workbook.SheetTabIds);
            Assert.Equal(new ushort[] { 1, 2 }, workbook.SheetTabIds!.TabIds);
            LegacyXlsWorkbookWindow window = Assert.Single(workbook.Windows);
            Assert.Equal((short)10, window.HorizontalPositionTwips);
            Assert.Equal((short)20, window.VerticalPositionTwips);
            Assert.Equal((short)5000, window.WidthTwips);
            Assert.Equal((short)4000, window.HeightTwips);
            Assert.False(window.Hidden);
            Assert.False(window.Minimized);
            Assert.False(window.VeryHidden);
            Assert.True(window.HorizontalScrollBarVisible);
            Assert.True(window.VerticalScrollBarVisible);
            Assert.True(window.SheetTabsVisible);
            Assert.True(window.AutoFilterDatesGroupedChronologically);
            Assert.Equal((ushort)0, window.ActiveSheetIndex);
            Assert.Equal((ushort)0, window.FirstVisibleSheetTabIndex);
            Assert.Equal((ushort)1, window.SelectedSheetTabCount);
            Assert.Equal((ushort)600, window.SheetTabRatio);
            WorkbookView projectedWindow = result.Document.WorkbookRoot
                .GetFirstChild<BookViews>()!
                .GetFirstChild<WorkbookView>()!;
            Assert.Equal(10, projectedWindow.XWindow!.Value);
            Assert.Equal(20, projectedWindow.YWindow!.Value);
            Assert.Equal(5000U, projectedWindow.WindowWidth!.Value);
            Assert.Equal(4000U, projectedWindow.WindowHeight!.Value);
            Assert.Equal(VisibilityValues.Visible, projectedWindow.Visibility!.Value);
            Assert.False(projectedWindow.Minimized!.Value);
            Assert.True(projectedWindow.ShowHorizontalScroll!.Value);
            Assert.True(projectedWindow.ShowVerticalScroll!.Value);
            Assert.True(projectedWindow.ShowSheetTabs!.Value);
            Assert.Equal(600U, projectedWindow.TabRatio!.Value);
            WorkbookProperties projectedWorkbookProperties = result.Document.WorkbookRoot.GetFirstChild<WorkbookProperties>()!;
            Assert.Equal("ThisWorkbook", projectedWorkbookProperties.CodeName!.Value);
            Assert.True(projectedWorkbookProperties.BackupFile!.Value);
            Assert.False(projectedWorkbookProperties.SaveExternalLinkValues!.Value);
            Assert.Equal(ObjectDisplayValues.None, projectedWorkbookProperties.ShowObjects!.Value);
            Assert.False(projectedWorkbookProperties.ShowBorderUnselectedTables!.Value);
            WorksheetPart projectedWorksheetPart = result.Document.WorkbookPartRoot.WorksheetParts.Single();
            SheetProperties projectedSheetProperties = projectedWorksheetPart.Worksheet.GetFirstChild<SheetProperties>()!;
            Assert.Equal("MetadataSheet", projectedSheetProperties.CodeName!.Value);
            Assert.True(workbook.SaveBackup.GetValueOrDefault());
            Assert.Equal((ushort)2, workbook.HiddenObjectsMode.GetValueOrDefault());
            Assert.True(workbook.DoNotSaveExternalLinkValues.GetValueOrDefault());
            Assert.True(workbook.HasEnvelope.GetValueOrDefault());
            Assert.True(workbook.EnvelopeVisible.GetValueOrDefault());
            Assert.True(workbook.EnvelopeInitialized.GetValueOrDefault());
            Assert.Equal((byte)2, workbook.ExternalLinkUpdateMode.GetValueOrDefault());
            Assert.True(workbook.HideBordersForInactiveTables.GetValueOrDefault());
            Assert.Equal((ushort)2, workbook.PrintSize.GetValueOrDefault());
            Assert.True(workbook.RevisionTrackingLocked.GetValueOrDefault());
            Assert.Equal((ushort)0x1234, workbook.RevisionTrackingPasswordHash.GetValueOrDefault());
            Assert.True(workbook.HasVbaProjectMarker);
            Assert.True(workbook.HasVbaProjectWithoutMacros);
            Assert.True(workbook.UsesNaturalLanguageFormulas.GetValueOrDefault());
            Assert.NotNull(workbook.Country);
            Assert.Equal((ushort)48, workbook.Country!.DefaultCountryCode);
            Assert.Equal((ushort)1, workbook.Country.SystemCountryCode);
            LegacyXlsWorkbookFutureMetadataRecord recalcId = Assert.Single(workbook.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.RecalculationIdentifier);
            Assert.Equal((ushort)0x01c0, recalcId.RecordType);
            Assert.Equal(4, recalcId.PayloadLength);
            Assert.Equal("RawPayload", recalcId.HeaderState);
            Assert.Equal(4, recalcId.BodyByteCount);
            LegacyXlsWorkbookFutureMetadataRecord compat12 = Assert.Single(workbook.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.Compatibility12);
            Assert.True(compat12.HasMatchingFutureRecordHeader);
            Assert.Equal((ushort)0x088c, compat12.HeaderRecordType);
            Assert.Equal((ushort)0x0001, compat12.HeaderFlags);
            Assert.Equal(4, compat12.BodyByteCount);
            LegacyXlsWorksheet sheet = Assert.Single(workbook.Worksheets);
            Assert.Equal("MetadataSheet", sheet.CodeName);
            Assert.Equal(1, sheet.MetadataRecords.Count(record => record.Kind == LegacyXlsWorksheetMetadataKind.CodeName));
            Assert.Equal(1, sheet.MetadataRecords.Count(record => record.Kind == LegacyXlsWorksheetMetadataKind.FutureMetadata));
            LegacyXlsSheetFutureMetadataRecord sheetPageLayoutView = Assert.Single(sheet.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.PageLayoutView);
            Assert.True(sheetPageLayoutView.HasMatchingFutureRecordHeader);
            Assert.Equal((ushort)0x088b, sheetPageLayoutView.RecordType);
            Assert.Equal((ushort)0x0005, sheetPageLayoutView.HeaderFlags);
            Assert.Equal(4, sheetPageLayoutView.BodyByteCount);
            Assert.Equal(25, result.ImportReport.WorkbookMetadataRecordCount);
            Assert.Equal(6, result.ImportReport.WorkbookFutureMetadataRecordCount);
            Assert.Equal(2, result.ImportReport.WorksheetMetadataRecordCount);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordCount);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.CodeName]);
            Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.FutureMetadata]);
            Assert.Equal(1, result.ImportReport.WorkbookCodeNameStates["Present"]);
            Assert.Equal(1, result.ImportReport.WorkbookCodeNames["ThisWorkbook"]);
            Assert.Equal(1, result.ImportReport.WorkbookOptionStates["SaveBackup:True"]);
            Assert.Equal(1, result.ImportReport.WorkbookOptionStates["DoNotSaveExternalLinkValues:True"]);
            Assert.Equal(1, result.ImportReport.WorkbookOptionStates["HasEnvelope:True"]);
            Assert.Equal(1, result.ImportReport.WorkbookOptionStates["EnvelopeVisible:True"]);
            Assert.Equal(1, result.ImportReport.WorkbookOptionStates["EnvelopeInitialized:True"]);
            Assert.Equal(1, result.ImportReport.WorkbookOptionStates["ExternalLinkUpdateMode:2"]);
            Assert.Equal(1, result.ImportReport.WorkbookOptionStates["HideBordersForInactiveTables:True"]);
            Assert.Equal(1, result.ImportReport.WorksheetCodeNameStates["Present"]);
            Assert.Equal(1, result.ImportReport.WorksheetCodeNames["MetadataSheet"]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.Backup]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.BookExtension]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.BookOptions]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.Compatibility12]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.PageLayoutView]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.CompressPictures]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.CodePage]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.CodeName]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.Country]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.HeaderFooter]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.HiddenObjects]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.InterfaceCodePage]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.InterfaceEnd]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.MultiThreadedRecalculationSettings]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.NaturalLanguageFormulas]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.PrinterSettings]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.PrintSize]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.RecalculationIdentifier]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.RevisionProtection]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.RevisionProtectionPassword]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.SheetTabIds]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.VbaProjectNoMacrosMarker]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.Window]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.WindowProtection]);
            Assert.Equal(1, result.ImportReport.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.WriteAccess]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByKind["PageLayoutView"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByKind["Compatibility12"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByRecordType["0x088B"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByRecordType["0x088C"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByRecordName["PLV"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByRecordName["Compat12"]);
            Assert.Equal(5, result.ImportReport.WorkbookFutureMetadataRecordsByHeaderState["MatchingFutureHeader"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByHeaderState["RawPayload"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByHeaderRecordType["0x088B"]);
            Assert.Equal(1, result.ImportReport.WorkbookFutureMetadataRecordsByHeaderRecordType["0x088C"]);
            Assert.Equal(2, result.ImportReport.WorkbookFutureMetadataRecordsByHeaderFlags["Flags:0x0001"]);
            Assert.Equal(3, result.ImportReport.WorkbookFutureMetadataRecordsByPayloadLength["Bytes:16"]);
            Assert.Equal(4, result.ImportReport.WorkbookFutureMetadataRecordsByBodyByteCount["Bytes:4"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsByKind["PageLayoutView"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsBySheet["Metadata"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsBySheetAndKind["Metadata|PageLayoutView"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsByRecordType["0x088B"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsByRecordName["PLV"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsByHeaderState["MatchingFutureHeader"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsByHeaderFlags["Flags:0x0005"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsByPayloadLength["Bytes:16"]);
            Assert.Equal(1, result.ImportReport.WorksheetFutureMetadataRecordsByBodyByteCount["Bytes:4"]);
            Assert.DoesNotContain(workbook.UnsupportedFeatures, feature => feature.RecordType is 0x0040 or 0x00da or 0x0863 or 0x01c0 or 0x088b or 0x089a or 0x089b or 0x089c or 0x0042 or 0x01ba or 0x008c or 0x008d or 0x00e1 or 0x00e2 or 0x0033 or 0x01af or 0x01bc or 0x01bd or 0x013d or 0x0160 or 0x003d or 0x0019 or 0x005c);
            Assert.Contains("Workbook metadata records: 25", result.ImportReport.ToMarkdown());
            Assert.Contains("Workbook future metadata records: 6", result.ImportReport.ToMarkdown());
            Assert.Contains("Worksheet future metadata records: 1", result.ImportReport.ToMarkdown());
            Assert.Contains("Workbook Metadata Records By Kind", result.ImportReport.ToMarkdown());
            Assert.Contains("Workbook Future Metadata Records By Header State", result.ImportReport.ToMarkdown());
            Assert.Contains("Worksheet Future Metadata Records By Sheet And Kind", result.ImportReport.ToMarkdown());
            Assert.Contains("Workbook CodeNames", result.ImportReport.ToMarkdown());
            Assert.Contains("Workbook Option States", result.ImportReport.ToMarkdown());
            Assert.Contains("Worksheet CodeNames", result.ImportReport.ToMarkdown());
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4PrintOptions() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4PrintOptionsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            Assert.NotNull(legacySheet.PageSetup);
            Assert.True(legacySheet.PageSetup!.PrintHeadings);
            Assert.True(legacySheet.PageSetup.PrintGridLines);
            Assert.True(legacySheet.PageSetup.HorizontalCentered);
            Assert.False(legacySheet.PageSetup.VerticalCentered);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelSheetPrintOptions options = document.Sheets[0].GetPrintOptions();
            Assert.True(options.PrintHeadings);
            Assert.True(options.PrintGridLines);
            Assert.True(options.HorizontalCentered);
            Assert.False(options.VerticalCentered);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            PrintOptions printOptions = Assert.Single(worksheetPart.Worksheet.Elements<PrintOptions>());
            Assert.True(printOptions.Headings!.Value);
            Assert.True(printOptions.GridLines!.Value);
            Assert.True(printOptions.GridLinesSet!.Value);
            Assert.True(printOptions.HorizontalCentered!.Value);
            Assert.False(printOptions.VerticalCentered!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4ManualPageBreaks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4PageBreaksWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            LegacyXlsPageBreak rowBreak = Assert.Single(legacySheet.RowPageBreaks);
            Assert.Equal(3, rowBreak.Position);
            Assert.Equal(1, rowBreak.Start);
            Assert.Equal(256, rowBreak.End);
            LegacyXlsPageBreak columnBreak = Assert.Single(legacySheet.ColumnPageBreaks);
            Assert.Equal(2, columnBreak.Position);
            Assert.Equal(1, columnBreak.Start);
            Assert.Equal(21, columnBreak.End);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Equal(new[] { 3 }, document.Sheets[0].GetManualRowPageBreaks());
            Assert.Equal(new[] { 2 }, document.Sheets[0].GetManualColumnPageBreaks());

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            RowBreaks rowBreaks = Assert.Single(worksheetPart.Worksheet.Elements<RowBreaks>());
            Assert.Equal(1U, rowBreaks.Count!.Value);
            Assert.Equal(1U, rowBreaks.ManualBreakCount!.Value);
            Break projectedRowBreak = Assert.Single(rowBreaks.Elements<Break>());
            Assert.Equal(3U, projectedRowBreak.Id!.Value);
            Assert.Equal(0U, projectedRowBreak.Min!.Value);
            Assert.Equal(16383U, projectedRowBreak.Max!.Value);
            Assert.True(projectedRowBreak.ManualPageBreak!.Value);
            ColumnBreaks columnBreaks = Assert.Single(worksheetPart.Worksheet.Elements<ColumnBreaks>());
            Assert.Equal(1U, columnBreaks.Count!.Value);
            Assert.Equal(1U, columnBreaks.ManualBreakCount!.Value);
            Break projectedColumnBreak = Assert.Single(columnBreaks.Elements<Break>());
            Assert.Equal(2U, projectedColumnBreak.Id!.Value);
            Assert.Equal(0U, projectedColumnBreak.Min!.Value);
            Assert.Equal(1048575U, projectedColumnBreak.Max!.Value);
            Assert.True(projectedColumnBreak.ManualPageBreak!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsSingleCellPageBreakSpans() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4SingleCellPageBreaksWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            LegacyXlsPageBreak rowBreak = Assert.Single(legacySheet.RowPageBreaks);
            Assert.Equal(3, rowBreak.Position);
            Assert.Equal(6, rowBreak.Start);
            Assert.Equal(6, rowBreak.End);
            LegacyXlsPageBreak columnBreak = Assert.Single(legacySheet.ColumnPageBreaks);
            Assert.Equal(2, columnBreak.Position);
            Assert.Equal(8, columnBreak.Start);
            Assert.Equal(8, columnBreak.End);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType is 0x001a or 0x001b);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4ZoomScale() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4ZoomScaleWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            Assert.Equal(150U, legacySheet.ZoomScale);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Equal(150U, document.Sheets[0].GetZoomScale());

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            SheetView sheetView = Assert.Single(worksheetPart.Worksheet.Elements<SheetViews>().Single().Elements<SheetView>());
            Assert.Equal(150U, sheetView.ZoomScale!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4HeaderFooter() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4HeaderFooterWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            Assert.NotNull(legacySheet.PageSetup);
            Assert.Equal("&LLeft &P&L&E Again&CQuarterly&RConfidential", legacySheet.PageSetup!.HeaderText);
            Assert.Equal("&CPage &P of &N", legacySheet.PageSetup.FooterText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelSheet.HeaderFooterSnapshot snapshot = document.Sheets[0].GetHeaderFooter();
            Assert.Equal("Left &P&E Again", snapshot.HeaderLeft);
            Assert.Equal("Quarterly", snapshot.HeaderCenter);
            Assert.Equal("Confidential", snapshot.HeaderRight);
            Assert.Equal("Page &P of &N", snapshot.FooterCenter);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            HeaderFooter headerFooter = Assert.Single(worksheetPart.Worksheet.Elements<HeaderFooter>());
            Assert.Equal("&LLeft &P&E Again&CQuarterly&RConfidential", headerFooter.OddHeader!.Text);
            Assert.Equal("&CPage &P of &N", headerFooter.OddFooter!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4DefinedNames() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4DefinedNamesWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(5, legacy.DefinedNames.Count);
            Assert.Contains(legacy.DefinedNames, name => name.Name == "DataRange" && name.Reference == "'Names'!$A$1:$B$2" && name.LocalSheetIndex == null);
            Assert.Contains(legacy.DefinedNames, name => name.Name == "HiddenCell" && name.Reference == "'Names'!$C$3" && name.LocalSheetIndex == 0 && name.Hidden);
            Assert.Contains(legacy.DefinedNames, name => name.Name == "_xlnm.Print_Area" && name.Reference == "'Names'!$A$1:$B$4" && name.LocalSheetIndex == 0 && name.BuiltIn);
            Assert.Contains(legacy.DefinedNames, name => name.Name == "_xlnm.Print_Titles" && name.Reference == "'Names'!$1:$2,'Names'!$A:$B" && name.LocalSheetIndex == 0 && name.BuiltIn);
            Assert.Contains(legacy.DefinedNames, name => name.Name == "_FilterDatabase" && name.Reference == "'Names'!$A$1:$B$3" && name.LocalSheetIndex == 0 && name.Hidden && name.BuiltIn);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelSheet sheet = document.Sheets[0];
            Assert.Equal("'Names'!$A$1:$B$2", document.GetNamedRange("DataRange"));
            Assert.Equal("$C$3", sheet.GetNamedRange("HiddenCell"));
            Assert.Equal("$A$1:$B$3", sheet.GetNamedRange("_FilterDatabase"));
            Assert.Equal("$A$1:$B$4", sheet.GetPrintArea());
            ExcelPrintTitles printTitles = sheet.GetPrintTitles();
            Assert.Equal(1, printTitles.FirstRow);
            Assert.Equal(2, printTitles.LastRow);
            Assert.Equal(1, printTitles.FirstColumn);
            Assert.Equal(2, printTitles.LastColumn);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            List<DefinedName> definedNames = spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
            Assert.Contains(definedNames, name => name.Name == "DataRange" && name.LocalSheetId == null && name.Text == "'Names'!$A$1:$B$2");
            Assert.Contains(definedNames, name => name.Name == "HiddenCell" && name.LocalSheetId?.Value == 0U && name.Hidden?.Value == true && name.Text == "'Names'!$C$3");
            Assert.Contains(definedNames, name => name.Name == "_xlnm.Print_Area" && name.LocalSheetId?.Value == 0U && name.Text == "'Names'!$A$1:$B$4");
            Assert.Contains(definedNames, name => name.Name == "_xlnm.Print_Titles" && name.LocalSheetId?.Value == 0U && name.Text == "'Names'!$1:$2,'Names'!$A:$B");
            Assert.Contains(definedNames, name => name.Name == "_FilterDatabase" && name.LocalSheetId?.Value == 0U && name.Hidden?.Value == true && name.Text == "'Names'!$A$1:$B$3");
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.Single();
            AutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<AutoFilter>());
            Assert.Equal("A1:B3", autoFilter.Reference!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4AutoFilterCriteria() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AutoFilterCriteriaWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            Assert.Equal((ushort)2, legacySheet.AutoFilterDropDownCount);
            Assert.Equal(2, legacySheet.AutoFilterCriteria.Count);
            LegacyXlsAutoFilterCriteria statusCriteria = legacySheet.AutoFilterCriteria[0];
            Assert.Equal(0U, statusCriteria.ColumnId);
            LegacyXlsAutoFilterCondition statusCondition = Assert.Single(statusCriteria.Conditions);
            Assert.Equal(LegacyXlsAutoFilterOperator.Equal, statusCondition.Operator);
            Assert.Equal("Open", statusCondition.Value);
            Assert.Equal(LegacyXlsAutoFilterValueKind.Text, statusCondition.ValueKind);
            LegacyXlsAutoFilterCriteria amountCriteria = legacySheet.AutoFilterCriteria[1];
            Assert.Equal(1U, amountCriteria.ColumnId);
            LegacyXlsAutoFilterCondition amountCondition = Assert.Single(amountCriteria.Conditions);
            Assert.Equal(LegacyXlsAutoFilterOperator.GreaterThanOrEqual, amountCondition.Operator);
            Assert.Equal("10", amountCondition.Value);
            Assert.Equal(LegacyXlsAutoFilterValueKind.Number, amountCondition.ValueKind);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(2, result.ImportReport.AutoFilterCriteriaCount);

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<AutoFilter>());
            Assert.Equal("A1:B4", autoFilter.Reference!.Value);
            List<FilterColumn> filterColumns = autoFilter.Elements<FilterColumn>().OrderBy(column => column.ColumnId?.Value ?? 0U).ToList();
            Assert.Equal(2, filterColumns.Count);

            FilterColumn statusColumn = filterColumns[0];
            Assert.Equal(0U, statusColumn.ColumnId!.Value);
            Filter statusFilter = Assert.Single(statusColumn.GetFirstChild<Filters>()!.Elements<Filter>());
            Assert.Equal("Open", statusFilter.Val!.Value);

            FilterColumn amountColumn = filterColumns[1];
            Assert.Equal(1U, amountColumn.ColumnId!.Value);
            CustomFilter amountFilter = Assert.Single(amountColumn.GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
            Assert.Equal(FilterOperatorValues.GreaterThanOrEqual, amountFilter.Operator!.Value);
            Assert.Equal("10", amountFilter.Val!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4AutoFilterAndCriteria() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AutoFilterAndCriteriaWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            LegacyXlsAutoFilterCriteria criteria = Assert.Single(legacySheet.AutoFilterCriteria);
            Assert.Equal(0U, criteria.ColumnId);
            Assert.Equal(LegacyXlsAutoFilterKind.Custom, criteria.Kind);
            Assert.True(criteria.MatchAll);
            Assert.Equal(LegacyXlsAutoFilterJoinOperator.And, criteria.JoinOperator);
            Assert.Equal(2, criteria.Conditions.Count);
            Assert.Equal(LegacyXlsAutoFilterOperator.GreaterThanOrEqual, criteria.Conditions[0].Operator);
            Assert.Equal("10", criteria.Conditions[0].Value);
            Assert.Equal(LegacyXlsAutoFilterOperator.LessThanOrEqual, criteria.Conditions[1].Operator);
            Assert.Equal("20", criteria.Conditions[1].Value);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaCount);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByJoinOperator["And"]);
            Assert.Empty(result.Document.ValidateOpenXml());

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<AutoFilter>());
            Assert.Equal("A1:A5", autoFilter.Reference!.Value);
            FilterColumn filterColumn = Assert.Single(autoFilter.Elements<FilterColumn>());
            CustomFilters filters = filterColumn.GetFirstChild<CustomFilters>()!;
            Assert.True(filters.And!.Value);
            List<CustomFilter> customFilters = filters.Elements<CustomFilter>().ToList();
            Assert.Equal(2, customFilters.Count);
            Assert.Equal(FilterOperatorValues.GreaterThanOrEqual, customFilters[0].Operator!.Value);
            Assert.Equal("10", customFilters[0].Val!.Value);
            Assert.Equal(FilterOperatorValues.LessThanOrEqual, customFilters[1].Operator!.Value);
            Assert.Equal("20", customFilters[1].Val!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4AutoFilterWildcardTextCriteria() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AutoFilterWildcardTextWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            LegacyXlsAutoFilterCriteria criteria = Assert.Single(legacySheet.AutoFilterCriteria);
            LegacyXlsAutoFilterCondition condition = Assert.Single(criteria.Conditions);
            Assert.Equal(0U, criteria.ColumnId);
            Assert.Equal(LegacyXlsAutoFilterKind.Custom, criteria.Kind);
            Assert.Equal(LegacyXlsAutoFilterOperator.Equal, condition.Operator);
            Assert.Equal("*Open*", condition.Value);
            Assert.Equal(LegacyXlsAutoFilterValueKind.Text, condition.ValueKind);
            Assert.Equal(LegacyXlsAutoFilterTextPatternKind.Contains, condition.TextPatternKind);
            Assert.True(condition.HasTextWildcardPattern);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaCount);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByTextPattern["Contains"]);
            Assert.Empty(result.Document.ValidateOpenXml());

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<AutoFilter>());
            Assert.Equal("A1:A4", autoFilter.Reference!.Value);
            FilterColumn filterColumn = Assert.Single(autoFilter.Elements<FilterColumn>());
            Assert.Equal(0U, filterColumn.ColumnId!.Value);
            Assert.Empty(filterColumn.Elements<Filters>());
            CustomFilter customFilter = Assert.Single(filterColumn.GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
            Assert.Equal(FilterOperatorValues.Equal, customFilter.Operator!.Value);
            Assert.Equal("*Open*", customFilter.Val!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4AutoFilterTop10Criteria() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AutoFilterTop10WorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            Assert.Equal((ushort)1, legacySheet.AutoFilterDropDownCount);
            LegacyXlsAutoFilterCriteria criteria = Assert.Single(legacySheet.AutoFilterCriteria);
            Assert.Equal(0U, criteria.ColumnId);
            Assert.Equal(LegacyXlsAutoFilterKind.Top10, criteria.Kind);
            Assert.True(criteria.IsTop10);
            Assert.Equal((ushort)10, criteria.Top10Value);
            Assert.True(criteria.Top10IsTop);
            Assert.False(criteria.Top10IsPercent);
            Assert.Empty(criteria.Conditions);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaCount);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByKind["Top10"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByColumn["Column:0"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByConditionCount["Conditions:0"]);
            Assert.Equal(1, result.ImportReport.AutoFilterTop10Kinds["TopItems"]);
            Assert.Equal(1, result.ImportReport.AutoFilterTop10Values["TopItems:10"]);
            Assert.Equal(1, result.ImportReport.AutoFilterTop10Directions["Top"]);
            Assert.Equal(1, result.ImportReport.AutoFilterTop10Units["Items"]);
            Assert.Empty(result.Document.ValidateOpenXml());

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<AutoFilter>());
            Assert.Equal("A1:A5", autoFilter.Reference!.Value);
            FilterColumn filterColumn = Assert.Single(autoFilter.Elements<FilterColumn>());
            Assert.Equal(0U, filterColumn.ColumnId!.Value);
            Top10 top10 = Assert.Single(filterColumn.Elements<Top10>());
            Assert.True(top10.Top!.Value);
            Assert.False(top10.Percent!.Value);
            Assert.Equal(10d, top10.Val!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4AutoFilterBlankNonBlankCriteria() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AutoFilterBlankNonBlankWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet legacySheet = Assert.Single(legacy.Worksheets);
            Assert.Equal((ushort)2, legacySheet.AutoFilterDropDownCount);
            Assert.Equal(2, legacySheet.AutoFilterCriteria.Count);

            LegacyXlsAutoFilterCriteria blankCriteria = legacySheet.AutoFilterCriteria[0];
            Assert.Equal(0U, blankCriteria.ColumnId);
            Assert.Equal(LegacyXlsAutoFilterKind.Blanks, blankCriteria.Kind);
            LegacyXlsAutoFilterCondition blankCondition = Assert.Single(blankCriteria.Conditions);
            Assert.Equal(LegacyXlsAutoFilterOperator.Equal, blankCondition.Operator);
            Assert.Equal(string.Empty, blankCondition.Value);
            Assert.Equal(LegacyXlsAutoFilterValueKind.Blank, blankCondition.ValueKind);

            LegacyXlsAutoFilterCriteria nonBlankCriteria = legacySheet.AutoFilterCriteria[1];
            Assert.Equal(1U, nonBlankCriteria.ColumnId);
            Assert.Equal(LegacyXlsAutoFilterKind.NonBlanks, nonBlankCriteria.Kind);
            LegacyXlsAutoFilterCondition nonBlankCondition = Assert.Single(nonBlankCriteria.Conditions);
            Assert.Equal(LegacyXlsAutoFilterOperator.NotEqual, nonBlankCondition.Operator);
            Assert.Equal(string.Empty, nonBlankCondition.Value);
            Assert.Equal(LegacyXlsAutoFilterValueKind.NonBlank, nonBlankCondition.ValueKind);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(2, result.ImportReport.AutoFilterCriteriaCount);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByKind["Blanks"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByKind["NonBlanks"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByOperator["Equal"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByOperator["NotEqual"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByValueKind["Blank"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByValueKind["NonBlank"]);
            Assert.Empty(result.Document.ValidateOpenXml());

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<AutoFilter>());
            Assert.Equal("A1:B5", autoFilter.Reference!.Value);
            List<FilterColumn> filterColumns = autoFilter.Elements<FilterColumn>().OrderBy(column => column.ColumnId?.Value ?? 0U).ToList();
            Assert.Equal(2, filterColumns.Count);

            FilterColumn blankColumn = filterColumns[0];
            Assert.Equal(0U, blankColumn.ColumnId!.Value);
            Filters blankFilters = Assert.Single(blankColumn.Elements<Filters>());
            Assert.True(blankFilters.Blank!.Value);
            Assert.Empty(blankFilters.Elements<Filter>());

            FilterColumn nonBlankColumn = filterColumns[1];
            Assert.Equal(1U, nonBlankColumn.ColumnId!.Value);
            CustomFilter customFilter = Assert.Single(nonBlankColumn.GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
            Assert.Equal(FilterOperatorValues.NotEqual, customFilter.Operator!.Value);
            Assert.Equal(" ", customFilter.Val!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ReportsUnsupportedBoundSheetTypes() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5UnsupportedSheetTypesWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Data", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Imported"));
            Assert.Contains(legacy.UnsupportedSheets, sheet => sheet.Name == "Macro1" && sheet.Kind == LegacyXlsUnsupportedSheetKind.MacroSheet && sheet.SheetType == 0x01);
            Assert.Contains(legacy.UnsupportedSheets, sheet => sheet.Name == "Module1" && sheet.Kind == LegacyXlsUnsupportedSheetKind.VbaModuleSheet && sheet.SheetType == 0x06);
            LegacyXlsChartSheet chartSheet = Assert.Single(legacy.ChartSheets);
            Assert.Equal("Chart1", chartSheet.Name);
            Assert.Equal(0x02, chartSheet.SheetType);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.MacroSheet && feature.SheetName == "Macro1" && feature.RecordType == 0x0085);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.VbaModuleSheet && feature.SheetName == "Module1" && feature.RecordType == 0x0085);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.DetailCode == "Sheet:MacroSheet");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.DetailCode == "Sheet:VbaModuleSheet");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ChartSheet);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.DetailCode == "Sheet:ChartSheet");
            Assert.DoesNotContain(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.ChartSheet);
            Assert.DoesNotContain(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.MacroSheet);
            Assert.DoesNotContain(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.VbaModuleSheet);
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED" && d.SheetName == "Macro1" && d.RecordType == 0x0085);
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-VBA-MODULE-SHEET-UNSUPPORTED" && d.SheetName == "Module1" && d.RecordType == 0x0085);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-CHART-SHEET-UNSUPPORTED");
            Assert.DoesNotContain(legacy.Diagnostics, d => d.DetailCode == "Sheet:ChartSheet");
        }

        [Fact]
        public void LegacyXls_Load_ReportsUnsupportedDialogSheets() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5DialogSheetWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Data", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Imported"));
            LegacyXlsDefinedName localName = Assert.Single(legacy.DefinedNames, name => name.Name == "DataLocal");
            Assert.Equal(0, localName.LocalSheetIndex);
            LegacyXlsUnsupportedSheet unsupportedSheet = Assert.Single(legacy.UnsupportedSheets, sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.DialogSheet);
            Assert.Equal("Dialog1", unsupportedSheet.Name);
            Assert.Equal(0x00, unsupportedSheet.SheetType);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DialogSheet && feature.SheetName == "Dialog1" && feature.RecordType == 0x0081);
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-DIALOG-SHEET-UNSUPPORTED" && d.SheetName == "Dialog1" && d.RecordType == 0x0081);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelSheet projectedSheet = Assert.Single(document.Sheets);
            Assert.Equal("Data", projectedSheet.Name);
            Assert.True(projectedSheet.TryGetCellText(1, 1, out string? text));
            Assert.Equal("Imported", text);
        }

        [Fact]
        public void LegacyXls_Load_PreservesZeroFitToHeightAsUnlimitedPagesTall() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5PrintPageSetupUnlimitedHeightWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Equal((ushort)0, Assert.Single(result.Workbook.Worksheets).PageSetup!.FitToHeight);
            ExcelSheetPageSetup setup = result.Document.Sheets[0].GetPageSetup();
            Assert.Equal(1U, setup.FitToWidth);
            Assert.Equal(0U, setup.FitToHeight);
        }

        [Fact]
        public void LegacyXls_Load_IgnoresSetupOrientationWhenPrinterSettingsAreAbsent() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5PrintPageSetupNoPrinterSettingsOrientationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Null(Assert.Single(result.Workbook.Worksheets).PageSetup!.Landscape);
            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            PageSetup openXmlSetup = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<PageSetup>()!;
            Assert.Null(openXmlSetup.Orientation);
        }

        [Fact]
        public void LegacyXls_Load_ReportsShortCellFormatWithoutReadingPastPayload() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ShortCellFormatWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Contains(legacy.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-XF-SHORT");
            Assert.Empty(legacy.CellFormats);
            Assert.Contains(Assert.Single(legacy.Worksheets).Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Imported"));
        }

        [Fact]
        public void LegacyXls_Load_PreservesExternalReferenceMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ExternalReferencesWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                PreserveExternalWorkbookLinks = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Data", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Imported"));
            LegacyXlsExternalReference reference = Assert.Single(legacy.ExternalReferences);
            Assert.Equal(LegacyXlsExternalReferenceKind.ExternalWorkbook, reference.Kind);
            Assert.Equal("C:\\Data\\Budget.xls", reference.Target);
            Assert.Equal(2, reference.SheetCount);
            Assert.Equal(new[] { "Jan", "Feb" }, reference.SheetNames);
            Assert.Equal(2, reference.SheetNameCount);
            Assert.Equal(3, reference.ExternalNameCount);
            Assert.Equal(0, reference.CachedCellCacheCount);
            Assert.Equal(0, reference.CachedCellCount);
            LegacyXlsDataConsolidationReference consolidationReference = Assert.Single(legacy.DataConsolidationReferences);
            Assert.Equal(LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath, consolidationReference.SourceKind);
            Assert.Equal("C:\\Data\\Budget.xls", consolidationReference.Source);
            Assert.Equal("B2:D4", consolidationReference.CellRange);
            Assert.Equal(2, consolidationReference.FirstRow);
            Assert.Equal(4, consolidationReference.LastRow);
            Assert.Equal(2, consolidationReference.FirstColumn);
            Assert.Equal(4, consolidationReference.LastColumn);
            Assert.Equal(3, consolidationReference.RowSpan);
            Assert.Equal(3, consolidationReference.ColumnSpan);
            Assert.Equal((byte)0x01, consolidationReference.SourcePrefix);
            Assert.Equal(0, consolidationReference.UnusedByteCount);
            Assert.True(legacy.HasRefreshAllMarker);
            Assert.Equal(1, legacy.MetadataRecords.Count(record => record.Kind == LegacyXlsWorkbookMetadataKind.RefreshAll));
            Assert.Equal(3, reference.ExternalNames.Count);
            LegacyXlsExternalName externalName = Assert.Single(reference.ExternalNames, name => name.Name == "TaxRate");
            Assert.Null(externalName.LocalSheetIndex);
            Assert.False(externalName.BuiltIn);
            Assert.Equal(LegacyXlsExternalNameBodyKind.ExternalDefinedName, externalName.BodyKind);
            Assert.False(externalName.WantsAdvise);
            Assert.False(externalName.WantsPicture);
            Assert.False(externalName.Ole);
            Assert.False(externalName.OleLink);
            Assert.Equal(0, externalName.CachedClipboardFormat);
            Assert.Equal("TextOrExternalName", externalName.CachedClipboardFormatName);
            Assert.False(externalName.Icon);
            LegacyXlsExternalName localExternalName = Assert.Single(reference.ExternalNames, name => name.Name == "FebTaxRate");
            Assert.Equal(1, localExternalName.LocalSheetIndex);
            Assert.False(localExternalName.BuiltIn);
            Assert.Equal(LegacyXlsExternalNameBodyKind.ExternalDefinedName, localExternalName.BodyKind);
            LegacyXlsExternalName builtInExternalName = Assert.Single(reference.ExternalNames, name => name.Name == "_xlnm.Print_Area");
            Assert.Null(builtInExternalName.LocalSheetIndex);
            Assert.True(builtInExternalName.BuiltIn);
            Assert.Equal(LegacyXlsExternalNameBodyKind.ExternalDefinedName, builtInExternalName.BodyKind);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.ExternalReferenceCount);
            Assert.Equal(2, report.ExternalSheetNameCount);
            Assert.Equal(3, report.ExternalNameCount);
            Assert.Equal(1, report.ExternalReferencesByKind[LegacyXlsExternalReferenceKind.ExternalWorkbook]);
            Assert.Equal(1, report.ExternalReferencesByTarget["C:\\Data\\Budget.xls"]);
            Assert.Equal(1, report.ExternalReferencesByShape["ExternalWorkbook|Sheets:2|Names:3|Caches:0|CachedCells:0"]);
            Assert.Equal(1, report.ExternalReferenceWorkbookStates["ExternalWorkbooks:Present|Self:Missing|AddIns:Missing|DdeOle:Missing|SheetTables:Present|ExternalNames:Present|CellCaches:Missing|CachedCells:Missing|CacheLinks:None"]);
            Assert.Equal(1, report.ExternalReferencesBySheetNameCount["Sheets:2"]);
            Assert.Equal(1, report.ExternalReferencesByExternalNameCount["Names:3"]);
            Assert.Equal(1, report.ExternalReferencesByCacheCount["Caches:0"]);
            Assert.Equal(1, report.ExternalReferencesByCachedCellCount["CachedCells:0"]);
            Assert.Equal(1, report.DataConsolidationReferenceCount);
            Assert.Equal(1, report.DataConsolidationReferencesBySourceKind["ExternalVirtualPath"]);
            Assert.Equal(1, report.DataConsolidationReferencesBySourcePrefix["Prefix:0x01"]);
            Assert.Equal(1, report.DataConsolidationReferencesBySource["C:\\Data\\Budget.xls"]);
            Assert.Equal(1, report.DataConsolidationReferencesByRange["B2:D4"]);
            Assert.Equal(1, report.DataConsolidationReferencesByShape["Rows:3;Columns:3"]);
            Assert.Equal(1, report.DataConsolidationReferencesBySourceAndRange["C:\\Data\\Budget.xls|B2:D4"]);
            Assert.Equal(1, report.DataConsolidationReferencesByUnusedByteCount["UnusedBytes:0"]);
            Assert.Equal(2, report.ExternalSheetNamesByReferenceKind[LegacyXlsExternalReferenceKind.ExternalWorkbook]);
            Assert.Equal(3, report.ExternalNamesByReferenceKind[LegacyXlsExternalReferenceKind.ExternalWorkbook]);
            Assert.Equal(1, report.ExternalNamesByName["TaxRate"]);
            Assert.Equal(1, report.ExternalNamesByName["FebTaxRate"]);
            Assert.Equal(1, report.ExternalNamesByName["_xlnm.Print_Area"]);
            Assert.Equal(2, report.ExternalNamesByScope["Workbook"]);
            Assert.Equal(1, report.ExternalNamesByScope["SheetLocal"]);
            Assert.Equal(2, report.ExternalNamesByBuiltInState["Custom"]);
            Assert.Equal(1, report.ExternalNamesByBuiltInState["BuiltIn"]);
            Assert.Equal(3, report.ExternalNamesByBodyKind["ExternalDefinedName"]);
            Assert.Equal(3, report.ExternalNamesByCachedClipboardFormat["TextOrExternalName:0"]);
            Assert.Equal(3, report.ExternalNamesByAdviseState["Missing"]);
            Assert.Equal(3, report.ExternalNamesByPictureState["Missing"]);
            Assert.Equal(3, report.ExternalNamesByOleState["Missing"]);
            Assert.Equal(3, report.ExternalNamesByOleLinkState["Missing"]);
            Assert.Equal(3, report.ExternalNamesByIconState["Missing"]);
            Assert.Equal(2, report.ExternalNamesByFlagShape["Body:ExternalDefinedName|BuiltIn:Missing|Advise:Missing|Picture:Missing|Ole:Missing|OleLink:Missing|Icon:Missing"]);
            Assert.Equal(1, report.ExternalNamesByFlagShape["Body:ExternalDefinedName|BuiltIn:Present|Advise:Missing|Picture:Missing|Ole:Missing|OleLink:Missing|Icon:Missing"]);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.RecordType == 0x01ae && feature.Description.Contains("C:\\Data\\Budget.xls", StringComparison.Ordinal));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.DetailCode == "ExternalReference:ExternalWorkbook");
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED" && d.RecordType == 0x01ae);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.DetailCode == "ExternalReference:ExternalWorkbook");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.DetailCode == "ExternalReference:DConRef");
            Assert.DoesNotContain(legacy.Diagnostics, d => d.DetailCode == "ExternalReference:DConRef");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.RecordType == 0x0023);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.RecordType == 0x01b7);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED" && d.RecordType == 0x0023);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED" && d.RecordType == 0x01b7);
        }

        [Fact]
        public void LegacyXls_Load_ReportsExternalNameFlagStates() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5DdeOleExternalNameFlagsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsExternalReference reference = Assert.Single(legacy.ExternalReferences);
            Assert.Equal(LegacyXlsExternalReferenceKind.DdeOrOle, reference.Kind);
            Assert.Equal("Excel|Topic", reference.Target);
            Assert.Equal(3, reference.ExternalNameCount);

            LegacyXlsExternalName ddePicture = Assert.Single(reference.ExternalNames, name => name.Name == "DdePicture");
            Assert.True(ddePicture.WantsAdvise);
            Assert.True(ddePicture.WantsPicture);
            Assert.False(ddePicture.Ole);
            Assert.False(ddePicture.OleLink);
            Assert.Equal(9, ddePicture.CachedClipboardFormat);
            Assert.Equal("Bitmap", ddePicture.CachedClipboardFormatName);
            Assert.True(ddePicture.Icon);
            Assert.Equal(LegacyXlsExternalNameBodyKind.OleDdeLink, ddePicture.BodyKind);

            LegacyXlsExternalName oleNoOper = Assert.Single(reference.ExternalNames, name => name.Name == "OleNoOper");
            Assert.True(oleNoOper.Ole);
            Assert.False(oleNoOper.OleLink);
            Assert.Equal(LegacyXlsExternalNameBodyKind.DdeLinkNoOper, oleNoOper.BodyKind);

            LegacyXlsExternalName oleData = Assert.Single(reference.ExternalNames, name => name.Name == "OleData");
            Assert.False(oleData.Ole);
            Assert.True(oleData.OleLink);
            Assert.Equal(LegacyXlsExternalNameBodyKind.OleDataItem, oleData.BodyKind);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.ExternalReferencesByKind[LegacyXlsExternalReferenceKind.DdeOrOle]);
            Assert.Equal(1, report.ExternalReferenceWorkbookStates["ExternalWorkbooks:Missing|Self:Missing|AddIns:Missing|DdeOle:Present|SheetTables:Missing|ExternalNames:Present|CellCaches:Missing|CachedCells:Missing|CacheLinks:None"]);
            Assert.Equal(1, report.ExternalNamesByBodyKind["OleDdeLink"]);
            Assert.Equal(1, report.ExternalNamesByBodyKind["DdeLinkNoOper"]);
            Assert.Equal(1, report.ExternalNamesByBodyKind["OleDataItem"]);
            Assert.Equal(1, report.ExternalNamesByCachedClipboardFormat["Bitmap:9"]);
            Assert.Equal(2, report.ExternalNamesByCachedClipboardFormat["TextOrExternalName:0"]);
            Assert.Equal(1, report.ExternalNamesByAdviseState["Present"]);
            Assert.Equal(2, report.ExternalNamesByAdviseState["Missing"]);
            Assert.Equal(1, report.ExternalNamesByPictureState["Present"]);
            Assert.Equal(2, report.ExternalNamesByPictureState["Missing"]);
            Assert.Equal(1, report.ExternalNamesByOleState["Present"]);
            Assert.Equal(2, report.ExternalNamesByOleState["Missing"]);
            Assert.Equal(1, report.ExternalNamesByOleLinkState["Present"]);
            Assert.Equal(2, report.ExternalNamesByOleLinkState["Missing"]);
            Assert.Equal(1, report.ExternalNamesByIconState["Present"]);
            Assert.Equal(2, report.ExternalNamesByIconState["Missing"]);
            Assert.Equal(1, report.ExternalNamesByFlagShape["Body:OleDdeLink|BuiltIn:Missing|Advise:Present|Picture:Present|Ole:Missing|OleLink:Missing|Icon:Present"]);
            Assert.Equal(1, report.ExternalNamesByFlagShape["Body:DdeLinkNoOper|BuiltIn:Missing|Advise:Missing|Picture:Missing|Ole:Present|OleLink:Missing|Icon:Missing"]);
            Assert.Equal(1, report.ExternalNamesByFlagShape["Body:OleDataItem|BuiltIn:Missing|Advise:Missing|Picture:Missing|Ole:Missing|OleLink:Present|Icon:Missing"]);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.DetailCode == "ExternalReference:DdeOrOle");
        }

        [Fact]
        public void LegacyXls_Load_ImportsWholeNumberDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4DataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.WholeNumber, validation.Type);
            Assert.Equal(LegacyXlsDataValidationOperator.Between, validation.Operator);
            Assert.Equal("18", validation.Formula1);
            Assert.Equal("65", validation.Formula2);
            Assert.True(validation.AllowBlank);
            Assert.True(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);
            Assert.Equal("Age", validation.PromptTitle);
            Assert.Equal("B2:B4", Assert.Single(validation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedValidation = Assert.Single(document.Sheets[0].GetDataValidations("B2:B4"));
            Assert.Equal("whole", projectedValidation.Type);
            Assert.Equal("between", projectedValidation.Operator);
            Assert.Equal("18", projectedValidation.Formula1);
            Assert.Equal("65", projectedValidation.Formula2);
            Assert.Equal("Age", projectedValidation.PromptTitle);
            Assert.Equal("Invalid age", projectedValidation.ErrorTitle);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.Whole, openXmlValidation.Type!.Value);
            Assert.Equal(DataValidationOperatorValues.Between, openXmlValidation.Operator!.Value);
            Assert.Equal("B2:B4", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("18", openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Equal("65", openXmlValidation.GetFirstChild<Formula2>()!.Text);
            Assert.Equal("Age", openXmlValidation.PromptTitle!.Value);
            Assert.Equal("Invalid age", openXmlValidation.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsDecimalDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4DecimalDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.Decimal, validation.Type);
            Assert.Equal(LegacyXlsDataValidationOperator.GreaterThan, validation.Operator);
            Assert.Equal("5.5", validation.Formula1);
            Assert.Null(validation.Formula2);
            Assert.False(validation.AllowBlank);
            Assert.False(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);
            Assert.Equal("C2:C4", Assert.Single(validation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedValidation = Assert.Single(document.Sheets[0].GetDataValidations("C2:C4"));
            Assert.Equal("decimal", projectedValidation.Type);
            Assert.Equal("greaterThan", projectedValidation.Operator);
            Assert.Equal("5.5", projectedValidation.Formula1);
            Assert.Null(projectedValidation.Formula2);
            Assert.Equal("Invalid discount", projectedValidation.ErrorTitle);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.Decimal, openXmlValidation.Type!.Value);
            Assert.Equal(DataValidationOperatorValues.GreaterThan, openXmlValidation.Operator!.Value);
            Assert.Equal("C2:C4", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("5.5", openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlValidation.GetFirstChild<Formula2>());
            Assert.Equal("Invalid discount", openXmlValidation.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsAnyValueDataValidationMessages() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4AnyValueDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsDataValidation validation = Assert.Single(Assert.Single(result.Workbook.Worksheets).DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.None, validation.Type);
            Assert.Equal(string.Empty, validation.Formula1);
            Assert.Equal("A2", Assert.Single(validation.Ranges));
            Assert.True(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.None, openXmlValidation.Type!.Value);
            Assert.Equal("A2", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("Instructions", openXmlValidation.PromptTitle!.Value);
            Assert.Equal("Check value", openXmlValidation.ErrorTitle!.Value);
            Assert.Null(openXmlValidation.GetFirstChild<Formula1>());
            Assert.Null(openXmlValidation.GetFirstChild<Formula2>());
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaBackedDataValidationBounds() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4FormulaBoundDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsDataValidation validation = Assert.Single(Assert.Single(result.Workbook.Worksheets).DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.WholeNumber, validation.Type);
            Assert.Equal("$A$1", validation.Formula1);
            Assert.Equal("$B$1", validation.Formula2);
            Assert.Equal(new[] { "C2:C6", "E2:E6" }, validation.Ranges);

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.Whole, openXmlValidation.Type!.Value);
            Assert.Equal(DataValidationOperatorValues.Between, openXmlValidation.Operator!.Value);
            Assert.Equal("C2:C6 E2:E6", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("$A$1", openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Equal("$B$1", openXmlValidation.GetFirstChild<Formula2>()!.Text);
            Assert.Equal("Out of range", openXmlValidation.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsInlineListDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4ListDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.List, validation.Type);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.InlineList, validation.ListSourceKind);
            Assert.Equal(new[] { "Open", "Closed", "Pending" }, validation.ListItems);
            Assert.Equal("\"Open,Closed,Pending\"", validation.Formula1);
            Assert.Null(validation.Formula2);
            Assert.True(validation.AllowBlank);
            Assert.True(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);
            Assert.Equal(LegacyXlsDataValidationErrorStyle.Warning, validation.ErrorStyle);
            Assert.True(validation.SuppressDropDown);
            Assert.Equal("D2:D5", Assert.Single(validation.Ranges));
            Assert.Equal(1, validation.RangeCount);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DataValidationsByAllowBlankState["AllowBlank"]);
            Assert.Equal(1, report.DataValidationsByInputMessageState["ShowInputMessage"]);
            Assert.Equal(1, report.DataValidationsByErrorMessageState["ShowErrorMessage"]);
            Assert.Equal(1, report.DataValidationsByPromptTextState["Present"]);
            Assert.Equal(1, report.DataValidationsByErrorTextState["Present"]);
            Assert.Equal(1, report.DataValidationsByDropDownState["Suppressed"]);
            Assert.Equal(1, report.DataValidationsByRangeCount["Ranges:1"]);
            Assert.Equal(1, report.DataValidationsByRange["D2:D5"]);
            Assert.Equal(1, report.DataValidationsByFormula1State["Present"]);
            Assert.Equal(1, report.DataValidationsByFormula2State["Missing"]);
            Assert.Equal(1, report.DataValidationsByFormulaPairState["Formula1:Present|Formula2:Missing"]);
            Assert.Equal(1, report.DataValidationListSourcesByKind["InlineList"]);
            Assert.Equal(1, report.DataValidationListSourcesByItemCount["Items:3"]);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedValidation = Assert.Single(document.Sheets[0].GetDataValidations("D2:D5"));
            Assert.Equal("list", projectedValidation.Type);
            Assert.Equal("\"Open,Closed,Pending\"", projectedValidation.Formula1);
            Assert.Equal("Status", projectedValidation.PromptTitle);
            Assert.Equal("Invalid status", projectedValidation.ErrorTitle);
            Assert.Equal("warning", projectedValidation.ErrorStyle);
            Assert.True(projectedValidation.ShowInputMessage);
            Assert.True(projectedValidation.ShowErrorMessage);
            Assert.True(projectedValidation.SuppressDropDown);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.List, openXmlValidation.Type!.Value);
            Assert.Equal("D2:D5", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("\"Open,Closed,Pending\"", openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlValidation.GetFirstChild<Formula2>());
            Assert.Equal("Status", openXmlValidation.PromptTitle!.Value);
            Assert.Equal("Invalid status", openXmlValidation.ErrorTitle!.Value);
            Assert.Equal(DataValidationErrorStyleValues.Warning, openXmlValidation.ErrorStyle!.Value);
            Assert.True(openXmlValidation.ShowInputMessage!.Value);
            Assert.True(openXmlValidation.ShowErrorMessage!.Value);
            Assert.True(openXmlValidation.ShowDropDown!.Value);
        }

        [Fact]
        public void LegacyXls_Load_PreservesHiddenValidationMessageFlags() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5HiddenMessageDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedValidation = Assert.Single(document.Sheets[0].GetDataValidations("A1:A1"));
            Assert.Equal("Hidden prompt", projectedValidation.PromptTitle);
            Assert.Equal("Hidden error", projectedValidation.ErrorTitle);
            Assert.False(projectedValidation.ShowInputMessage);
            Assert.False(projectedValidation.ShowErrorMessage);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            DataValidation openXmlValidation = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<DataValidation>());
            Assert.False(openXmlValidation.ShowInputMessage!.Value);
            Assert.False(openXmlValidation.ShowErrorMessage!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsRangeBackedListDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4RangeListDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Open"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 1 && Equals(cell.Value, "Closed"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 3 && cell.Column == 1 && Equals(cell.Value, "Pending"));
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.List, validation.Type);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.Range, validation.ListSourceKind);
            Assert.Empty(validation.ListItems);
            Assert.Equal("$A$1:$A$3", validation.Formula1);
            Assert.Equal("A1:A3", validation.ListSourceRange);
            Assert.Null(validation.Formula2);
            Assert.True(validation.AllowBlank);
            Assert.True(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);
            Assert.Equal("H2:H5", Assert.Single(validation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DataValidationListSourcesByKind["Range"]);
            Assert.Equal(1, report.DataValidationListSourcesByItemCount["Items:0"]);
            Assert.Equal(1, report.DataValidationListSourcesByRange["A1:A3"]);
            Assert.Equal(1, report.DataValidationsByFormulaPairState["Formula1:Present|Formula2:Missing"]);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedValidation = Assert.Single(document.Sheets[0].GetDataValidations("H2:H5"));
            Assert.Equal("list", projectedValidation.Type);
            Assert.Equal("=A1:A3", projectedValidation.Formula1);
            Assert.Equal("Status", projectedValidation.PromptTitle);
            Assert.Equal("Invalid status", projectedValidation.ErrorTitle);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.List, openXmlValidation.Type!.Value);
            Assert.Equal("H2:H5", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("=A1:A3", openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlValidation.GetFirstChild<Formula2>());
            Assert.Equal("Status", openXmlValidation.PromptTitle!.Value);
            Assert.Equal("Invalid status", openXmlValidation.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsCrossSheetRangeBackedListDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4CrossSheetRangeListDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(2, legacy.Worksheets.Count);
            LegacyXlsWorksheet optionsSheet = legacy.Worksheets[0];
            Assert.Equal("Options", optionsSheet.Name);
            Assert.Contains(optionsSheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Open"));
            Assert.Contains(optionsSheet.Cells, cell => cell.Row == 2 && cell.Column == 1 && Equals(cell.Value, "Closed"));
            Assert.Contains(optionsSheet.Cells, cell => cell.Row == 3 && cell.Column == 1 && Equals(cell.Value, "Pending"));
            LegacyXlsWorksheet validationSheet = legacy.Worksheets[1];
            Assert.Equal("CrossSheetValidation", validationSheet.Name);
            LegacyXlsDataValidation validation = Assert.Single(validationSheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.List, validation.Type);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.SheetQualifiedRange, validation.ListSourceKind);
            Assert.Empty(validation.ListItems);
            Assert.Equal("'Options'!$A$1:$A$3", validation.Formula1);
            Assert.Equal("A1:A3", validation.ListSourceRange);
            Assert.Equal("Options", validation.ListSourceSheetName);
            Assert.Null(validation.ListSourceName);
            Assert.Equal("H2:H5", Assert.Single(validation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DataValidationListSourcesByKind["SheetQualifiedRange"]);
            Assert.Equal(1, report.DataValidationListSourcesByItemCount["Items:0"]);
            Assert.Equal(1, report.DataValidationListSourcesByRange["A1:A3"]);
            Assert.Equal(1, report.DataValidationListSourcesBySheetName["Options"]);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelSheet projectedValidationSheet = document.Sheets[1];
            ExcelDataValidationInfo projectedValidation = Assert.Single(projectedValidationSheet.GetDataValidations("H2:H5"));
            Assert.Equal("list", projectedValidation.Type);
            Assert.Equal("='Options'!A1:A3", projectedValidation.Formula1);
            Assert.Equal("Status", projectedValidation.PromptTitle);
            Assert.Equal("Invalid status", projectedValidation.ErrorTitle);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.Uri.OriginalString.EndsWith("/sheet2.xml", StringComparison.Ordinal));
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.List, openXmlValidation.Type!.Value);
            Assert.Equal("H2:H5", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("='Options'!A1:A3", openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlValidation.GetFirstChild<Formula2>());
            Assert.Equal("Status", openXmlValidation.PromptTitle!.Value);
            Assert.Equal("Invalid status", openXmlValidation.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsNamedRangeListDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4NamedListDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Contains(legacy.DefinedNames, name => name.Name == "StatusOptions" && name.Reference == "'NamedListValidation'!$A$1:$A$3" && name.LocalSheetIndex == null);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Open"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 1 && Equals(cell.Value, "Closed"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 3 && cell.Column == 1 && Equals(cell.Value, "Pending"));
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.List, validation.Type);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.DefinedName, validation.ListSourceKind);
            Assert.Empty(validation.ListItems);
            Assert.Equal("StatusOptions", validation.Formula1);
            Assert.Null(validation.ListSourceRange);
            Assert.Equal("StatusOptions", validation.ListSourceName);
            Assert.Null(validation.Formula2);
            Assert.True(validation.AllowBlank);
            Assert.True(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);
            Assert.Equal("H2:H5", Assert.Single(validation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DataValidationListSourcesByKind["DefinedName"]);
            Assert.Equal(1, report.DataValidationListSourcesByItemCount["Items:0"]);
            Assert.Equal(1, report.DataValidationListSourcesByName["StatusOptions"]);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedValidation = Assert.Single(document.Sheets[0].GetDataValidations("H2:H5"));
            Assert.Equal("list", projectedValidation.Type);
            Assert.Equal("=StatusOptions", projectedValidation.Formula1);
            Assert.Equal("Status", projectedValidation.PromptTitle);
            Assert.Equal("Invalid status", projectedValidation.ErrorTitle);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            DefinedName openXmlName = Assert.Single(workbookPart.Workbook.DefinedNames!.Elements<DefinedName>(), name => name.Name == "StatusOptions");
            Assert.Equal("'NamedListValidation'!$A$1:$A$3", openXmlName.Text);
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.List, openXmlValidation.Type!.Value);
            Assert.Equal("H2:H5", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("=StatusOptions", openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlValidation.GetFirstChild<Formula2>());
            Assert.Equal("Status", openXmlValidation.PromptTitle!.Value);
            Assert.Equal("Invalid status", openXmlValidation.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsSheetLocalNamedRangeListDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4SheetLocalNamedListDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Contains(legacy.DefinedNames, name => name.Name == "StatusOptions" && name.Reference == "'LocalNamedListValidation'!$A$1:$A$3" && name.LocalSheetIndex == 0);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.List, validation.Type);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.DefinedName, validation.ListSourceKind);
            Assert.Empty(validation.ListItems);
            Assert.Equal("StatusOptions", validation.Formula1);
            Assert.Null(validation.ListSourceRange);
            Assert.Equal("StatusOptions", validation.ListSourceName);
            Assert.Equal("H2:H5", Assert.Single(validation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DataValidationListSourcesByKind["DefinedName"]);
            Assert.Equal(1, report.DataValidationListSourcesByItemCount["Items:0"]);
            Assert.Equal(1, report.DataValidationListSourcesByName["StatusOptions"]);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelSheet projectedSheet = Assert.Single(document.Sheets);
            Assert.Equal("$A$1:$A$3", projectedSheet.GetNamedRange("StatusOptions"));
            ExcelDataValidationInfo projectedValidation = Assert.Single(projectedSheet.GetDataValidations("H2:H5"));
            Assert.Equal("list", projectedValidation.Type);
            Assert.Equal("=StatusOptions", projectedValidation.Formula1);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            DefinedName openXmlName = Assert.Single(workbookPart.Workbook.DefinedNames!.Elements<DefinedName>(), name => name.Name == "StatusOptions");
            Assert.Equal(0U, openXmlName.LocalSheetId!.Value);
            Assert.Equal("'LocalNamedListValidation'!$A$1:$A$3", openXmlName.Text);
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.List, openXmlValidation.Type!.Value);
            Assert.Equal("H2:H5", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal("=StatusOptions", openXmlValidation.GetFirstChild<Formula1>()!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsDateTimeAndTextLengthDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4TypedDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string expectedStartDate = new DateTime(2024, 1, 1).ToOADate().ToString("G15", CultureInfo.InvariantCulture);
            string expectedEndDate = new DateTime(2024, 12, 31).ToOADate().ToString("G15", CultureInfo.InvariantCulture);
            string expectedStartTime = TimeSpan.FromHours(9).TotalDays.ToString("G15", CultureInfo.InvariantCulture);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsDataValidationCollectionRecord collectionRecord = Assert.Single(sheet.DataValidationCollections);
            Assert.Equal("TypedValidation", collectionRecord.SheetName);
            Assert.Equal((ushort)BiffRecordType.DVal, collectionRecord.RecordType);
            Assert.Equal(3U, collectionRecord.DeclaredValidationCount);
            Assert.Equal(3, sheet.DataValidations.Count);
            LegacyXlsDataValidation dateValidation = Assert.Single(sheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.Date);
            Assert.Equal(LegacyXlsDataValidationOperator.Between, dateValidation.Operator);
            Assert.Equal(expectedStartDate, dateValidation.Formula1);
            Assert.Equal(expectedEndDate, dateValidation.Formula2);
            Assert.Equal("E2:E5", Assert.Single(dateValidation.Ranges));
            LegacyXlsDataValidation timeValidation = Assert.Single(sheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.Time);
            Assert.Equal(LegacyXlsDataValidationOperator.GreaterThanOrEqual, timeValidation.Operator);
            Assert.Equal(expectedStartTime, timeValidation.Formula1);
            Assert.Null(timeValidation.Formula2);
            Assert.Equal("F2:F5", Assert.Single(timeValidation.Ranges));
            LegacyXlsDataValidation textLengthValidation = Assert.Single(sheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.TextLength);
            Assert.Equal(LegacyXlsDataValidationOperator.LessThanOrEqual, textLengthValidation.Operator);
            Assert.Equal("12", textLengthValidation.Formula1);
            Assert.Null(textLengthValidation.Formula2);
            Assert.Equal("G2:G5", Assert.Single(textLengthValidation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedDate = Assert.Single(document.Sheets[0].GetDataValidations("E2:E5"));
            Assert.Equal("date", projectedDate.Type);
            Assert.Equal("between", projectedDate.Operator);
            Assert.Equal(expectedStartDate, projectedDate.Formula1);
            Assert.Equal(expectedEndDate, projectedDate.Formula2);
            Assert.Equal("Ship date", projectedDate.PromptTitle);
            Assert.Equal("Invalid date", projectedDate.ErrorTitle);
            ExcelDataValidationInfo projectedTime = Assert.Single(document.Sheets[0].GetDataValidations("F2:F5"));
            Assert.Equal("time", projectedTime.Type);
            Assert.Equal("greaterThanOrEqual", projectedTime.Operator);
            Assert.Equal(expectedStartTime, projectedTime.Formula1);
            Assert.Null(projectedTime.Formula2);
            Assert.Equal("Invalid time", projectedTime.ErrorTitle);
            ExcelDataValidationInfo projectedTextLength = Assert.Single(document.Sheets[0].GetDataValidations("G2:G5"));
            Assert.Equal("textLength", projectedTextLength.Type);
            Assert.Equal("lessThanOrEqual", projectedTextLength.Operator);
            Assert.Equal("12", projectedTextLength.Formula1);
            Assert.Null(projectedTextLength.Formula2);
            Assert.Equal("Invalid text", projectedTextLength.ErrorTitle);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            List<DataValidation> openXmlValidations = worksheetPart.Worksheet.Descendants<DataValidation>().ToList();
            Assert.Equal(3, openXmlValidations.Count);
            DataValidation openXmlDate = Assert.Single(openXmlValidations, validation => validation.Type!.Value == DataValidationValues.Date);
            Assert.Equal(DataValidationOperatorValues.Between, openXmlDate.Operator!.Value);
            Assert.Equal("E2:E5", openXmlDate.SequenceOfReferences!.InnerText);
            Assert.Equal(expectedStartDate, openXmlDate.GetFirstChild<Formula1>()!.Text);
            Assert.Equal(expectedEndDate, openXmlDate.GetFirstChild<Formula2>()!.Text);
            Assert.Equal("Ship date", openXmlDate.PromptTitle!.Value);
            Assert.Equal("Invalid date", openXmlDate.ErrorTitle!.Value);
            DataValidation openXmlTime = Assert.Single(openXmlValidations, validation => validation.Type!.Value == DataValidationValues.Time);
            Assert.Equal(DataValidationOperatorValues.GreaterThanOrEqual, openXmlTime.Operator!.Value);
            Assert.Equal("F2:F5", openXmlTime.SequenceOfReferences!.InnerText);
            Assert.Equal(expectedStartTime, openXmlTime.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlTime.GetFirstChild<Formula2>());
            Assert.Equal("Invalid time", openXmlTime.ErrorTitle!.Value);
            DataValidation openXmlTextLength = Assert.Single(openXmlValidations, validation => validation.Type!.Value == DataValidationValues.TextLength);
            Assert.Equal(DataValidationOperatorValues.LessThanOrEqual, openXmlTextLength.Operator!.Value);
            Assert.Equal("G2:G5", openXmlTextLength.SequenceOfReferences!.InnerText);
            Assert.Equal("12", openXmlTextLength.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlTextLength.GetFirstChild<Formula2>());
            Assert.Equal("Invalid text", openXmlTextLength.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_Projects1904DateValidationBoundsAsOpenXmlDateSerials() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4Date1904DataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string expectedStartDate = new DateTime(1904, 1, 2).ToOADate().ToString("G15", CultureInfo.InvariantCulture);
            string expectedEndDate = new DateTime(1904, 1, 3).ToOADate().ToString("G15", CultureInfo.InvariantCulture);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(legacy.Uses1904DateSystem);
            LegacyXlsDataValidation dateValidation = Assert.Single(Assert.Single(legacy.Worksheets).DataValidations);
            Assert.Equal("1", dateValidation.Formula1);
            Assert.Equal("2", dateValidation.Formula2);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlDate = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.Date, openXmlDate.Type!.Value);
            Assert.Equal(expectedStartDate, openXmlDate.GetFirstChild<Formula1>()!.Text);
            Assert.Equal(expectedEndDate, openXmlDate.GetFirstChild<Formula2>()!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsCustomFormulaDataValidation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4CustomFormulaDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            const string ExpectedFormula = "SUM($A$1:$B$1)>10";

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.Custom, validation.Type);
            Assert.Equal(ExpectedFormula, validation.Formula1);
            Assert.Null(validation.Formula2);
            Assert.True(validation.AllowBlank);
            Assert.True(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);
            Assert.Equal("F2:F11", Assert.Single(validation.Ranges));
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            ExcelDataValidationInfo projectedValidation = Assert.Single(document.Sheets[0].GetDataValidations("F2:F11"));
            Assert.Equal("custom", projectedValidation.Type);
            Assert.Null(projectedValidation.Operator);
            Assert.Equal(ExpectedFormula, projectedValidation.Formula1);
            Assert.Null(projectedValidation.Formula2);
            Assert.Equal("Custom", projectedValidation.PromptTitle);
            Assert.Equal("Invalid total", projectedValidation.ErrorTitle);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            DataValidation openXmlValidation = Assert.Single(worksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Equal(DataValidationValues.Custom, openXmlValidation.Type!.Value);
            Assert.Null(openXmlValidation.Operator);
            Assert.Equal("F2:F11", openXmlValidation.SequenceOfReferences!.InnerText);
            Assert.Equal(ExpectedFormula, openXmlValidation.GetFirstChild<Formula1>()!.Text);
            Assert.Null(openXmlValidation.GetFirstChild<Formula2>());
            Assert.Equal("Custom", openXmlValidation.PromptTitle!.Value);
            Assert.Equal("Invalid total", openXmlValidation.ErrorTitle!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ReportsDataValidationRecordsAsPreserveOnly() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5DataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Validation", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Validated"));
            Assert.Equal(2, legacy.UnsupportedFeatures.Count(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation));
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Code == "XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED" && feature.SheetName == "Validation" && feature.RecordType == 0x01b2);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Code == "XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED" && feature.SheetName == "Validation" && feature.RecordType == 0x01be);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.DetailCode == "DataValidation:DVal");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.DetailCode == "DataValidation:Dv");
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED" && d.SheetName == "Validation" && d.RecordType == 0x01b2);
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED" && d.SheetName == "Validation" && d.RecordType == 0x01be);
            Assert.Contains(legacy.Diagnostics, d => d.DetailCode == "DataValidation:DVal");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.True(result.HasUnsupportedFeatures);
            Assert.Equal(2, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DataValidation]);
            Assert.Equal(2, result.ImportReport.UnsupportedFeaturesByCode["XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["DataValidation|XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED|DataValidation:DVal"]);
        }

        [Fact]
        public void LegacyXls_Load_ReportsDataValidationFormulaTokenBlocker() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFormulaDataValidationWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Empty(Assert.Single(result.Workbook.Worksheets).DataValidations);
            Assert.Contains(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation && feature.DetailCode == "DataValidation:Dv");
            LegacyXlsImportDiagnostic diagnostic = Assert.Single(result.Workbook.Diagnostics, d =>
                d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED"
                && d.RecordType == (ushort)BiffRecordType.Dv);
            Assert.Equal("FormulaToken0x01", diagnostic.DetailCode);
            Assert.Equal("DataValidation", diagnostic.FormulaContext);
            Assert.Equal((byte)0x01, diagnostic.FormulaToken);
            Assert.Equal("PtgExp", diagnostic.FormulaTokenName);
            Assert.Equal(0, diagnostic.FormulaTokenOffset);
            Assert.Contains("Data-validation formula", diagnostic.Message);
            Assert.Contains("PtgExp", diagnostic.Message);

            Assert.Equal(1, result.ImportReport.FormulaTokenBlockers["FormulaToken0x01"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByToken["Token:0x01"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByTokenName["PtgExp"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByOffset["Offset:0"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContext["DataValidation"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContextAndToken["DataValidation|Token:0x01"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContextAndTokenName["DataValidation|PtgExp"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContextAndDetail["DataValidation|FormulaToken0x01"]);
        }

        [Fact]
        public void LegacyXls_Load_ImportsConditionalFormattingCellIsRule() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4ConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(sheet.ConditionalFormattings);
            Assert.Equal(LegacyXlsConditionalFormattingType.CellIs, conditionalFormatting.Type);
            Assert.Equal(LegacyXlsConditionalFormattingOperator.GreaterThan, conditionalFormatting.Operator);
            Assert.Equal("10", conditionalFormatting.Formula1);
            Assert.Null(conditionalFormatting.Formula2);
            Assert.Equal(new[] { "A1:A3" }, conditionalFormatting.Ranges);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingCount);

            ExcelSheet projectedSheet = result.Document.Sheets[0];
            ExcelConditionalFormattingInfo info = Assert.Single(projectedSheet.GetConditionalFormattingRules("A1:A3"));
            Assert.Equal("CellIs", info.Type);
            Assert.Equal(nameof(ConditionalFormattingOperatorValues.GreaterThan), info.Operator);
            Assert.Equal(new[] { "10" }, info.Formulas);
            Assert.Empty(result.Document.ValidateOpenXml());

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            ConditionalFormatting openXmlFormatting = Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>());
            Assert.Equal("A1:A3", openXmlFormatting.SequenceOfReferences!.InnerText);
            ConditionalFormattingRule openXmlRule = Assert.Single(openXmlFormatting.Elements<ConditionalFormattingRule>());
            Assert.Equal(ConditionalFormatValues.CellIs, openXmlRule.Type!.Value);
            Assert.Equal(ConditionalFormattingOperatorValues.GreaterThan, openXmlRule.Operator!.Value);
            Assert.Equal("10", Assert.Single(openXmlRule.Elements<Formula>()).Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorksheetFeaturesAfterEmbeddedChartSubstream() {
            Assert.True(
                BiffDataValidationReader.TryRead(
                    LegacyXlsTestWorkbookBuilder.CreateExcelInlineListDataValidationPayloadForTest(),
                    out LegacyXlsDataValidation? directValidation),
                "The Excel-style inline-list DV payload should parse directly.");
            Assert.Equal(new[] { "Open", "Closed", "Pending" }, directValidation!.ListItems);

            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5EmbeddedChartBeforeWorksheetFeaturesWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(sheet.ConditionalFormattings);
            Assert.Equal(LegacyXlsConditionalFormattingType.CellIs, conditionalFormatting.Type);
            Assert.Equal(LegacyXlsConditionalFormattingOperator.GreaterThan, conditionalFormatting.Operator);
            Assert.Equal("100", conditionalFormatting.Formula1);
            Assert.Equal(new[] { "B2:B5" }, conditionalFormatting.Ranges);
            string unsupportedDetails = string.Join(",", legacy.UnsupportedFeatures.Select(feature => feature.DetailCode ?? feature.Code));
            Assert.True(sheet.DataValidations.Count == 1, unsupportedDetails);
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations);
            Assert.Equal(LegacyXlsDataValidationType.List, validation.Type);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.InlineList, validation.ListSourceKind);
            Assert.Equal("\"Open,Closed,Pending\"", validation.Formula1);
            Assert.Equal(new[] { "Open", "Closed", "Pending" }, validation.ListItems);
            Assert.Equal(new[] { "A2:A5" }, validation.Ranges);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DataValidation);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Equal(1, result.ImportReport.DataValidationCount);
            Assert.Equal(1, result.ImportReport.DataValidationListSourcesByKind["InlineList"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingCount);
            Assert.Equal(1, result.ImportReport.WorksheetFeatureStates["DataValidations:1|ConditionalFormatting:1|AutoFilterCriteria:0|AutoFilterDropDowns:Missing"]);
            ExcelSheet projectedSheet = result.Document.Sheets[0];
            Assert.Single(projectedSheet.GetDataValidations("A2:A5"));
            Assert.Single(projectedSheet.GetConditionalFormattingRules("B2:B5"));
        }

        [Fact]
        public void LegacyXls_Load_ImportsConditionalFormattingFormulaRule() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4ConditionalFormulaFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingCount);

            LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(legacySheet.ConditionalFormattings);
            Assert.Equal(LegacyXlsConditionalFormattingType.Formula, conditionalFormatting.Type);
            Assert.Null(conditionalFormatting.Operator);
            Assert.Equal("A1>10", conditionalFormatting.Formula1);
            Assert.Equal(new[] { "A1:A3" }, conditionalFormatting.Ranges);
            Assert.Equal(1, conditionalFormatting.RangeCount);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByRangeCount["Ranges:1"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByRange["A1:A3"]);

            ExcelSheet projectedSheet = result.Document.Sheets[0];
            ExcelConditionalFormattingInfo info = Assert.Single(projectedSheet.GetConditionalFormattingRules("A1:A3"));
            Assert.Equal("Expression", info.Type);
            Assert.Null(info.Operator);
            Assert.Equal(new[] { "A1>10" }, info.Formulas);
            Assert.Empty(result.Document.ValidateOpenXml());
        }

        [Fact]
        public void LegacyXls_Load_DiagnosesInlineConditionalFormattingBlocks() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4InlineFormattedConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.Equal(LegacyXlsConditionalFormattingType.CellIs, conditionalFormatting.Type);
            Assert.Equal("10", conditionalFormatting.Formula1);
            Assert.Contains(result.Workbook.UnsupportedFeatures, feature =>
                feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting
                && feature.DetailCode == "ConditionalFormatting:Cf");
            Assert.Equal(1, result.ImportReport.ConditionalFormattingCount);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.ConditionalFormatting]);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsInlineConditionalFormattingFill() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5InlineFillConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.Equal(LegacyXlsConditionalFormattingType.CellIs, conditionalFormatting.Type);
            Assert.Equal(LegacyXlsConditionalFormattingOperator.GreaterThan, conditionalFormatting.Operator);
            Assert.NotNull(conditionalFormatting.DifferentialFormat);
            Assert.Equal((byte)1, conditionalFormatting.DifferentialFormat!.FillPattern);
            Assert.Equal("FFFFFF00", conditionalFormatting.DifferentialFormat.FillForegroundColor);
            Assert.Equal("FFFFFF00", conditionalFormatting.DifferentialFormat.FillBackgroundColor);
            Assert.Equal((ushort)BiffRecordType.Cf, conditionalFormatting.DifferentialFormat.RecordType);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFill["Pattern:1"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFill["Foreground:FFFFFF00"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFill["Background:FFFFFF00"]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            ConditionalFormattingRule rule = Assert.Single(Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>()).Elements<ConditionalFormattingRule>());
            Assert.Equal(0U, rule.FormatId!.Value);
            DifferentialFormat dxf = Assert.Single(spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>());
            Fill fill = Assert.Single(dxf.Elements<Fill>());
            PatternFill patternFill = fill.PatternFill!;
            Assert.Equal(PatternValues.Solid, patternFill.PatternType!.Value);
            Assert.Equal("FFFFFF00", patternFill.ForegroundColor!.Rgb!.Value);
            Assert.Equal("FFFFFF00", patternFill.BackgroundColor!.Rgb!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsInlineConditionalFormattingFont() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5InlineFontConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            LegacyXlsDifferentialFormat differentialFormat = conditionalFormatting.DifferentialFormat!;
            Assert.Equal("FFFF0000", differentialFormat.FontColor);
            Assert.True(differentialFormat.FontBold);
            Assert.True(differentialFormat.FontItalic);
            Assert.Null(differentialFormat.FillPattern);
            Assert.Equal((ushort)BiffRecordType.Cf, differentialFormat.RecordType);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFont["Color:FFFF0000"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFont["Bold"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFont["Italic"]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            DifferentialFormat dxf = Assert.Single(spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>());
            Font font = Assert.Single(dxf.Elements<Font>());
            Assert.Equal("FFFF0000", Assert.Single(font.Elements<Color>()).Rgb!.Value);
            Assert.Single(font.Elements<Bold>());
            Assert.Single(font.Elements<Italic>());
            Assert.Empty(dxf.Elements<Fill>());
        }

        [Fact]
        public void LegacyXls_Load_ProjectsInlineConditionalFormattingFontAndFill() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5InlineFontAndFillConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsDifferentialFormat differentialFormat = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings).DifferentialFormat!;
            Assert.Equal("FFFF0000", differentialFormat.FontColor);
            Assert.True(differentialFormat.FontBold);
            Assert.True(differentialFormat.FontItalic);
            Assert.Equal((byte)1, differentialFormat.FillPattern);
            Assert.Equal("FFFFFF00", differentialFormat.FillForegroundColor);
            Assert.Equal("FFFFFF00", differentialFormat.FillBackgroundColor);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFont["Color:FFFF0000"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFill["Background:FFFFFF00"]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            DifferentialFormat dxf = Assert.Single(spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>());
            Assert.Single(dxf.Elements<Font>());
            Fill fill = Assert.Single(dxf.Elements<Fill>());
            Assert.Equal("FFFFFF00", fill.PatternFill!.ForegroundColor!.Rgb!.Value);
            Assert.Equal("FFFFFF00", fill.PatternFill.BackgroundColor!.Rgb!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsConditionalFormattingExtensionPriorityAndStopIfTrue() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ConditionalFormattingExtensionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.Equal(7, conditionalFormatting.Priority);
            Assert.True(conditionalFormatting.StopIfTrue);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByPriorityState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByPriority["Priority:7"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByStopIfTrueState["StopIfTrue"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Missing"]);
            Assert.Empty(result.ImportReport.ConditionalFormattingExtensionInlineFormattingByteCounts);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingExtensionDxfProjectionStates["NoDxfRequested"]);

            ExcelConditionalFormattingInfo info = Assert.Single(result.Document.Sheets[0].GetConditionalFormattingRules("A1:A3"));
            Assert.Equal(7, info.Priority);
            Assert.True(info.StopIfTrue);
            Assert.Empty(result.Document.ValidateOpenXml());

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            ConditionalFormattingRule rule = Assert.Single(Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>()).Elements<ConditionalFormattingRule>());
            Assert.Equal(1, (int)rule.Priority!.Value);
            Assert.True(rule.StopIfTrue!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ReportsConditionalFormattingRecordsAsPreserveOnly() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Conditional", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Formatted"));
            Assert.Equal(5, legacy.UnsupportedFeatures.Count(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting));
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Code == "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED" && feature.SheetName == "Conditional" && feature.RecordType == 0x01b0);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Code == "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED" && feature.SheetName == "Conditional" && feature.RecordType == 0x01b1);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Code == "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED" && feature.SheetName == "Conditional" && feature.RecordType == 0x087a);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Code == "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED" && feature.SheetName == "Conditional" && feature.RecordType == 0x087b);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Code == "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED" && feature.SheetName == "Conditional" && feature.RecordType == 0x088d);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.DetailCode == "ConditionalFormatting:CondFmt");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.DetailCode == "ConditionalFormatting:Cf12");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.DetailCode == "ConditionalFormatting:Dxf");
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED" && d.SheetName == "Conditional" && d.RecordType == 0x01b0);
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED" && d.SheetName == "Conditional" && d.RecordType == 0x088d);
            Assert.Contains(legacy.Diagnostics, d => d.DetailCode == "ConditionalFormatting:Dxf");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.True(result.HasUnsupportedFeatures);
            Assert.Equal(5, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.ConditionalFormatting]);
            Assert.Equal(5, result.ImportReport.UnsupportedFeaturesByCode["XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["ConditionalFormatting|XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED|ConditionalFormatting:CfEx"]);
        }

        [Fact]
        public void LegacyXls_Load_ReportsConditionalFormattingFormulaTokenBlocker() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFormulaConditionalFormattingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Empty(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.Contains(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting && feature.DetailCode == "ConditionalFormatting:Cf");
            LegacyXlsImportDiagnostic diagnostic = Assert.Single(result.Workbook.Diagnostics, d =>
                d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED"
                && d.RecordType == (ushort)BiffRecordType.Cf);
            Assert.Equal("FormulaToken0x01", diagnostic.DetailCode);
            Assert.Equal("ConditionalFormatting", diagnostic.FormulaContext);
            Assert.Equal((byte)0x01, diagnostic.FormulaToken);
            Assert.Equal("PtgExp", diagnostic.FormulaTokenName);
            Assert.Equal(0, diagnostic.FormulaTokenOffset);
            Assert.Contains("Conditional-formatting formula", diagnostic.Message);
            Assert.Contains("PtgExp", diagnostic.Message);

            Assert.Equal(1, result.ImportReport.FormulaTokenBlockers["FormulaToken0x01"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByToken["Token:0x01"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByTokenName["PtgExp"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByOffset["Offset:0"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContext["ConditionalFormatting"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContextAndToken["ConditionalFormatting|Token:0x01"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContextAndTokenName["ConditionalFormatting|PtgExp"]);
            Assert.Equal(1, result.ImportReport.FormulaTokenBlockersByContextAndDetail["ConditionalFormatting|FormulaToken0x01"]);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsSimpleTableDefinitionRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5TableDefinitionFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
            Assert.Equal("ListTable", legacySheet.Name);
            Assert.Contains(legacySheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Name"));
            LegacyXlsTableDefinition tableDefinition = Assert.Single(legacySheet.TableDefinitions);
            Assert.Equal("LegacySales", tableDefinition.Name);
            Assert.Equal("A1:B4", tableDefinition.Range);
            Assert.True(tableDefinition.HasHeaderRow);
            Assert.True(tableDefinition.HasTotalsRow);
            Assert.Equal(1U, tableDefinition.TotalRowCount);
            Assert.True(tableDefinition.HasAutoFilter);
            Assert.Equal("TableStyleMedium4", tableDefinition.StyleName);
            Assert.Equal("LegacySalesDisplay", tableDefinition.DisplayName);
            Assert.Equal("Imported legacy table comment", tableDefinition.Comment);
            LegacyXlsTableBlockLevelFormatting blockLevelFormatting = Assert.IsType<LegacyXlsTableBlockLevelFormatting>(tableDefinition.BlockLevelFormatting);
            Assert.Equal(0, blockLevelFormatting.HeaderStyleRecordIndex);
            Assert.Equal("LegacyHeaderStyle", blockLevelFormatting.HeaderStyleName);
            Assert.Equal(1, blockLevelFormatting.DataStyleRecordIndex);
            Assert.Equal("LegacyDataStyle", blockLevelFormatting.DataStyleName);
            Assert.Equal(2, blockLevelFormatting.TotalStyleRecordIndex);
            Assert.Equal("LegacyTotalStyle", blockLevelFormatting.TotalStyleName);
            Assert.True(tableDefinition.ShowFirstColumn);
            Assert.False(tableDefinition.ShowLastColumn);
            Assert.True(tableDefinition.ShowRowStripes);
            Assert.True(tableDefinition.ShowColumnStripes);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.TableDefinition);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.FeatureExtension);
            Assert.DoesNotContain(result.Workbook.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.TableDefinition);
            Assert.False(result.ImportReport.UnsupportedFeaturesByKind.ContainsKey(LegacyXlsUnsupportedFeatureKind.TableDefinition));
            Assert.False(result.ImportReport.PreservedFeatureRecordsByKind.ContainsKey(LegacyXlsUnsupportedFeatureKind.TableDefinition));
            Assert.False(result.ImportReport.UnsupportedProjectionGapsByKind.ContainsKey(LegacyXlsUnsupportedFeatureKind.TableDefinition));

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts);
            TableDefinitionPart tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
            Table table = tablePart.Table!;
            Assert.Equal("LegacySalesDisplay", table.Name?.Value);
            Assert.Equal("LegacySalesDisplay", table.DisplayName?.Value);
            Assert.Equal("Imported legacy table comment", table.Comment?.Value);
            Assert.Equal("A1:B4", table.Reference?.Value);
            Assert.Equal(1U, table.HeaderRowCount?.Value);
            Assert.True(table.TotalsRowShown?.Value);
            Assert.Equal(1U, table.TotalsRowCount?.Value);
            Assert.NotNull(table.GetFirstChild<AutoFilter>());
            TableStyleInfo styleInfo = table.TableStyleInfo!;
            Assert.Equal("TableStyleMedium4", styleInfo.Name?.Value);
            Assert.True(styleInfo.ShowFirstColumn?.Value);
            Assert.False(styleInfo.ShowLastColumn?.Value);
            Assert.True(styleInfo.ShowRowStripes?.Value);
            Assert.True(styleInfo.ShowColumnStripes?.Value);
            TableColumns columns = table.TableColumns!;
            Assert.Equal(2U, columns.Count?.Value);
            Assert.Equal(new[] { "Name", "Amount" }, columns.Elements<TableColumn>().Select(column => column.Name?.Value).ToArray());
            Dictionary<string, uint?> styleIndexes = worksheetPart.Worksheet.Descendants<Cell>()
                .Where(cell => cell.CellReference?.Value is "A1" or "B1" or "A2" or "B2" or "A3" or "B3" or "A4" or "B4")
                .ToDictionary(cell => cell.CellReference!.Value!, cell => cell.StyleIndex?.Value);
            Assert.True(styleIndexes["A1"] > 0U);
            Assert.Equal(styleIndexes["A1"], styleIndexes["B1"]);
            Assert.True(styleIndexes["A2"] > 0U);
            Assert.Equal(styleIndexes["A2"], styleIndexes["B2"]);
            Assert.Equal(styleIndexes["A2"], styleIndexes["A3"]);
            Assert.Equal(styleIndexes["B2"], styleIndexes["B3"]);
            Assert.True(styleIndexes["A4"] > 0U);
            Assert.Equal(styleIndexes["A4"], styleIndexes["B4"]);
            Assert.NotEqual(styleIndexes["A1"], styleIndexes["A2"]);
            Assert.NotEqual(styleIndexes["A2"], styleIndexes["A4"]);
        }

        [Fact]
        public void LegacyXls_Load_ModelsDifferentialFormatFillFromDxfRecord() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5DifferentialFormatWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            Assert.Equal(1, result.ImportReport.DifferentialFormatCount);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByRecordType["RecordType:0x088D"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByContentState["FillOnly"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByFill["Background:FFFFFF00"]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);
            Assert.Empty(result.ImportReport.DifferentialFormatsByFont);
            LegacyXlsDifferentialFormat format = Assert.Single(result.Workbook.DifferentialFormats);
            Assert.Equal(0, format.Index);
            Assert.Null(format.FillPattern);
            Assert.Null(format.FillForegroundColor);
            Assert.Equal("FFFFFF00", format.FillBackgroundColor);
            Assert.Equal(0x088d, format.RecordType);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.DetailCode == "ConditionalFormatting:Dxf");
        }

        [Fact]
        public void LegacyXls_Load_ProjectsConditionalFormattingDxfFillWhenSingleStyleIsAvailable() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ConditionalFormattingExtensionWithDxfWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.NotNull(conditionalFormatting.DifferentialFormat);
            Assert.Equal("FFFFFF00", conditionalFormatting.DifferentialFormat!.FillBackgroundColor);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByPriorityState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByPriority["Priority:1"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByStopIfTrueState["StopIfTrue"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFill["Background:FFFFFF00"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingExtensionInlineFormattingByteCounts["Bytes:16"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingExtensionDxfProjectionStates["ProjectedInlineDxfBytes:16"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByContentState["FillOnly"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByFill["Background:FFFFFF00"]);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            ConditionalFormattingRule rule = Assert.Single(Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>()).Elements<ConditionalFormattingRule>());
            Assert.Equal(0U, rule.FormatId!.Value);
            DifferentialFormat dxf = Assert.Single(spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>());
            Fill fill = Assert.Single(dxf.Elements<Fill>());
            PatternFill patternFill = fill.PatternFill!;
            Assert.Equal(PatternValues.Solid, patternFill.PatternType!.Value);
            Assert.Equal("FFFFFF00", patternFill.ForegroundColor!.Rgb!.Value);
            Assert.Equal("FFFFFF00", patternFill.BackgroundColor!.Rgb!.Value);
        }

        [Fact]
        public void LegacyXls_Load_AttachesConditionalFormattingDxfToMatchingPriorityRule() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ConditionalFormattingPriorityDxfWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.Equal(2, Assert.Single(result.Workbook.Worksheets).ConditionalFormattings.Count);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            Dictionary<int, ConditionalFormattingRule> rulesByPriority = worksheetPart.Worksheet
                .Elements<ConditionalFormatting>()
                .SelectMany(formatting => formatting.Elements<ConditionalFormattingRule>())
                .ToDictionary(rule => rule.Priority!.Value);

            Assert.Null(rulesByPriority[2].FormatId);
            Assert.Equal(0U, rulesByPriority[1].FormatId!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsConditionalFormattingDxfFontWhenSingleStyleIsAvailable() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ConditionalFormattingExtensionWithFontDxfWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.NotNull(conditionalFormatting.DifferentialFormat);
            Assert.Equal("FFFF0000", conditionalFormatting.DifferentialFormat!.FontColor);
            Assert.True(conditionalFormatting.DifferentialFormat.FontBold);
            Assert.True(conditionalFormatting.DifferentialFormat.FontItalic);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFont["Color:FFFF0000"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFont["Bold"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFont["Italic"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByRecordType["RecordType:0x088D"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByContentState["FontOnly"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByFont["Color:FFFF0000"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByFont["Bold"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByFont["Italic"]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("Differential Formats By Content State", markdown);
            Assert.Contains("Differential Formats By Font", markdown);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            ConditionalFormattingRule rule = Assert.Single(Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>()).Elements<ConditionalFormattingRule>());
            Assert.Equal(0U, rule.FormatId!.Value);
            DifferentialFormat dxf = Assert.Single(spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>());
            Font font = Assert.Single(dxf.Elements<Font>());
            Assert.Equal("FFFF0000", Assert.Single(font.Elements<Color>()).Rgb!.Value);
            Assert.Single(font.Elements<Bold>());
            Assert.Single(font.Elements<Italic>());
        }

        [Fact]
        public void LegacyXls_Load_ProjectsConditionalFormattingDxfBorderWhenSingleStyleIsAvailable() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ConditionalFormattingExtensionWithBorderDxfWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.NotNull(conditionalFormatting.DifferentialFormat);
            Assert.NotNull(conditionalFormatting.DifferentialFormat!.Border);
            LegacyXlsDifferentialBorder border = conditionalFormatting.DifferentialFormat.Border!;
            Assert.NotNull(border.Top);
            LegacyXlsDifferentialBorderSide top = border.Top!;
            Assert.Equal(1, top.Style);
            Assert.Equal("FF336699", top.Color);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialBorder["Top:Style:1"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialBorder["Top:Color:FF336699"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByRecordType["RecordType:0x088D"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByContentState["BorderOnly"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByBorder["Top:Style:1"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByBorder["Top:Color:FF336699"]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("Differential Formats By Border", markdown);
            Assert.Contains("Conditional Formatting By Differential Border", markdown);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            ConditionalFormattingRule rule = Assert.Single(Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>()).Elements<ConditionalFormattingRule>());
            Assert.Equal(0U, rule.FormatId!.Value);
            DifferentialFormat dxf = Assert.Single(spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>());
            Border openXmlBorder = Assert.Single(dxf.Elements<Border>());
            Assert.Equal(BorderStyleValues.Thin, openXmlBorder.TopBorder!.Style!.Value);
            Assert.Equal("FF336699", openXmlBorder.TopBorder.Color!.Rgb!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsConditionalFormattingDxfNumberFormatWhenSingleStyleIsAvailable() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5ConditionalFormattingExtensionWithNumberFormatDxfWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            Assert.False(result.HasUnsupportedFeatures);
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(Assert.Single(result.Workbook.Worksheets).ConditionalFormattings);
            Assert.NotNull(conditionalFormatting.DifferentialFormat);
            LegacyXlsDifferentialFormat differentialFormat = conditionalFormatting.DifferentialFormat!;
            Assert.Equal((ushort)164, differentialFormat.NumberFormatId);
            Assert.Equal("$#,##0.00", differentialFormat.NumberFormatCode);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialFormatState["Present"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialNumberFormat["Id:164"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByDifferentialNumberFormat["Code:$#,##0.00"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByRecordType["RecordType:0x088D"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByContentState["NumberFormatOnly"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByNumberFormat["Id:164"]);
            Assert.Equal(1, result.ImportReport.DifferentialFormatsByNumberFormat["Code:$#,##0.00"]);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting);
            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("Differential Formats By Number Format", markdown);
            Assert.Contains("Conditional Formatting By Differential Number Format", markdown);

            using var packageStream = new MemoryStream();
            result.Document.Save(packageStream);
            packageStream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(packageStream, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            ConditionalFormattingRule rule = Assert.Single(Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>()).Elements<ConditionalFormattingRule>());
            Assert.Equal(0U, rule.FormatId!.Value);
            DifferentialFormat dxf = Assert.Single(spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>());
            NumberingFormat numberingFormat = Assert.Single(dxf.Elements<NumberingFormat>());
            Assert.Equal(164U, numberingFormat.NumberFormatId!.Value);
            Assert.Equal("$#,##0.00", numberingFormat.FormatCode!.Value);
        }
    }
}
