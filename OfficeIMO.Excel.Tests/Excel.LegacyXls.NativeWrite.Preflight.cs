using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesProtectedRanges() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Protected");
                    sheet.CellValue(1, 1, "Protected range");
                    sheet.CellValue(2, 3, "Second range");
                    sheet.Protect();

                    sheet.WorksheetPart.Worksheet.AppendChild(new ProtectedRanges(
                        new ProtectedRange {
                            Name = "UnlockedBlock",
                            Password = "BEEF",
                            SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1 C2:D2" }
                        }));
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsProtectedRange protectedRange = Assert.Single(worksheet.ProtectedRanges);
                Assert.Equal("UnlockedBlock", protectedRange.Name);
                Assert.Equal("BEEF", protectedRange.LegacyPasswordHash);
                Assert.Equal(new[] { "A1", "C2:D2" }, protectedRange.References);
                AssertBiffRecordOccursBefore(xlsOutputPath, 0x0867, 0x0868);

                ProtectedRange projectedRange = Assert.Single(
                    result.Document.Sheets.Single()
                        .WorksheetPart.Worksheet
                        .Elements<ProtectedRanges>()
                        .Single()
                        .Elements<ProtectedRange>());
                Assert.Equal("UnlockedBlock", projectedRange.Name!.Value);
                Assert.Equal("BEEF", projectedRange.Password!.Value);
                Assert.Equal("A1 C2:D2", projectedRange.SequenceOfReferences!.InnerText);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_TreatsSingleCellProtectedRangeReferencesAsOneCellRanges() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("SingleProtected");
                    sheet.CellValue(1, 1, "Editable");
                    sheet.Protect();
                    sheet.WorksheetPart.Worksheet.AppendChild(new ProtectedRanges(
                        new ProtectedRange {
                            Name = "OneCell",
                            SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
                        }));
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsProtectedRange protectedRange = Assert.Single(worksheet.ProtectedRanges);
                Assert.Equal("A1", Assert.Single(protectedRange.References));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedProtectedRangePayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("protected range payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Protected range");
                sheet.Protect();
                sheet.WorksheetPart.Worksheet.AppendChild(new ProtectedRanges(
                    new ProtectedRange {
                        Name = new string('P', 9000),
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                    }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksProtectedRangeUnsupportedAttributesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("protected ranges with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Protected range metadata");
                sheet.Protect();
                var protectedRange = new ProtectedRange {
                    Name = "UnlockedBlock",
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                };
                protectedRange.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.AppendChild(new ProtectedRanges(protectedRange));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksProtectedRangeUnsupportedCollectionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("protected ranges with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Protected range collection metadata");
                sheet.Protect();
                var protectedRanges = new ProtectedRanges(
                    new ProtectedRange {
                        Name = "UnlockedBlock",
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                    });
                protectedRanges.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.AppendChild(protectedRanges);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksModernWorkbookProtectionHashesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("modern workbook protection hashes", (document, sheet) => {
                sheet.CellValue(1, 1, "Modern workbook hash");
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    ProtectStructure = true,
                    LegacyPasswordHash = "CAFE"
                });

                WorkbookProtection protection = document.WorkbookRoot.GetFirstChild<WorkbookProtection>()!;
                protection.SetAttribute(new OpenXmlAttribute("workbookAlgorithmName", string.Empty, "SHA-512"));
                protection.SetAttribute(new OpenXmlAttribute("workbookHashValue", string.Empty, Convert.ToBase64String(new byte[] { 1, 2, 3, 4 })));
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksModernWorkbookRevisionProtectionHashesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("workbook revision protection", (document, sheet) => {
                sheet.CellValue(1, 1, "Revision protection");
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    ProtectStructure = true,
                    LegacyPasswordHash = "CAFE"
                });

                WorkbookProtection protection = document.WorkbookRoot.GetFirstChild<WorkbookProtection>()!;
                protection.LockRevision = true;
                protection.RevisionsPassword = "BEEF";
                protection.SetAttribute(new OpenXmlAttribute("revisionsAlgorithmName", string.Empty, "SHA-512"));
                protection.SetAttribute(new OpenXmlAttribute("revisionsHashValue", string.Empty, Convert.ToBase64String(new byte[] { 9, 8, 7, 6 })));
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksModernWorksheetProtectionHashesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("modern worksheet protection hashes", (document, sheet) => {
                sheet.CellValue(1, 1, "Modern worksheet hash");
                sheet.Protect(new ExcelSheetProtectionOptions {
                    LegacyPasswordHash = "BEEF"
                });

                SheetProtection protection = sheet.WorksheetPart.Worksheet.Elements<SheetProtection>().Single();
                protection.SetAttribute(new OpenXmlAttribute("algorithmName", string.Empty, "SHA-512"));
                protection.SetAttribute(new OpenXmlAttribute("hashValue", string.Empty, Convert.ToBase64String(new byte[] { 5, 6, 7, 8 })));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInvalidWorkbookProtectionHashesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("invalid workbook protection password hashes", (document, sheet) => {
                sheet.CellValue(1, 1, "Invalid workbook hash");
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    ProtectStructure = true,
                    LegacyPasswordHash = "CAFE"
                });

                WorkbookProtection protection = document.WorkbookRoot.GetFirstChild<WorkbookProtection>()!;
                protection.SetAttribute(new OpenXmlAttribute("workbookPassword", string.Empty, "10000"));
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInvalidWorksheetProtectionHashesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("invalid worksheet protection password hashes", (document, sheet) => {
                sheet.CellValue(1, 1, "Invalid worksheet hash");
                sheet.Protect(new ExcelSheetProtectionOptions {
                    LegacyPasswordHash = "BEEF"
                });

                SheetProtection protection = sheet.WorksheetPart.Worksheet.Elements<SheetProtection>().Single();
                protection.SetAttribute(new OpenXmlAttribute("password", string.Empty, "10000"));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInvalidWriteReservationHashesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("invalid write-reservation password hashes", (document, sheet) => {
                sheet.CellValue(1, 1, "Invalid write reservation hash");
                document.SetWriteReservation(new ExcelWorkbookWriteReservationOptions {
                    ReadOnlyRecommended = true,
                    UserName = "Reviewer"
                });

                FileSharing fileSharing = document.WorkbookRoot.GetFirstChild<FileSharing>()!;
                fileSharing.SetAttribute(new OpenXmlAttribute("reservationPassword", string.Empty, "10000"));
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorkbookExtensionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("workbook extension metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Workbook extension metadata");
                document.WorkbookRoot.Append(new WorkbookExtensionList(
                    new WorkbookExtension { Uri = "{4F3E2D1C-0000-4000-8000-000000000001}" }));
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetExtensionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet extension metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Worksheet extension metadata");
                sheet.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(
                    new WorksheetExtension { Uri = "{4F3E2D1C-0000-4000-8000-000000000002}" }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSparklinesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("sparklines", (document, sheet) => {
                sheet.CellValue(1, 1, "Jan");
                sheet.CellValue(1, 2, "Feb");
                sheet.CellValue(2, 1, 10d);
                sheet.CellValue(2, 2, 12d);
                sheet.AddSparklines("A2:B2", "C2");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksIgnoredCalculatedColumnErrorsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("ignored calculated-column errors", (document, sheet) => {
                sheet.CellValue(1, 1, "Text number");
                sheet.WorksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors(
                    new IgnoredError {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" },
                        CalculatedColumn = true
                    }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksIgnoredErrorsUnsupportedCollectionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("ignored errors with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Text number");
                var ignoredErrors = new DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors(
                    new IgnoredError {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" },
                        NumberStoredAsText = true
                    });
                ignoredErrors.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.Append(ignoredErrors);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedCellWatchReferencesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell watch references outside BIFF8 worksheet limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Watched");
                sheet.WorksheetPart.Worksheet.Append(new CellWatches(
                    new CellWatch { CellReference = "XFD1048576" }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCellWatchesUnsupportedCollectionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell watches with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Watched");
                var cellWatches = new CellWatches(
                    new CellWatch { CellReference = "A1" });
                cellWatches.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.Append(cellWatches);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterDropdownControlMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter dropdown control metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                autoFilter.Append(new FilterColumn(
                    new Filters(new Filter { Val = "Alpha" })) {
                    ColumnId = 0U,
                    HiddenButton = true
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterExtensionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter extension metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                autoFilter.Append(new ExtensionList(
                    new Extension { Uri = "{4F3E2D1C-0000-4000-8000-000000000004}" }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterSortMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter sort metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                autoFilter.Append(new SortState { Reference = "A1:A2" });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterUnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                autoFilter.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterColumnUnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter column metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                var filterColumn = new FilterColumn(
                    new Filters(new Filter { Val = "Alpha" })) {
                    ColumnId = 0U
                };
                filterColumn.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                autoFilter.Append(filterColumn);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterCriteriaUnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter criteria metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                var filters = new Filters(new Filter { Val = "Alpha" });
                filters.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                autoFilter.Append(new FilterColumn(filters) {
                    ColumnId = 0U
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateAutoFilterCriteriaContainersBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter columns with multiple filter types", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                autoFilter.Append(new FilterColumn(
                    new Filters(new Filter { Val = "Alpha" }),
                    new Filters(new Filter { Val = "Beta" })) {
                    ColumnId = 0U
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedAutoFilterDateGroupCriteriaBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter date-group criteria", (document, sheet) => {
                AppendAutoFilterColumn(sheet, new Filters(
                    new DateGroupItem {
                        Year = 2026,
                        Month = 6,
                        Day = 28,
                        DateTimeGrouping = DateTimeGroupingValues.Day
                    },
                    new DateGroupItem {
                        Year = 2026,
                        Month = 6,
                        Day = 29,
                        DateTimeGrouping = DateTimeGroupingValues.Day
                    }));
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterDynamicCriteriaBeforeWriting() {
            AssertNativeXlsSaveNotSupported("dynamic, color, icon, or extension AutoFilter criteria", (document, sheet) => {
                AppendAutoFilterColumn(sheet, new DynamicFilter {
                    Type = DynamicFilterValues.ThisMonth
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterColorCriteriaBeforeWriting() {
            AssertNativeXlsSaveNotSupported("dynamic, color, icon, or extension AutoFilter criteria", (document, sheet) => {
                AppendAutoFilterColumn(sheet, new ColorFilter {
                    CellColor = true
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterIconCriteriaBeforeWriting() {
            AssertNativeXlsSaveNotSupported("dynamic, color, icon, or extension AutoFilter criteria", (document, sheet) => {
                AppendAutoFilterColumn(sheet, new IconFilter {
                    IconSet = IconSetValues.ThreeTrafficLights1,
                    IconId = 0U
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterEqualityListsOutsideBiffLimitBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter equality lists with more than two values", (document, sheet) => {
                AppendAutoFilterColumn(sheet, new Filters(
                    new Filter { Val = "Alpha" },
                    new Filter { Val = "Beta" },
                    new Filter { Val = "Gamma" }));
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAutoFilterBlankPlusMultipleValuesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("AutoFilter blank criteria combined with more than one value", (document, sheet) => {
                AppendAutoFilterColumn(sheet, new Filters(
                    new Filter { Val = "Alpha" },
                    new Filter { Val = "Beta" }) {
                    Blank = true
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateWorksheetSingletonElementsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("multiple worksheet AutoFilter", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AutoFilterAdd("A1:A2");

                sheet.WorksheetPart.Worksheet.Append(new AutoFilter { Reference = "B1:B2" });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCustomSheetViewsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("custom sheet views", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom sheet view");
                sheet.WorksheetPart.Worksheet.Append(new CustomSheetViews(
                    new CustomSheetView { Guid = "{4F3E2D1C-0000-4000-8000-000000000003}" }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetScenariosWithoutChangedCellsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet scenario cell counts outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Scenario value");
                sheet.WorksheetPart.Worksheet.Append(new Scenarios(
                    new Scenario { Name = "Best case" }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataConsolidationSourceReferencesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("external data consolidation source relationships", (document, sheet) => {
                sheet.CellValue(1, 1, "Consolidated");
                sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                    new DataReferences(
                        new DataReference {
                            Reference = "A1:A2",
                            Id = "rIdMissingExternalWorkbook"
                        })) {
                    Function = DataConsolidateFunctionValues.Sum
                });
                sheet.WorksheetPart.Worksheet.Save();
            });

            AssertNativeXlsSaveNotSupported("external data consolidation sheet-qualified source references", (document, sheet) => {
                sheet.CellValue(1, 1, "Consolidated");
                sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                    new DataReferences(
                        new DataReference {
                            Reference = "A1:A2",
                            Sheet = "Source",
                            Id = "rIdExternalWorkbook"
                        })) {
                    Function = DataConsolidateFunctionValues.Sum
                });
                sheet.WorksheetPart.Worksheet.Save();
            });

            AssertNativeXlsSaveNotSupported("named data consolidation source references with explicit cell ranges", (document, sheet) => {
                sheet.CellValue(1, 1, "Consolidated");
                sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                    new DataReferences(
                        new DataReference {
                            Name = "LocalSource",
                            Sheet = "Source",
                            Reference = "A1:A2"
                        })) {
                    Function = DataConsolidateFunctionValues.Sum
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataConsolidationSourceReferenceCollectionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data consolidation source references", (document, sheet) => {
                sheet.CellValue(1, 1, "Consolidated");
                var dataReferences = new DataReferences(
                    new DataReference {
                        Reference = "A1:A2",
                        Sheet = "Source"
                    });
                dataReferences.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(dataReferences) {
                    Function = DataConsolidateFunctionValues.Sum
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataValidationCollectionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation collection metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 5d);

                var validations = new DataValidations {
                    Count = 1U
                };
                validations.SetAttribute(new OpenXmlAttribute("xWindow", string.Empty, "42"));
                validations.Append(new DataValidation(
                    new Formula1("1"),
                    new Formula2("10")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                });
                sheet.WorksheetPart.Worksheet.Append(validations);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateDataValidationCollectionsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("multiple worksheet data-validation collections", (document, sheet) => {
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 5d);
                sheet.CellValue(3, 1, 7d);

                sheet.WorksheetPart.Worksheet.Append(new DataValidations(
                    new DataValidation(
                        new Formula1("1"),
                        new Formula2("10")) {
                        Type = DataValidationValues.Whole,
                        Operator = DataValidationOperatorValues.Between,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                    }) {
                    Count = 1U
                });
                sheet.WorksheetPart.Worksheet.Append(new DataValidations(
                    new DataValidation(
                        new Formula1("1"),
                        new Formula2("10")) {
                        Type = DataValidationValues.Whole,
                        Operator = DataValidationOperatorValues.Between,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A3:A3" }
                    }) {
                    Count = 1U
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateDataValidationFormula1ElementsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation formulas with duplicate Formula1 elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 5d);

                var validation = new DataValidation(
                    new Formula1("1"),
                    new Formula1("2"),
                    new Formula2("10")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                };
                sheet.WorksheetPart.Worksheet.Append(new DataValidations(validation) { Count = 1U });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateDataValidationFormula2ElementsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation formulas with duplicate Formula2 elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 5d);

                var validation = new DataValidation(
                    new Formula1("1"),
                    new Formula2("10"),
                    new Formula2("20")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                };
                sheet.WorksheetPart.Worksheet.Append(new DataValidations(validation) { Count = 1U });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataValidationFormula1UnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation formula metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 5d);

                var formula = new Formula1("1");
                formula.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                var validation = new DataValidation(
                    formula,
                    new Formula2("10")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                };
                sheet.WorksheetPart.Worksheet.Append(new DataValidations(validation) { Count = 1U });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataValidationFormula2UnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation formula metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 5d);

                var formula = new Formula2("10");
                formula.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                var validation = new DataValidation(
                    new Formula1("1"),
                    formula) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                };
                sheet.WorksheetPart.Worksheet.Append(new DataValidations(validation) { Count = 1U });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataValidationExtensionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation extension metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(2, 1, 5d);

                var validation = new DataValidation(
                    new Formula1("1"),
                    new Formula2("10")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                };
                validation.Append(new ExtensionList(
                    new Extension { Uri = "{4F3E2D1C-0000-4000-8000-000000000005}" }));
                sheet.WorksheetPart.Worksheet.Append(new DataValidations(validation) { Count = 1U });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataValidationImeModeMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation IME mode metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Input");
                sheet.CellValue(2, 1, "Alpha");

                var validation = new DataValidation(
                    new Formula1("1"),
                    new Formula2("10")) {
                    Type = DataValidationValues.TextLength,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                };
                validation.SetAttribute(new OpenXmlAttribute("imeMode", string.Empty, "hiragana"));
                sheet.WorksheetPart.Worksheet.Append(new DataValidations(validation) { Count = 1U });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataValidationFormulasOutsideNativeSubsetBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation formulas outside the native XLS formula subset", (document, sheet) => {
                AppendDataValidation(sheet, new DataValidation(
                    new Formula1("SUM(Table1[Amount])>0")) {
                    Type = DataValidationValues.Custom,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksMissingRequiredDataValidationFormulasBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation formulas are required for this validation type", (document, sheet) => {
                AppendDataValidation(sheet, new DataValidation {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A2" }
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInvalidDataValidationRangesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation ranges", (document, sheet) => {
                AppendDataValidation(sheet, new DataValidation(
                    new Formula1("1"),
                    new Formula2("10")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "NotARange" }
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksTooManyDataValidationRangesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation range counts outside BIFF8 limits", (document, sheet) => {
                string ranges = string.Join(" ", Enumerable.Range(1, 433).Select(row => "A" + row.ToString(CultureInfo.InvariantCulture)));
                AppendDataValidation(sheet, new DataValidation(
                    new Formula1("1"),
                    new Formula2("10")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = ranges }
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDataValidationRangesOutsideBiffLimitsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation ranges outside BIFF8 worksheet limits", (document, sheet) => {
                AppendDataValidation(sheet, new DataValidation(
                    new Formula1("1"),
                    new Formula2("10")) {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "XFE1:XFE1" }
                });
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingPivotMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting pivot metadata", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("A1>0")) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    Pivot = true,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingExtensionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting extension metadata", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                var conditionalFormatting = new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("A1>0")) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                };
                conditionalFormatting.Append(new ExtensionList(
                    new Extension { Uri = "{4F3E2D1C-0000-4000-8000-000000000006}" }));
                sheet.WorksheetPart.Worksheet.Append(conditionalFormatting);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingUnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting metadata", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                var conditionalFormatting = new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("A1>0")) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                };
                conditionalFormatting.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                sheet.WorksheetPart.Worksheet.Append(conditionalFormatting);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingRuleUnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting rule metadata", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                var rule = new ConditionalFormattingRule(
                    new Formula("A1>0")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1
                };
                rule.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(rule) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingFormulaUnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting formula metadata", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                var formula = new Formula("A1>0");
                formula.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(formula) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingDifferentialFormatsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting differential formats", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("A1>0")) {
                        Type = ConditionalFormatValues.Expression,
                        FormatId = 0U,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingVisualPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting visual or extension payloads", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);
                sheet.AddConditionalColorScale("A1:A1", "FFFF0000", "FF00FF00");
            });

            AssertNativeXlsSaveNotSupported("conditional formatting visual or extension payloads", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);
                sheet.AddConditionalIconSet("A1:A1");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedConditionalFormattingRuleTypesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting rule types outside the BIFF8 classic rule subset", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule {
                        Type = ConditionalFormatValues.DataBar,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedConditionalFormattingOperatorsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting operators outside the BIFF8 classic rule subset", (document, sheet) => {
                sheet.CellValue(1, 1, "Ready");

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("\"Ready\"")) {
                        Type = ConditionalFormatValues.CellIs,
                        Operator = ConditionalFormattingOperatorValues.ContainsText,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingFormulasOutsideNativeSubsetBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting formulas outside the native XLS formula subset", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("SUM(Table1[Amount])>0")) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingFormulaBackedMultiRangesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting formula-backed rule ranges outside the native XLS subset", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);
                sheet.CellValue(1, 3, 5d);

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule {
                        Type = ConditionalFormatValues.DuplicateValues,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A1 C1:C1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInvalidConditionalFormattingRangesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting ranges", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("A1>0")) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "NotARange" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksTooManyConditionalFormattingRangesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting range counts outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);
                string ranges = string.Join(" ", Enumerable.Range(1, 8193).Select(row => "A" + row.ToString(CultureInfo.InvariantCulture)));

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("A1>0")) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = ranges }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConditionalFormattingRangesOutsideBiffLimitsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting ranges outside BIFF8 worksheet limits", (document, sheet) => {
                sheet.CellValue(1, 1, 5d);

                sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                    new ConditionalFormattingRule(
                        new Formula("A1>0")) {
                        Type = ConditionalFormatValues.Expression,
                        Priority = 1
                    }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "XFE1:XFE1" }
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedWorksheetCalculationPropertiesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet calculation properties", (document, sheet) => {
                sheet.CellValue(1, 1, "Sheet calculation");
                var properties = new SheetCalculationProperties {
                    FullCalculationOnLoad = true
                };
                properties.SetAttribute(new OpenXmlAttribute("unsupported", string.Empty, "1"));
                sheet.WorksheetPart.Worksheet.Append(properties);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedWorkbookCalculationPropertiesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("workbook calculation properties", (document, sheet) => {
                sheet.CellValue(1, 1, "Workbook calculation");
                document.WorkbookRoot.Append(new CalculationProperties {
                    FullCalculationOnLoad = true,
                    ForceFullCalculation = true
                });
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateWorkbookSingletonElementsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("multiple workbook calculation property elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Duplicate workbook calculation");
                document.WorkbookRoot.RemoveAllChildren<CalculationProperties>();
                document.WorkbookRoot.Append(new CalculationProperties {
                    CalculationMode = CalculateModeValues.Manual
                });
                document.WorkbookRoot.Append(new CalculationProperties {
                    CalculationMode = CalculateModeValues.Auto
                });
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedWorksheetPhoneticSettingsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet phonetic settings", (document, sheet) => {
                sheet.CellValue(1, 1, "Phonetic settings");
                var properties = new PhoneticProperties {
                    FontId = 0U
                };
                properties.SetAttribute(new OpenXmlAttribute("unsupported", string.Empty, "1"));
                sheet.WorksheetPart.Worksheet.Append(properties);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksHeaderFooterImagesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("header or footer images", (document, sheet) => {
                sheet.CellValue(1, 1, "Header image");
                sheet.SetHeaderImage(HeaderFooterPosition.Center, new byte[] { 0x89, 0x50, 0x4E, 0x47 }, "image/png", widthPoints: 24D, heightPoints: 16D);
            });

            AssertNativeXlsSaveNotSupported("header or footer images", (document, sheet) => {
                sheet.CellValue(1, 1, "Footer image");
                sheet.SetFooterImage(HeaderFooterPosition.Center, new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 }, "image/jpeg", widthPoints: 24D, heightPoints: 16D);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedHeaderFooterTextBeforeWriting() {
            AssertNativeXlsSaveNotSupported("header or footer text lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Oversized header");
                sheet.SetHeaderFooter(headerCenter: new string('H', 70000));
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedHeaderFooterExtensionTextBeforeWriting() {
            AssertNativeXlsSaveNotSupported("header or footer text lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Oversized first-page header");
                sheet.SetFirstPageHeaderFooter(headerCenter: new string('F', 70000));
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCustomWorkbookViewsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("custom workbook views", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom workbook view");
                Workbook workbook = document.WorkbookRoot;
                workbook.RemoveAllChildren<CustomWorkbookViews>();
                var customWorkbookViews = new CustomWorkbookViews(new CustomWorkbookView());
                Sheets? sheets = workbook.GetFirstChild<Sheets>();
                if (sheets != null) {
                    workbook.InsertBefore(customWorkbookViews, sheets);
                } else {
                    workbook.Append(customWorkbookViews);
                }

                workbook.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksVbaProjectsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("VBA projects or macros", (document, sheet) => {
                sheet.CellValue(1, 1, "Macro workbook");
                VbaProjectPart vbaProjectPart = document.WorkbookPartRoot.AddNewPart<VbaProjectPart>();
                using var stream = new MemoryStream(new byte[] { 1, 2, 3, 4 });
                vbaProjectPart.FeedData(stream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCustomXmlPartsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("custom XML parts", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom XML");
                CustomXmlPart customXmlPart = document.WorkbookPartRoot.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<metadata><owner>OfficeIMO</owner></metadata>"));
                customXmlPart.FeedData(stream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorkbookDataPartRelationshipsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data part relationships", (document, sheet) => {
                sheet.CellValue(1, 1, "Workbook data relationship");
                AddMediaDataReference(document, document.WorkbookPartRoot);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetDataPartRelationshipsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data part relationships", (document, sheet) => {
                sheet.CellValue(1, 1, "Worksheet data relationship");
                AddMediaDataReference(document, sheet.WorksheetPart);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksThreadedCommentsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("threaded comments", (document, sheet) => {
                sheet.CellValue(1, 1, "Threaded");
                sheet.AddThreadedComment("A1", "Review this before release", "Modern Reviewer");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentObjectFillMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment object fill metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom comment fill");
                sheet.SetLegacyComment(1, 1, "Supported text with unsupported fill", "Reviewer", visible: false, anchor: null);
                SetCommentVmlFillColor(sheet, "#ffcccc");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentObjectLineMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment object line metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom comment outline");
                sheet.SetLegacyComment(1, 1, "Supported text with unsupported outline", "Reviewer", visible: false, anchor: null);
                SetCommentVmlLineColor(sheet, "#cc0000");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentObjectShadowMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment object shadow metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom comment shadow");
                sheet.SetLegacyComment(1, 1, "Supported text with unsupported shadow", "Reviewer", visible: false, anchor: null);
                SetCommentVmlShadowColor(sheet, "#666666");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentObjectTextboxMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment object textbox metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom comment textbox");
                sheet.SetLegacyComment(1, 1, "Supported text with unsupported textbox", "Reviewer", visible: false, anchor: null);
                SetCommentVmlTextboxStyle(sheet, "mso-direction-alt:auto;mso-fit-shape-to-text:t");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentObjectClientMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment object client metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom comment client data");
                sheet.SetLegacyComment(1, 1, "Supported text with unsupported client data", "Reviewer", visible: false, anchor: null);
                AddCommentVmlClientDataElement(sheet, "TextHAlign", "Center");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentObjectShapeMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment object shape metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom comment shape");
                sheet.SetLegacyComment(1, 1, "Supported text with unsupported shape style", "Reviewer", visible: false, anchor: null);
                SetCommentVmlShapeStyle(sheet, "position:absolute;margin-left:4pt;margin-top:0pt;width:108pt;height:59pt;z-index:1;visibility:hidden");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksCommentObjectPathMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("comment object shape metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Custom comment path");
                sheet.SetLegacyComment(1, 1, "Supported text with unsupported path", "Reviewer", visible: false, anchor: null);
                SetCommentVmlPathConnectType(sheet, "rect");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksChartSheetsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("unsupported sheet types", (document, sheet) => {
                sheet.CellValue(1, 1, "Chart sheet marker");
                document.WorkbookPartRoot.AddNewPart<ChartsheetPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDialogSheetsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("unsupported sheet types", (document, sheet) => {
                sheet.CellValue(1, 1, "Dialog sheet marker");
                document.WorkbookPartRoot.AddNewPart<DialogsheetPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksMacroSheetsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("unsupported sheet types", (document, sheet) => {
                sheet.CellValue(1, 1, "Macro sheet marker");
                document.WorkbookPartRoot.AddNewPart<MacroSheetPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInternationalMacroSheetsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("unsupported sheet types", (document, sheet) => {
                sheet.CellValue(1, 1, "International macro sheet marker");
                document.WorkbookPartRoot.AddNewPart<InternationalMacroSheetPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksPivotTableMarkersBeforeWriting() {
            AssertNativeXlsSaveNotSupported("PivotTables", (document, sheet) => {
                sheet.CellValue(1, 1, "Pivot marker");
                document.WorkbookRoot.Append(new PivotCaches(
                    new PivotCache {
                        CacheId = 1U,
                        Id = "rIdPivotCache"
                    }));
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksExternalWorkbookLinksBeforeWriting() {
            AssertNativeXlsSaveNotSupported("external workbook links", (document, sheet) => {
                sheet.CellValue(1, 1, "External link");
                document.WorkbookRoot.Append(new ExternalReferences(
                    new ExternalReference {
                        Id = "rIdExternalWorkbook"
                    }));
                document.WorkbookRoot.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksConnectionsAndQueryTablesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("connections or query tables", (document, sheet) => {
                sheet.CellValue(1, 1, "Connection");
                ExtendedPart connectionPart = document.WorkbookPartRoot.AddExtendedPart(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
                    "xml");
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"1\" name=\"SalesConnection\" type=\"5\" refreshedVersion=\"7\"/></connections>"))) {
                    connectionPart.FeedData(stream);
                }

                ExtendedPart queryTablePart = sheet.WorksheetPart.AddExtendedPart(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
                    "xml");
                using var queryStream = new MemoryStream(Encoding.UTF8.GetBytes("<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"SalesQuery\" connectionId=\"1\"/>"));
                queryTablePart.FeedData(queryStream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksEmbeddedOlePackagesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("embedded OLE objects or packages", (document, sheet) => {
                sheet.CellValue(1, 1, "Embedded package");
                EmbeddedPackagePart embeddedPackagePart = sheet.WorksheetPart.AddEmbeddedPackagePart(EmbeddedPackagePartType.Xlsx);
                using var stream = new MemoryStream(new byte[] { 5, 6, 7, 8 });
                embeddedPackagePart.FeedData(stream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksEmbeddedOleObjectsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("embedded OLE objects or packages", (document, sheet) => {
                sheet.CellValue(1, 1, "Embedded object");
                EmbeddedObjectPart embeddedObjectPart = sheet.WorksheetPart.AddEmbeddedObjectPart("application/vnd.openxmlformats-officedocument.oleObject");
                using var stream = new MemoryStream(new byte[] { 9, 10, 11, 12 });
                embeddedObjectPart.FeedData(stream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksFormControlsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("form controls", (document, sheet) => {
                sheet.CellValue(1, 1, "Form control");
                const string controlsXml = "<x:controls xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:control shapeId=\"1026\" name=\"ApproveButton\" r:id=\"rIdControl1\" /></x:controls>";
                sheet.WorksheetPart.Worksheet.Append(new Controls(controlsXml));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetSlicersBeforeWriting() {
            AssertNativeXlsSaveNotSupported("slicers or timelines", (document, sheet) => {
                sheet.CellValue(1, 1, "Worksheet slicer");
                sheet.WorksheetPart.AddNewPart<SlicersPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetTimelinesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("slicers or timelines", (document, sheet) => {
                sheet.CellValue(1, 1, "Worksheet timeline");
                sheet.WorksheetPart.AddNewPart<TimeLinePart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorkbookMetadataExtensionsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("workbook metadata extensions", (document, sheet) => {
                sheet.CellValue(1, 1, "Metadata extension");
                document.WorkbookPartRoot.AddNewPart<CellMetadataPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksRichDataFeaturesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("rich data features", (document, sheet) => {
                sheet.CellValue(1, 1, "Rich data");
                document.WorkbookPartRoot.AddNewPart<RdRichValuePart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksRichStylesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("rich styles", (document, sheet) => {
                sheet.CellValue(1, 1, "Rich styles");
                document.WorkbookPartRoot.AddNewPart<RichStylesPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateCellFontPropertiesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell font properties with duplicate font name elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Duplicate font");
                sheet.CellAt(1, 1).SetFontName("Arial");

                Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                var duplicateFont = new Font(
                    new FontName { Val = "Arial" },
                    new FontName { Val = "Calibri" });
                stylesheet.Fonts!.Append(duplicateFont);
                stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                var duplicateFontFormat = (CellFormat)baseFormat.CloneNode(true);
                duplicateFontFormat.FontId = stylesheet.Fonts.Count!.Value - 1U;
                duplicateFontFormat.ApplyFont = true;
                stylesheet.CellFormats.Append(duplicateFontFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksUnsupportedCellFontMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell font properties with unsupported metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Font metadata");
                sheet.CellAt(1, 1).SetFontName("Arial");

                Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                var metadataFont = new Font(new FontName { Val = "Arial" });
                var scheme = new OpenXmlUnknownElement(string.Empty, "scheme", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                scheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, "minor"));
                metadataFont.AppendChild(scheme);
                stylesheet.Fonts!.Append(metadataFont);
                stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

                var metadataFontFormat = (CellFormat)baseFormat.CloneNode(true);
                metadataFontFormat.FontId = stylesheet.Fonts.Count!.Value - 1U;
                metadataFontFormat.ApplyFont = true;
                stylesheet.CellFormats.Append(metadataFontFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateFillStyleChildrenBeforeWriting() {
            AssertNativeXlsSaveNotSupported("fill style with duplicate foreground color elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Duplicate fill");
                sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                var duplicateFill = new Fill(new PatternFill {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = "FF123456" },
                    BackgroundColor = new BackgroundColor { Indexed = 64U }
                });
                duplicateFill.PatternFill!.Append(new ForegroundColor { Rgb = "FF654321" });
                stylesheet.Fills!.Append(duplicateFill);
                stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

                var duplicateFillFormat = (CellFormat)baseFormat.CloneNode(true);
                duplicateFillFormat.FillId = stylesheet.Fills.Count!.Value - 1U;
                duplicateFillFormat.ApplyFill = true;
                stylesheet.CellFormats.Append(duplicateFillFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateBorderStyleChildrenBeforeWriting() {
            AssertNativeXlsSaveNotSupported("border style with duplicate left border elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Duplicate border");
                sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                var duplicateBorder = new Border(
                    new LeftBorder(new Color { Rgb = "FF123456" }) { Style = BorderStyleValues.Thin },
                    new LeftBorder(new Color { Rgb = "FF654321" }) { Style = BorderStyleValues.Medium });
                stylesheet.Borders!.Append(duplicateBorder);
                stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

                var duplicateBorderFormat = (CellFormat)baseFormat.CloneNode(true);
                duplicateBorderFormat.BorderId = stylesheet.Borders.Count!.Value - 1U;
                duplicateBorderFormat.ApplyBorder = true;
                stylesheet.CellFormats.Append(duplicateBorderFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateCellFormatAlignmentChildrenBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell-format style with duplicate alignment elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Duplicate alignment");
                sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                var duplicateAlignmentFormat = (CellFormat)baseFormat.CloneNode(true);
                duplicateAlignmentFormat.Append(
                    new Alignment { Horizontal = HorizontalAlignmentValues.Left },
                    new Alignment { Horizontal = HorizontalAlignmentValues.Right });
                duplicateAlignmentFormat.ApplyAlignment = true;
                stylesheet.CellFormats.Append(duplicateAlignmentFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDuplicateCellFormatProtectionChildrenBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell-format style with duplicate protection elements", (document, sheet) => {
                sheet.CellValue(1, 1, "Duplicate protection");
                sheet.CellAt(1, 1).SetFillColor("#ABCDEF");

                Stylesheet stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
                Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single();
                uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                CellFormat baseFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)baseStyleIndex);
                var duplicateProtectionFormat = (CellFormat)baseFormat.CloneNode(true);
                duplicateProtectionFormat.Append(
                    new Protection { Locked = true },
                    new Protection { Locked = false });
                duplicateProtectionFormat.ApplyProtection = true;
                stylesheet.CellFormats.Append(duplicateProtectionFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                stylesheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksVolatileDependencyMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("volatile dependency metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Volatile dependencies");
                document.WorkbookPartRoot.AddNewPart<VolatileDependenciesPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksRevisionOrUserDataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("revision or user data", (document, sheet) => {
                sheet.CellValue(1, 1, "Revision tracking");
                document.WorkbookPartRoot.AddNewPart<WorkbookRevisionHeaderPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksAttachedToolbarsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("attached toolbars", (document, sheet) => {
                sheet.CellValue(1, 1, "Attached toolbar");
                document.WorkbookPartRoot.AddNewPart<ExcelAttachedToolbarsPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDigitalSignatureOriginPartsBeforeWriting() {
            AssertNativeXlsSignedSaveBlocked((document, sheet) => {
                sheet.CellValue(1, 1, "Signed workbook");
                document._spreadSheetDocument.AddDigitalSignatureOriginPart();
                DigitalSignatureOriginPart originPart = document._spreadSheetDocument.DigitalSignatureOriginPart!;
                XmlSignaturePart signaturePart = originPart.AddNewPart<XmlSignaturePart>();
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes(
                    "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo /></Signature>"));
                signaturePart.FeedData(stream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDigitalSignatureApplicationPropertiesBeforeWriting() {
            AssertNativeXlsSignedSaveBlocked((document, sheet) => {
                sheet.CellValue(1, 1, "Signed workbook metadata");
                ExtendedFilePropertiesPart appPart = document._spreadSheetDocument.ExtendedFilePropertiesPart
                    ?? document._spreadSheetDocument.AddExtendedFilePropertiesPart();
                appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                appPart.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
                appPart.Properties.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksDirectWorksheetImagePartsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("drawings, images, or charts", (document, sheet) => {
                sheet.CellValue(1, 1, "Direct image relationship");
                ImagePart imagePart = sheet.WorksheetPart.AddImagePart(ImagePartType.Png);
                using var stream = new MemoryStream(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
                imagePart.FeedData(stream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksModel3DReferencesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("drawings, images, or charts", (document, sheet) => {
                sheet.CellValue(1, 1, "3D model relationship");
                sheet.WorksheetPart.AddNewPart<Model3DReferenceRelationshipPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSingleCellTablesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("tables", (document, sheet) => {
                sheet.CellValue(1, 1, "Single-cell table");
                sheet.WorksheetPart.AddNewPart<SingleCellTablePart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetTablesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("tables", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                sheet.AddTable("A1:A2", hasHeader: true, name: "NativeTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksTableDefinitionPartsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("tables", (document, sheet) => {
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alpha");
                TableDefinitionPart tablePart = sheet.WorksheetPart.AddNewPart<TableDefinitionPart>("rIdTableDefinition");
                tablePart.Table = new Table {
                    Id = 1U,
                    Name = "NativeTable",
                    DisplayName = "NativeTable",
                    Reference = "A1:A2"
                };
                tablePart.Table.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedCustomNumberFormatsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("custom number format lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, 123.45d);
                WorkbookStylesPart stylesPart = document.WorkbookPartRoot.WorkbookStylesPart
                    ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
                Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                stylesheet.NumberingFormats ??= new NumberingFormats();
                stylesheet.Fonts ??= new Fonts(new Font());
                stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }));
                stylesheet.Borders ??= new Borders(new Border());
                stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
                stylesheet.CellFormats ??= new CellFormats(new CellFormat());

                const uint numberFormatId = 200U;
                stylesheet.NumberingFormats.Append(new NumberingFormat {
                    NumberFormatId = numberFormatId,
                    FormatCode = new string('0', 70000)
                });
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
                stylesheet.CellFormats.Append(new CellFormat {
                    NumberFormatId = numberFormatId,
                    ApplyNumberFormat = true
                });
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

                Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                    .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1U;
                stylesheet.Save();
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksNamedSheetViewsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("named sheet views", (document, sheet) => {
                sheet.CellValue(1, 1, "Named sheet view");
                sheet.WorksheetPart.AddNewPart<NamedSheetViewsPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetSortMapsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet sort maps", (document, sheet) => {
                sheet.CellValue(1, 1, "Sort map");
                sheet.WorksheetPart.AddNewPart<WorksheetSortMapPart>();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetCustomPropertiesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("worksheet custom properties", (document, sheet) => {
                sheet.CellValue(1, 1, "Worksheet custom property");
                sheet.WorksheetPart.AddNewPart<CustomPropertyPart>("application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty+xml");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInlinePhoneticTextCellsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("phonetic cell text", (document, sheet) => {
                sheet.CellValue(1, 1, "Inline phonetic text");
                Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                    .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                cell.DataType = CellValues.InlineString;
                cell.RemoveAllChildren<CellValue>();
                cell.RemoveAllChildren<InlineString>();
                cell.Append(new InlineString(
                    new Text("Displayed"),
                    new PhoneticProperties { FontId = 0U }));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksSharedStringPhoneticTextCellsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("phonetic cell text", (document, sheet) => {
                sheet.CellValue(1, 1, "Shared phonetic text");
                SharedStringTablePart sharedStringPart = document.WorkbookPartRoot.SharedStringTablePart
                    ?? document.WorkbookPartRoot.AddNewPart<SharedStringTablePart>();
                sharedStringPart.SharedStringTable = new SharedStringTable(
                    new SharedStringItem(
                        new Text("Displayed"),
                        new PhoneticProperties { FontId = 0U }));
                sharedStringPart.SharedStringTable.Save();

                Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                    .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                cell.DataType = CellValues.SharedString;
                cell.RemoveAllChildren<CellValue>();
                cell.RemoveAllChildren<InlineString>();
                cell.Append(new CellValue("0"));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedCellTextBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cell text lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Oversized text");
                Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                    .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                cell.DataType = CellValues.InlineString;
                cell.RemoveAllChildren<CellValue>();
                cell.RemoveAllChildren<InlineString>();
                cell.Append(new InlineString(new Text(new string('T', 32768))));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedCachedFormulaTextBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cached formula text lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Formula text");
                Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                    .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                cell.DataType = CellValues.String;
                cell.CellFormula = new CellFormula("1");
                cell.RemoveAllChildren<CellValue>();
                cell.RemoveAllChildren<InlineString>();
                cell.Append(new CellValue(new string('F', 32768)));
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedHyperlinkPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("hyperlink payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Oversized internal link");
                Hyperlinks hyperlinks = sheet.WorksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault()
                    ?? sheet.WorksheetPart.Worksheet.AppendChild(new Hyperlinks());
                hyperlinks.Append(new Hyperlink {
                    Reference = "A1",
                    Display = "Long jump",
                    Location = new string('A', 40000)
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedHyperlinkTooltipsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("hyperlink tooltips outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Oversized tooltip");
                AddExternalHyperlink(sheet, "A1", "https://officeimo.net/legacy-xls", UriKind.Absolute, tooltip: new string('T', 256));
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksHyperlinkCollectionMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("hyperlink collection metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Hyperlink metadata");
                Hyperlinks hyperlinks = sheet.WorksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault()
                    ?? sheet.WorksheetPart.Worksheet.AppendChild(new Hyperlinks());
                hyperlinks.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                hyperlinks.Append(new Hyperlink {
                    Reference = "A1",
                    Location = "Sheet1!A1"
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksHyperlinkUnsupportedMetadataBeforeWriting() {
            AssertNativeXlsSaveNotSupported("hyperlink metadata", (document, sheet) => {
                sheet.CellValue(1, 1, "Hyperlink metadata");
                Hyperlinks hyperlinks = sheet.WorksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault()
                    ?? sheet.WorksheetPart.Worksheet.AppendChild(new Hyperlinks());
                var hyperlink = new Hyperlink {
                    Reference = "A1",
                    Location = "Sheet1!A1"
                };
                hyperlink.SetAttribute(new OpenXmlAttribute("customMetadata", string.Empty, "present"));
                hyperlinks.Append(hyperlink);
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksWorksheetImagesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("drawings, images, or charts", (document, sheet) => {
                sheet.CellValue(1, 1, "Worksheet image");
                sheet.AddImage(2, 1, new byte[] { 0x89, 0x50, 0x4E, 0x47 }, "image/png", widthPixels: 24, heightPixels: 16);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksChartsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("drawings, images, or charts", (document, sheet) => {
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 15d);
                sheet.AddChartFromRange("A1:B3", 1, 4, type: ExcelChartType.ColumnClustered, hasHeaders: true, title: "Chart");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksNonCommentVmlShapesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("legacy VML drawings or shapes", (document, sheet) => {
                sheet.CellValue(1, 1, "Comment plus shape");
                sheet.SetLegacyComment(1, 1, "Supported comment", "Reviewer", visible: false, anchor: null);
                AddVmlDrawing(sheet, "Button", row: 1, column: 0);
            });
        }

        private static void AppendAutoFilterColumn(ExcelSheet sheet, OpenXmlElement criteria) {
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 1, "Alpha");
            sheet.CellValue(3, 1, "Beta");
            sheet.CellValue(4, 1, "Gamma");
            sheet.AutoFilterAdd("A1:A4");

            AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
            autoFilter.Append(new FilterColumn(criteria) {
                ColumnId = 0U
            });
            sheet.WorksheetPart.Worksheet.Save();
        }

        private static void AppendDataValidation(ExcelSheet sheet, DataValidation validation) {
            sheet.CellValue(1, 1, "Amount");
            sheet.CellValue(2, 1, 5d);
            sheet.WorksheetPart.Worksheet.Append(new DataValidations(validation) { Count = 1U });
            sheet.WorksheetPart.Worksheet.Save();
        }

        private static void AddVmlDrawing(ExcelSheet sheet, string objectType, int row, int column) {
            VmlDrawingPart vmlPart = sheet.WorksheetPart.AddNewPart<VmlDrawingPart>();
            string relationshipId = sheet.WorksheetPart.GetIdOfPart(vmlPart);
            string vml = $@"<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">
<v:shape id=""_x0000_s1025"" type=""#_x0000_t201"" style=""position:absolute;margin-left:0;margin-top:0;width:80pt;height:20pt;z-index:1"">
<x:ClientData ObjectType=""{objectType}""><x:Row>{row}</x:Row><x:Column>{column}</x:Column></x:ClientData>
</v:shape>
</xml>";
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(vml))) {
                vmlPart.FeedData(stream);
            }

            sheet.WorksheetPart.Worksheet!.Append(new LegacyDrawing { Id = relationshipId });
            sheet.WorksheetPart.Worksheet.Save();
        }

        private static void SetCommentVmlLineColor(ExcelSheet sheet, string color) {
            MutateCommentVmlShape(sheet, shape => {
                XNamespace v = "urn:schemas-microsoft-com:vml";
                shape.SetAttributeValue("strokecolor", color);
                XElement? stroke = shape.Element(v + "stroke");
                if (stroke == null) {
                    stroke = new XElement(v + "stroke");
                    shape.AddFirst(stroke);
                }

                stroke.SetAttributeValue("color", color);
            });
        }

        private static void SetCommentVmlShadowColor(ExcelSheet sheet, string color) {
            MutateCommentVmlShape(sheet, shape => {
                XNamespace v = "urn:schemas-microsoft-com:vml";
                XElement? shadow = shape.Element(v + "shadow");
                if (shadow == null) {
                    shadow = new XElement(v + "shadow");
                    shape.AddFirst(shadow);
                }

                shadow.SetAttributeValue("on", "t");
                shadow.SetAttributeValue("color", color);
                shadow.SetAttributeValue("obscured", "t");
            });
        }

        private static void SetCommentVmlTextboxStyle(ExcelSheet sheet, string style) {
            MutateCommentVmlShape(sheet, shape => {
                XNamespace v = "urn:schemas-microsoft-com:vml";
                XElement? textbox = shape.Element(v + "textbox");
                if (textbox == null) {
                    textbox = new XElement(v + "textbox");
                    shape.Add(textbox);
                }

                textbox.SetAttributeValue("style", style);
            });
        }

        private static void AddCommentVmlClientDataElement(ExcelSheet sheet, string localName, string value) {
            MutateCommentVmlShape(sheet, shape => {
                XNamespace x = "urn:schemas-microsoft-com:office:excel";
                XElement clientData = shape.Element(x + "ClientData")!;
                clientData.Add(new XElement(x + localName, value));
            });
        }

        private static void SetCommentVmlShapeStyle(ExcelSheet sheet, string style) {
            MutateCommentVmlShape(sheet, shape => {
                shape.SetAttributeValue("style", style);
            });
        }

        private static void SetCommentVmlPathConnectType(ExcelSheet sheet, string connectType) {
            MutateCommentVmlShape(sheet, shape => {
                XNamespace v = "urn:schemas-microsoft-com:vml";
                XElement? path = shape.Element(v + "path");
                if (path == null) {
                    path = new XElement(v + "path");
                    shape.AddFirst(path);
                }

                path.SetAttributeValue(XName.Get("connecttype", "urn:schemas-microsoft-com:office:office"), connectType);
            });
        }

        private static void SetCommentVmlFillColor(ExcelSheet sheet, string color) {
            MutateCommentVmlShape(sheet, shape => {
                XNamespace v = "urn:schemas-microsoft-com:vml";
                shape.SetAttributeValue("fillcolor", color);
                XElement? fill = shape.Element(v + "fill");
                if (fill == null) {
                    fill = new XElement(v + "fill");
                    shape.AddFirst(fill);
                }

                fill.SetAttributeValue("color2", color);
            });
        }

        private static void MutateCommentVmlShape(ExcelSheet sheet, Action<XElement> mutate) {
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XDocument document;
            using (Stream stream = vmlPart.GetStream(FileMode.Open, FileAccess.Read)) {
                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }

            XElement shape = Assert.Single(document.Descendants(v + "shape"));
            mutate(shape);
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                document.Save(stream);
            }
        }

        private static void AddMediaDataReference(ExcelDocument document, OpenXmlPartContainer container) {
            MediaDataPart mediaDataPart = document._spreadSheetDocument.CreateMediaDataPart(MediaDataPartType.Mp3);
            using (var stream = new MemoryStream(new byte[] { 0x49, 0x44, 0x33, 0x03 })) {
                mediaDataPart.FeedData(stream);
            }

            MethodInfo? addRelationshipDefinition = null;
            foreach (MethodInfo method in typeof(OpenXmlPartContainer).GetMethods(BindingFlags.Instance | BindingFlags.NonPublic)) {
                ParameterInfo[] parameters = method.GetParameters();
                if (method.Name == "AddDataPartReferenceRelationship"
                    && method.ContainsGenericParameters
                    && parameters.Length == 1
                    && parameters[0].ParameterType == typeof(MediaDataPart)) {
                    addRelationshipDefinition = method;
                    break;
                }
            }

            Assert.NotNull(addRelationshipDefinition);
            MethodInfo addRelationship = addRelationshipDefinition!.MakeGenericMethod(typeof(AudioReferenceRelationship));
            addRelationship!.Invoke(container, new object[] { mediaDataPart });
            Assert.NotEmpty(container.DataPartReferenceRelationships);
        }
    }
}
