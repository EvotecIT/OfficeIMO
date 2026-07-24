using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private const string ExternalLinkPathRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        [Fact]
        public void LegacyXls_NativeSave_WritesDataConsolidationSettings() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Consolidation");
                    sheet.CellValue(1, 1, "Value");
                    sheet.WorksheetPart.Worksheet.Append(new DataConsolidate {
                        Function = DataConsolidateFunctionValues.Maximum,
                        LeftLabels = true,
                        TopLabels = true,
                        Link = true
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                Assert.NotNull(worksheet.DataConsolidationSettings);
                Assert.Equal(LegacyXlsDataConsolidationFunction.Maximum, worksheet.DataConsolidationSettings!.Function);
                Assert.True(worksheet.DataConsolidationSettings.UsesLeftLabels);
                Assert.True(worksheet.DataConsolidationSettings.UsesTopLabels);
                Assert.True(worksheet.DataConsolidationSettings.LinksToSourceData);

                DataConsolidate projected = result.Document.Sheets.Single()
                    .WorksheetPart.Worksheet
                    .Elements<DataConsolidate>()
                    .Single();
                Assert.Equal(DataConsolidateFunctionValues.Maximum, projected.Function!.Value);
                Assert.True(projected.LeftLabels!.Value);
                Assert.True(projected.TopLabels!.Value);
                Assert.True(projected.Link!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesDataConsolidationStartLabels() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Consolidation");
                    sheet.CellValue(1, 1, "Region");
                    sheet.CellValue(2, 1, "North");
                    sheet.WorksheetPart.Worksheet.Append(new DataConsolidate {
                        Function = DataConsolidateFunctionValues.Sum,
                        StartLabels = true
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                Assert.NotNull(worksheet.DataConsolidationSettings);
                Assert.Equal(LegacyXlsDataConsolidationFunction.Sum, worksheet.DataConsolidationSettings!.Function);
                Assert.True(worksheet.DataConsolidationSettings.UsesLeftLabels);
                Assert.False(worksheet.DataConsolidationSettings.UsesTopLabels);
                Assert.False(worksheet.DataConsolidationSettings.LinksToSourceData);

                DataConsolidate projected = result.Document.Sheets.Single()
                    .WorksheetPart.Worksheet
                    .Elements<DataConsolidate>()
                    .Single();
                Assert.Equal(DataConsolidateFunctionValues.Sum, projected.Function!.Value);
                Assert.True(projected.LeftLabels!.Value);
                Assert.Null(projected.StartLabels);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_Load_ReadsSpecShapedDataConsolidationSettings() {
            byte[] payload = {
                0x03, 0x00,
                0x01, 0x00,
                0x01, 0x00,
                0x01, 0x00
            };

            Assert.True(BiffDataConsolidationSettingsReader.TryRead(
                new BiffRecord((ushort)BiffRecordType.DCon, 0, payload),
                out LegacyXlsDataConsolidationSettings? settings));

            Assert.NotNull(settings);
            Assert.Equal(LegacyXlsDataConsolidationFunction.Maximum, settings!.Function);
            Assert.True(settings.UsesLeftLabels);
            Assert.True(settings.UsesTopLabels);
            Assert.True(settings.LinksToSourceData);
            Assert.Equal(0x0007, settings.OptionFlags);
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSelfDataConsolidationReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet source = document.AddWorksheet("Source");
                    source.CellValue(1, 1, "Region");
                    source.CellValue(2, 1, "North");
                    source.CellValue(2, 2, 12d);

                    ExcelSheet sheet = document.AddWorksheet("Consolidation");
                    sheet.CellValue(1, 1, "Consolidated");
                    sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                        new DataReferences(
                            new DataReference {
                                Sheet = "Source",
                                Reference = "A1:B2"
                            })) {
                        Function = DataConsolidateFunctionValues.Sum
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsDataConsolidationReference reference = Assert.Single(result.Workbook.DataConsolidationReferences);
                Assert.Equal(LegacyXlsDataConsolidationSourceKind.SelfReference, reference.SourceKind);
                Assert.Equal("Source", reference.Source);
                Assert.Equal("A1:B2", reference.CellRange);

                ExcelSheet projectedSheet = result.Document.Sheets.Single(sheet => sheet.Name == "Consolidation");
                DataConsolidate projected = projectedSheet.WorksheetPart.Worksheet
                    .Elements<DataConsolidate>()
                    .Single();
                DataReference projectedReference = projected.GetFirstChild<DataReferences>()!
                    .Elements<DataReference>()
                    .Single();
                Assert.Equal("Source", projectedReference.Sheet!.Value);
                Assert.Equal("A1:B2", projectedReference.Reference!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookNamedDataConsolidationReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet source = document.AddWorksheet("Source");
                    source.CellValue(1, 1, "Region");
                    source.CellValue(2, 1, "North");
                    source.CellValue(2, 2, 12d);

                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName("'Source'!$A$1:$B$2") {
                        Name = "ConsolidationSource"
                    });

                    ExcelSheet sheet = document.AddWorksheet("Consolidation");
                    sheet.CellValue(1, 1, "Consolidated");
                    sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                        new DataReferences(
                            new DataReference {
                                Name = "ConsolidationSource"
                            })) {
                        Function = DataConsolidateFunctionValues.Sum
                    });
                    sheet.WorksheetPart.Worksheet.Save();
                    document.WorkbookRoot.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsDataConsolidationName consolidationName = Assert.Single(result.Workbook.DataConsolidationNames);
                Assert.Equal(LegacyXlsDataConsolidationSourceKind.SelfReference, consolidationName.SourceKind);
                Assert.Equal("ConsolidationSource", consolidationName.Name);
                Assert.Equal(string.Empty, consolidationName.Source);
                Assert.Equal(1, result.ImportReport.DataConsolidationNameCount);
                Assert.Equal(1, result.ImportReport.DataConsolidationNamesBySourceKind["SelfReference"]);
                Assert.Equal(1, result.ImportReport.DataConsolidationNamesByName["ConsolidationSource"]);
                Assert.Equal(1, result.ImportReport.DataConsolidationNamesBySource["(self)"]);

                Assert.Contains(result.Workbook.DefinedNames, name => name.Name == "ConsolidationSource" && name.Reference == "'Source'!$A$1:$B$2");

                ExcelSheet projectedSheet = result.Document.Sheets.Single(sheet => sheet.Name == "Consolidation");
                DataConsolidate projected = projectedSheet.WorksheetPart.Worksheet
                    .Elements<DataConsolidate>()
                    .Single();
                DataReference projectedReference = projected.GetFirstChild<DataReferences>()!
                    .Elements<DataReference>()
                    .Single();
                Assert.Equal("ConsolidationSource", projectedReference.Name!.Value);
                Assert.Null(projectedReference.Sheet);
                Assert.Null(projectedReference.Reference);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetScopedNamedDataConsolidationReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet source = document.AddWorksheet("Source");
                    source.CellValue(1, 1, "Region");
                    source.CellValue(2, 1, "North");
                    source.CellValue(2, 2, 12d);

                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName("$A$1:$B$2") {
                        Name = "LocalSource",
                        LocalSheetId = 0
                    });

                    ExcelSheet sheet = document.AddWorksheet("Consolidation");
                    sheet.CellValue(1, 1, "Consolidated");
                    sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                        new DataReferences(
                            new DataReference {
                                Name = "LocalSource",
                                Sheet = "Source"
                            })) {
                        Function = DataConsolidateFunctionValues.Sum
                    });
                    sheet.WorksheetPart.Worksheet.Save();
                    document.WorkbookRoot.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsDataConsolidationName consolidationName = Assert.Single(result.Workbook.DataConsolidationNames);
                Assert.Equal(LegacyXlsDataConsolidationSourceKind.SelfReference, consolidationName.SourceKind);
                Assert.Equal("LocalSource", consolidationName.Name);
                Assert.Equal("Source", consolidationName.Source);
                Assert.Contains(result.Workbook.DefinedNames, name => name.Name == "LocalSource" && name.LocalSheetIndex == 0);

                ExcelSheet projectedSheet = result.Document.Sheets.Single(sheet => sheet.Name == "Consolidation");
                DataConsolidate projected = projectedSheet.WorksheetPart.Worksheet
                    .Elements<DataConsolidate>()
                    .Single();
                DataReference projectedReference = projected.GetFirstChild<DataReferences>()!
                    .Elements<DataReference>()
                    .Single();
                Assert.Equal("LocalSource", projectedReference.Name!.Value);
                Assert.Equal("Source", projectedReference.Sheet!.Value);
                Assert.Null(projectedReference.Reference);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalDataConsolidationReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Consolidation");
                    sheet.CellValue(1, 1, "Consolidated");
                    ExternalRelationship relationship = sheet.WorksheetPart.AddExternalRelationship(
                        ExternalLinkPathRelationshipType,
                        new Uri("../Data/Budget.xls", UriKind.Relative));
                    sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                        new DataReferences(
                            new DataReference {
                                Reference = "B2:D4",
                                Id = relationship.Id
                            })) {
                        Function = DataConsolidateFunctionValues.Sum
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                    xlsOutputPath,
                    new LegacyXlsImportOptions { PreserveExternalWorkbookLinks = true });
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsDataConsolidationReference reference = Assert.Single(result.Workbook.DataConsolidationReferences);
                Assert.Equal(LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath, reference.SourceKind);
                Assert.Equal("../Data/Budget.xls", reference.Source);
                Assert.Equal("B2:D4", reference.CellRange);
                Assert.Equal((byte)0x01, reference.SourcePrefix);
                Assert.Equal(1, result.ImportReport.DataConsolidationReferenceCount);
                Assert.Equal(1, result.ImportReport.DataConsolidationReferencesBySourceKind["ExternalVirtualPath"]);
                Assert.Equal(1, result.ImportReport.DataConsolidationReferencesBySource["../Data/Budget.xls"]);

                ExcelSheet projectedSheet = result.Document.Sheets.Single(sheet => sheet.Name == "Consolidation");
                ExternalRelationship projectedRelationship = Assert.Single(projectedSheet.WorksheetPart.ExternalRelationships);
                Assert.Equal(ExternalLinkPathRelationshipType, projectedRelationship.RelationshipType);
                Assert.Equal("../Data/Budget.xls", projectedRelationship.Uri.OriginalString);
                DataReference projectedReference = projectedSheet.WorksheetPart.Worksheet
                    .Elements<DataConsolidate>()
                    .Single()
                    .GetFirstChild<DataReferences>()!
                    .Elements<DataReference>()
                    .Single();
                Assert.Equal("B2:D4", projectedReference.Reference!.Value);
                Assert.Equal(projectedRelationship.Id, projectedReference.Id!.Value);
                Assert.Null(projectedReference.Sheet);
                Assert.Null(projectedReference.Name);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalNamedDataConsolidationReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Consolidation");
                    sheet.CellValue(1, 1, "Consolidated");
                    ExternalRelationship relationship = sheet.WorksheetPart.AddExternalRelationship(
                        ExternalLinkPathRelationshipType,
                        new Uri("../Data/Budget.xls", UriKind.Relative));
                    sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                        new DataReferences(
                            new DataReference {
                                Name = "ExternalBudget",
                                Id = relationship.Id
                            })) {
                        Function = DataConsolidateFunctionValues.Sum
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                    xlsOutputPath,
                    new LegacyXlsImportOptions { PreserveExternalWorkbookLinks = true });
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsDataConsolidationName consolidationName = Assert.Single(result.Workbook.DataConsolidationNames);
                Assert.Equal(LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath, consolidationName.SourceKind);
                Assert.Equal("ExternalBudget", consolidationName.Name);
                Assert.Equal("../Data/Budget.xls", consolidationName.Source);
                Assert.Equal(1, result.ImportReport.DataConsolidationNameCount);
                Assert.Equal(1, result.ImportReport.DataConsolidationNamesBySourceKind["ExternalVirtualPath"]);
                Assert.Equal(1, result.ImportReport.DataConsolidationNamesBySource["../Data/Budget.xls"]);

                ExcelSheet projectedSheet = result.Document.Sheets.Single(sheet => sheet.Name == "Consolidation");
                ExternalRelationship projectedRelationship = Assert.Single(projectedSheet.WorksheetPart.ExternalRelationships);
                Assert.Equal(ExternalLinkPathRelationshipType, projectedRelationship.RelationshipType);
                Assert.Equal("../Data/Budget.xls", projectedRelationship.Uri.OriginalString);
                DataReference projectedReference = projectedSheet.WorksheetPart.Worksheet
                    .Elements<DataConsolidate>()
                    .Single()
                    .GetFirstChild<DataReferences>()!
                    .Elements<DataReference>()
                    .Single();
                Assert.Equal("ExternalBudget", projectedReference.Name!.Value);
                Assert.Equal(projectedRelationship.Id, projectedReference.Id!.Value);
                Assert.Null(projectedReference.Sheet);
                Assert.Null(projectedReference.Reference);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedDataConsolidationSourceReferencePayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data consolidation source reference payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Consolidated");
                string oversizedSheetName = new string('\u0100', 33000);
                sheet.WorksheetPart.Worksheet.Append(new DataConsolidate(
                    new DataReferences(
                        new DataReference {
                            Sheet = oversizedSheetName,
                            Reference = "A1:A1"
                        })) {
                    Function = DataConsolidateFunctionValues.Sum
                });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }
    }
}
