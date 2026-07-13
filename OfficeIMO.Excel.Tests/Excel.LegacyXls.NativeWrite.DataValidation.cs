using OpenXmlDataValidationOperatorValues = DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues;
using OpenXmlDataValidationValues = DocumentFormat.OpenXml.Spreadsheet.DataValidationValues;
using OpenXmlFormula1 = DocumentFormat.OpenXml.Spreadsheet.Formula1;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesAdditionalWorksheetDataValidationTypes() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            DateTime minimumDate = new DateTime(2024, 1, 1);
            DateTime maximumDate = new DateTime(2024, 12, 31);

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("MoreValidation");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(1, 2, "Date");
                    sheet.CellValue(1, 3, "Code");
                    sheet.CellValue(2, 1, 5d);

                    sheet.ValidationDate("B2:B5", OpenXmlDataValidationOperatorValues.Between, minimumDate, maximumDate, errorTitle: "Invalid date", errorMessage: "Use a date in 2024.");
                    sheet.ValidationTextLength("C2:C5", OpenXmlDataValidationOperatorValues.LessThanOrEqual, 12, errorTitle: "Invalid code", errorMessage: "Use at most 12 characters.");
                    sheet.ValidationCustomFormula("A2:A5", "A2>0", errorTitle: "Invalid amount", errorMessage: "Use a positive amount.");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsDataValidationCollectionRecord collection = Assert.Single(legacySheet.DataValidationCollections);
                Assert.Equal(3U, collection.DeclaredValidationCount);
                Assert.Equal(3, legacySheet.DataValidations.Count);

                string expectedMinimumDate = minimumDate.ToOADate().ToString(CultureInfo.InvariantCulture);
                string expectedMaximumDate = maximumDate.ToOADate().ToString(CultureInfo.InvariantCulture);
                LegacyXlsDataValidation dateValidation = Assert.Single(legacySheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.Date);
                Assert.Equal(LegacyXlsDataValidationOperator.Between, dateValidation.Operator);
                Assert.Equal(expectedMinimumDate, dateValidation.Formula1);
                Assert.Equal(expectedMaximumDate, dateValidation.Formula2);
                Assert.Equal("B2:B5", Assert.Single(dateValidation.Ranges));

                LegacyXlsDataValidation textLengthValidation = Assert.Single(legacySheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.TextLength);
                Assert.Equal(LegacyXlsDataValidationOperator.LessThanOrEqual, textLengthValidation.Operator);
                Assert.Equal("12", textLengthValidation.Formula1);
                Assert.Equal("C2:C5", Assert.Single(textLengthValidation.Ranges));

                LegacyXlsDataValidation customValidation = Assert.Single(legacySheet.DataValidations, validation => validation.Type == LegacyXlsDataValidationType.Custom);
                Assert.Equal("A2>0", customValidation.Formula1);
                Assert.Equal("A2:A5", Assert.Single(customValidation.Ranges));

                ExcelSheet projectedSheet = result.Document.Sheets[0];
                ExcelDataValidationInfo projectedDate = Assert.Single(projectedSheet.GetDataValidations("B2:B5"));
                Assert.Equal("date", projectedDate.Type);
                Assert.Equal("between", projectedDate.Operator);
                Assert.Equal(expectedMinimumDate, projectedDate.Formula1);
                Assert.Equal(expectedMaximumDate, projectedDate.Formula2);

                ExcelDataValidationInfo projectedTextLength = Assert.Single(projectedSheet.GetDataValidations("C2:C5"));
                Assert.Equal("textLength", projectedTextLength.Type);
                Assert.Equal("lessThanOrEqual", projectedTextLength.Operator);
                Assert.Equal("12", projectedTextLength.Formula1);

                ExcelDataValidationInfo projectedCustom = Assert.Single(projectedSheet.GetDataValidations("A2:A5"));
                Assert.Equal("custom", projectedCustom.Type);
                Assert.Equal("A2>0", projectedCustom.Formula1);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesRangeAndNamedListDataValidations() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet options = document.AddWorksheet("Options");
                    ExcelSheet input = document.AddWorksheet("Input");

                    options.CellValue(1, 1, "Open");
                    options.CellValue(2, 1, "Closed");
                    options.CellValue(3, 1, "Pending");
                    options.ValidationListRange("B1:B3", "A1:A3");

                    document.SetNamedRange("StatusOptions", "'Options'!A1:A3", save: false);
                    input.ValidationListRange("B2:B5", "A1:A3", "Options");
                    input.ValidationListNamedRange("C2:C5", "StatusOptions");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet optionsSheet = result.Workbook.Worksheets[0];
                LegacyXlsWorksheet inputSheet = result.Workbook.Worksheets[1];

                LegacyXlsDataValidation sameSheetRange = Assert.Single(optionsSheet.DataValidations);
                Assert.Equal(LegacyXlsDataValidationType.List, sameSheetRange.Type);
                Assert.Equal(LegacyXlsDataValidationListSourceKind.Range, sameSheetRange.ListSourceKind);
                Assert.Equal("A1:A3", sameSheetRange.ListSourceRange);
                Assert.Null(sameSheetRange.ListSourceSheetName);
                Assert.Equal("B1:B3", Assert.Single(sameSheetRange.Ranges));

                Assert.Equal(2, inputSheet.DataValidations.Count);
                LegacyXlsDataValidation crossSheetRange = Assert.Single(inputSheet.DataValidations, validation => validation.ListSourceKind == LegacyXlsDataValidationListSourceKind.SheetQualifiedRange);
                Assert.Equal("Options", crossSheetRange.ListSourceSheetName);
                Assert.Equal("A1:A3", crossSheetRange.ListSourceRange);
                Assert.Equal("B2:B5", Assert.Single(crossSheetRange.Ranges));

                LegacyXlsDataValidation namedRange = Assert.Single(inputSheet.DataValidations, validation => validation.ListSourceKind == LegacyXlsDataValidationListSourceKind.DefinedName);
                Assert.Equal("StatusOptions", namedRange.ListSourceName);
                Assert.Equal("C2:C5", Assert.Single(namedRange.Ranges));

                ExcelDataValidationInfo projectedSameSheet = Assert.Single(result.Document.Sheets[0].GetDataValidations("B1:B3"));
                Assert.Equal("list", projectedSameSheet.Type);
                Assert.Equal("=A1:A3", projectedSameSheet.Formula1);

                ExcelDataValidationInfo projectedCrossSheet = Assert.Single(result.Document.Sheets[1].GetDataValidations("B2:B5"));
                Assert.Equal("list", projectedCrossSheet.Type);
                Assert.Equal("='Options'!A1:A3", projectedCrossSheet.Formula1);

                ExcelDataValidationInfo projectedNamedRange = Assert.Single(result.Document.Sheets[1].GetDataValidations("C2:C5"));
                Assert.Equal("list", projectedNamedRange.Type);
                Assert.Equal("=StatusOptions", projectedNamedRange.Formula1);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_RebasesListValidationRangeReferencesToSqrefAnchor() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Validation");
                    sheet.CellValue(1, 1, "Open");
                    sheet.CellValue(2, 1, "Closed");
                    sheet.CellValue(3, 1, "Pending");
                    sheet.ValidationListRange("B1:B10", "A1:A3");

                    document.Save(xlsOutputPath);
                }

                byte[] payload = Assert.Single(GetBiffRecordPayloads(xlsOutputPath, 0x01be));
                byte[] formulaTokens = ReadDataValidationFormulaTokens(payload, formulaIndex: 0);
                Assert.Contains((byte)0x4d, formulaTokens);
                Assert.DoesNotContain((byte)0x45, formulaTokens);

                int areaOffset = Array.IndexOf(formulaTokens, (byte)0x4d);
                Assert.True(areaOffset >= 0);
                Assert.Equal((ushort)0, ReadUInt16(formulaTokens, areaOffset + 1));
                Assert.Equal((ushort)2, ReadUInt16(formulaTokens, areaOffset + 3));
                Assert.Equal((ushort)0xffff, ReadUInt16(formulaTokens, areaOffset + 5));
                Assert.Equal((ushort)0xffff, ReadUInt16(formulaTokens, areaOffset + 7));

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsDataValidation validation = Assert.Single(Assert.Single(result.Workbook.Worksheets).DataValidations);
                Assert.Equal("A1:A3", validation.ListSourceRange);
                ExcelDataValidationInfo projectedValidation = Assert.Single(result.Document.Sheets[0].GetDataValidations("B1:B10"));
                Assert.Equal("=A1:A3", projectedValidation.Formula1);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_TreatsSingleCellDataValidationReferencesAsOneCellRanges() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("SingleDv");
                    sheet.CellValue(1, 1, 3d);
                    sheet.ValidationWholeNumber("A1", OpenXmlDataValidationOperatorValues.Between, 1, 5);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsDataValidation validation = Assert.Single(legacySheet.DataValidations);
                Assert.Equal("A1", Assert.Single(validation.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetRangeListDataValidations() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet input = document.AddWorksheet("Input");
                    ExcelSheet firstRegion = document.AddWorksheet("Region 1");
                    ExcelSheet secondRegion = document.AddWorksheet("Region 2");

                    firstRegion.CellValue(1, 1, "North");
                    firstRegion.CellValue(2, 1, "East");
                    secondRegion.CellValue(1, 1, "South");
                    secondRegion.CellValue(2, 1, "West");

                    input.ValidationListRange("A1:A5", "$A$1:$A$2", "Region 1:Region 2");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet inputSheet = result.Workbook.Worksheets[0];
                LegacyXlsDataValidation validation = Assert.Single(inputSheet.DataValidations);
                Assert.Equal(LegacyXlsDataValidationType.List, validation.Type);
                Assert.Equal(LegacyXlsDataValidationListSourceKind.SheetQualifiedRange, validation.ListSourceKind);
                Assert.Equal("Region 1:Region 2", validation.ListSourceSheetName);
                Assert.Equal("A1:A2", validation.ListSourceRange);
                Assert.Equal("A1:A5", Assert.Single(validation.Ranges));

                ExcelDataValidationInfo projectedValidation = Assert.Single(result.Document.Sheets[0].GetDataValidations("A1:A5"));
                Assert.Equal("list", projectedValidation.Type);
                Assert.Equal("='Region 1:Region 2'!A1:A2", projectedValidation.Formula1);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalWorkbookFormulaDataValidations() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalDv");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(2, 1, 5d);

                    sheet.WorksheetPart.Worksheet!.Append(new DocumentFormat.OpenXml.Spreadsheet.DataValidations(
                        new DocumentFormat.OpenXml.Spreadsheet.DataValidation(
                            new OpenXmlFormula1("COUNTIF('[Other.xlsx]Data'!$A$1:$A$3,\">0\")>0")) {
                            Type = OpenXmlDataValidationValues.Custom,
                            SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A2:A2" }
                        }) { Count = 1U });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Other.xlsx", externalReference.Target);
                Assert.Equal(new[] { "Data" }, externalReference.SheetNames);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsDataValidation validation = Assert.Single(legacySheet.DataValidations);
                Assert.Equal(LegacyXlsDataValidationType.Custom, validation.Type);
                Assert.Equal("COUNTIF('[Other.xlsx]Data'!$A$1:$A$3,\">0\")>0", validation.Formula1);
                Assert.Equal("A2", Assert.Single(validation.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalDefinedNameDataValidations() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalDvName");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(2, 1, 5d);

                    sheet.WorksheetPart.Worksheet!.Append(new DocumentFormat.OpenXml.Spreadsheet.DataValidations(
                        new DocumentFormat.OpenXml.Spreadsheet.DataValidation(
                            new OpenXmlFormula1("COUNTIF([Other.xlsx]HasPositiveValues,\">0\")>0")) {
                            Type = OpenXmlDataValidationValues.Custom,
                            SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A2:A2" }
                        }) { Count = 1U });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Other.xlsx", externalReference.Target);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("HasPositiveValues", externalName.Name);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsDataValidation validation = Assert.Single(legacySheet.DataValidations);
                Assert.Equal("COUNTIF('Other.xlsx'!HasPositiveValues,\">0\")>0", validation.Formula1);
                Assert.Equal("A2", Assert.Single(validation.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetScopedExternalDefinedNameDataValidations() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalDvName");
                    sheet.CellValue(1, 1, "Amount");
                    sheet.CellValue(2, 1, 5d);

                    sheet.WorksheetPart.Worksheet!.Append(new DocumentFormat.OpenXml.Spreadsheet.DataValidations(
                        new DocumentFormat.OpenXml.Spreadsheet.DataValidation(
                            new OpenXmlFormula1("COUNTIF('[Other.xlsx]Feb'!HasPositiveValues,\">0\")>0")) {
                            Type = OpenXmlDataValidationValues.Custom,
                            SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A2:A2" }
                        }) { Count = 1U });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Other.xlsx", externalReference.Target);
                Assert.Equal(new[] { "Feb" }, externalReference.SheetNames);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("HasPositiveValues", externalName.Name);
                Assert.Equal(0, externalName.LocalSheetIndex);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsDataValidation validation = Assert.Single(legacySheet.DataValidations);
                Assert.Equal("COUNTIF('[Other.xlsx]Feb'!HasPositiveValues,\">0\")>0", validation.Formula1);
                Assert.Equal("A2", Assert.Single(validation.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedDataValidationFormulaPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation formula token payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Validated");
                string longLiteral = "\"" + new string('A', 255) + "\"";
                string formula = string.Join("&", Enumerable.Repeat(longLiteral, 260));

                var validation = new DocumentFormat.OpenXml.Spreadsheet.DataValidation(new OpenXmlFormula1(formula)) {
                    Type = OpenXmlDataValidationValues.Custom,
                    SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1" }
                };

                sheet.WorksheetPart.Worksheet!.Append(new DocumentFormat.OpenXml.Spreadsheet.DataValidations(validation) { Count = 1U });
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedDataValidationTextPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("data validation text payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Validated");
                sheet.ValidationCustomFormula("A1:A1", "A1<>\"\"", errorMessage: new string('A', 33000));
            });
        }
    }
}
