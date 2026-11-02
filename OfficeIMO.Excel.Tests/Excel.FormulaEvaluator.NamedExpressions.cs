using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("numeric_expression", "Calc+Nested", "511")]
        [InlineData("index_reference", "SUM(Picked)", "170")]
        [InlineData("offset_reference", "SUM(Shifted)", "40")]
        [InlineData("local_scope", "Rate+'Data'!Rate", "12")]
        [InlineData("function_arguments", "Guarded", "11")]
        [InlineData("text_expression", "Textual", "Rate=7")]
        [InlineData("named_array", "SUM(Generated)", "21")]
        [InlineData("nested_named_array", "SUM(FILTER(Generated,SEQUENCE(3)>1))", "18")]
        public void FormulaEvaluator_NamedExpressionsResolveTypedValuesAndReferences(string caseId, string formula, string expected) {
            using var document = ExcelDocument.Create();
            var data = CreateLookupModeSheet(document);
            var report = document.AddSheet("Report");
            AddFormulaName(document, "Calc", "SUM(Data!$A$1:$A$6)*2");
            AddFormulaName(document, "Nested", "Calc/2+1");
            AddFormulaName(document, "Picked", "INDEX(Data!$A$1:$B$6,0,1)");
            AddFormulaName(document, "Shifted", "OFFSET(Data!$A$1,1,0,2,1)");
            AddFormulaName(document, "Rate", "2");
            AddFormulaName(document, "Rate", "5", 0);
            AddFormulaName(document, "Rate", "7", 1);
            AddFormulaName(document, "Guarded", "IFERROR(INDEX(Data!$A$1:$A$6,99),11)");
            AddFormulaName(document, "Textual", "CONCAT(\"Rate=\",Rate)");
            AddFormulaName(document, "Generated", "SEQUENCE(3,2)");
            report.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(report, 1, 1, expected);
            Assert.True(Assert.Single(report.InspectFormulas().Formulas).IsSupportedByOfficeIMO, caseId);
        }

        [Fact]
        public void FormulaEvaluator_NamedExpressionsRefreshAndPreserveSavedCaches() {
            using var document = ExcelDocument.Create();
            var data = CreateLookupModeSheet(document);
            var report = document.AddSheet("Report");
            AddFormulaName(document, "Total", "SUM(Data!$A$1:$A$6)");
            report.CellFormula(1, 1, "Total*2");
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(report, 1, 1, "340");
            data.CellValue(6, 1, 60d);
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(report, 1, 1, "360");
            var name = Assert.Single(document.OpenXmlDocument.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>());
            name.Text = "SUM(Data!$A$1:$A$6)+1";
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(report, 1, 1, "362");
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            AssertFormulaCache(reopened.Sheets[1], 1, 1, "362");
            Assert.Equal(1, reopened.Calculate());
            AssertFormulaCache(reopened.Sheets[1], 1, 1, "362");
        }

        [Fact]
        public void FormulaEvaluator_NamedFormulaAuthoringCalculatesScopesAndRoundTrips() {
            using var document = ExcelDocument.Create();
            var data = CreateLookupModeSheet(document);
            var report = document.AddSheet("Report");
            document.SetNamedFormula("Total", "=0", save: false);
            document.SetNamedFormula("tOtAl", "=SUM(Data!$A$1:$A$6)*2", save: false);
            Assert.Equal("SUM(Data!$A$1:$A$6)*2", document.GetNamedRange("TOTAL"));
            data.SetNamedFormula("Rate", "5", save: false, hidden: true);
            report.SetNamedFormula("Rate", "7", save: false);
            report.CellFormula(1, 1, "Total+Rate+Data!Rate");
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(report, 1, 1, "352");
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(3, reopened.OpenXmlDocument.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().Count());
            AssertFormulaCache(reopened.Sheets[1], 1, 1, "352");
            Assert.Equal(1, reopened.Calculate());
            AssertFormulaCache(reopened.Sheets[1], 1, 1, "352");
        }

        [Fact]
        public void FormulaEvaluator_NameChangesInvalidateCachePreflightAndPersistDirtyState() {
            using var document = ExcelDocument.Create();
            var source = document.AddSheet("Data");
            source.CellValue(1, 1, 3d);
            source.CellValue(2, 1, 6d);
            var report = document.AddSheet("Report");
            document.SetNamedFormula("Rate", "1", save: false);
            document.SetNamedRange("Chosen", "Data!A1", save: false);
            report.CellFormula(1, 1, "Rate*2");
            report.CellFormula(1, 2, "Chosen");
            Assert.Equal(2, document.Calculate());
            Assert.All(report.InspectFormulas().Formulas, formula => Assert.False(formula.IsDirty));
            Assert.True(document.InspectFeatures().Can(ExcelPreflightCapability.UseCachedFormulaValues));

            document.SetNamedFormula("Rate", "3", save: false);
            AssertFormulaCache(report, 1, 1, "2");
            Assert.All(report.InspectFormulas().Formulas, formula => Assert.True(formula.IsDirty));
            Assert.False(document.InspectFeatures().Can(ExcelPreflightCapability.UseCachedFormulaValues));
            Assert.False(document.InspectFeatures().Can(ExcelPreflightCapability.ExportPdfReport));
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            var reopenedReport = reopened.Sheets[1];
            Assert.All(reopenedReport.InspectFormulas().Formulas, formula => Assert.True(formula.IsDirty));
            Assert.Equal(2, reopened.Calculate());
            AssertFormulaCache(reopenedReport, 1, 1, "6");
            Assert.All(reopenedReport.InspectFormulas().Formulas, formula => Assert.False(formula.IsDirty));
            reopened.SetNamedRange("Chosen", "Data!A2", save: false);
            Assert.True(reopenedReport.InspectFormulas().Formulas.Single(formula => formula.CellReference == "B1").IsDirty);
            Assert.Equal(2, reopened.Calculate());
            AssertFormulaCache(reopenedReport, 1, 2, "6");
            Assert.True(reopened.RemoveNamedRange("cHoSeN", save: false));
            Assert.True(reopenedReport.InspectFormulas().Formulas.Single(formula => formula.CellReference == "B1").IsDirty);
            Assert.False(reopened.InspectFeatures().Can(ExcelPreflightCapability.UseCachedFormulaValues));
        }

        [Fact]
        public void FormulaEvaluator_NamedArraySpillsRefreshTheirShape() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.CellValue(1, 1, 2d);
            document.SetNamedFormula("Series", "SEQUENCE(Data!A1,2)", save: false);
            sheet.CellFormula(1, 3, "SUM(Series)");
            sheet.SetDynamicArrayFormula("D1", "Series");
            Assert.Equal(2, document.Calculate());
            AssertFormulaCache(sheet, 1, 3, "10");
            Assert.Equal(4d, sheet.CellAt(2, 5).GetValue().Value);
            sheet.CellValue(1, 1, 3d);
            Assert.Equal(2, document.Calculate());
            AssertFormulaCache(sheet, 1, 3, "21");
            Assert.Equal(6d, sheet.CellAt(3, 5).GetValue().Value);
            sheet.CellValue(1, 1, 1d);
            Assert.Equal(2, document.Calculate());
            AssertFormulaCache(sheet, 1, 3, "3");
            Assert.Null(sheet.CellAt(3, 5).GetValue().Value);
        }

        [Fact]
        public void FormulaEvaluator_CyclicNamedExpressionsStayUncalculatedWithoutBlockingOtherCells() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Names");
            AddFormulaName(document, "CycleLeft", "CycleRight+1");
            AddFormulaName(document, "CycleRight", "CycleLeft+1");
            sheet.CellValue(1, 1, 42d);
            sheet.CellFormula(1, 1, "CycleLeft");
            sheet.CellFormula(2, 1, "2+5");
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 1, "42");
            AssertFormulaCache(sheet, 2, 1, "7");
            Assert.False(sheet.InspectFormulas().Formulas.Single(formula => formula.CellReference == "A1").IsSupportedByOfficeIMO);
        }

        private static void AddFormulaName(ExcelDocument document, string name, string formula, uint? localSheet = null) {
            var names = document.OpenXmlDocument.WorkbookPart!.Workbook.DefinedNames ??= new DefinedNames();
            var definition = new DefinedName(formula) { Name = name };
            if (localSheet.HasValue) definition.LocalSheetId = localSheet.Value;
            names.Append(definition);
        }
    }
}
