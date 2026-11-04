using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaEvaluator_PreservesRoundTripPrecisionInCachedNumbers() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.RoundTripPrecision.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Precision");
                sheet.Cell(1, 1).SetFormula("1/7");

                Assert.Equal(1, document.Calculate());
                string cached = Assert.Single(sheet.InspectFormulas().Formulas).CachedValue!;
                Assert.Equal(1d / 7d, double.Parse(cached, System.Globalization.CultureInfo.InvariantCulture));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesIndexAndExactMatch() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IndexMatch.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(2, 1).SetValue("EU");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(3, 1).SetValue("US");
                sheet.Cell(3, 2).SetValue(20d);
                sheet.Cell(4, 1).SetValue("APAC");
                sheet.Cell(4, 2).SetValue(30d);
                sheet.Cell(6, 1).SetFormula("MATCH(\"US\",A2:A4,0)");
                sheet.Cell(7, 1).SetFormula("INDEX(B2:B4,MATCH(\"APAC\",A2:A4,0))");
                sheet.Cell(8, 1).SetFormula("INDEX(A2:B4,2,1)");
                sheet.Cell(9, 1).SetFormula("INDEX(A1:B1,2)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(4, before.TotalFormulas);
                Assert.Equal(4, before.SupportedFormulas);
                Assert.Contains("INDEX", before.Capabilities.SupportedFunctions);
                Assert.Contains("MATCH", before.Capabilities.SupportedFunctions);

                Assert.Equal(4, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A6" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A7" && formula.CachedValue == "30");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A8" && formula.CachedValue == "US");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A9" && formula.CachedValue == "Amount");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTypedLookupShapeErrors() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IndexMatchUnsupported.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue("EU");
                sheet.Cell(1, 2).SetValue("US");
                sheet.Cell(2, 1).SetValue(10d);
                sheet.Cell(2, 2).SetValue(20d);
                sheet.Cell(4, 1).SetFormula("INDEX(A1:B2,2)");
                sheet.Cell(5, 1).SetFormula("MATCH(\"US\",A1:B2,0)");
                sheet.Cell(6, 1).SetFormula("MATCH(\"ZZ\",A1:B1,1)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();

                Assert.Equal(3, inspection.TotalFormulas);
                Assert.Equal(3, inspection.SupportedFormulas);
                Assert.True(inspection.Formulas.Single(formula => formula.CellReference == "A4").IsSupportedByOfficeIMO);
                Assert.Equal(3, document.Calculate());
                AssertFormulaCache(sheet, 4, 1, "#REF!");
                AssertFormulaCache(sheet, 5, 1, "#N/A");
                AssertFormulaCache(sheet, 6, 1, "2");
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesApproximateLookupReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ApproximateLookup.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Lookup");
                sheet.Cell(1, 1).SetValue(10d);
                sheet.Cell(2, 1).SetValue(20d);
                sheet.Cell(3, 1).SetValue(30d);
                sheet.Cell(4, 1).SetValue(40d);
                sheet.Cell(1, 2).SetValue("Low");
                sheet.Cell(2, 2).SetValue("Mid");
                sheet.Cell(3, 2).SetValue("High");
                sheet.Cell(4, 2).SetValue("Top");
                sheet.Cell(1, 3).SetValue("East");
                sheet.Cell(2, 3).SetValue("West");
                sheet.Cell(3, 3).SetValue("East");
                sheet.Cell(4, 3).SetValue("North");
                sheet.Cell(1, 4).SetValue(40d);
                sheet.Cell(2, 4).SetValue(30d);
                sheet.Cell(3, 4).SetValue(20d);
                sheet.Cell(4, 4).SetValue(10d);

                sheet.Cell(1, 5).SetFormula("MATCH(25,A1:A4,1)");
                sheet.Cell(2, 5).SetFormula("MATCH(25,D1:D4,-1)");
                sheet.Cell(3, 5).SetFormula("MATCH(30,A1:A4)");
                sheet.Cell(4, 5).SetFormula("XMATCH(\"East\",C1:C4,0,-1)");
                sheet.Cell(5, 5).SetFormula("XMATCH(25,A1:A4,-1)");
                sheet.Cell(6, 5).SetFormula("XMATCH(25,A1:A4,1)");
                sheet.Cell(7, 5).SetFormula("XLOOKUP(25,A1:A4,B1:B4,\"Missing\",-1)");
                sheet.Cell(8, 5).SetFormula("XLOOKUP(25,A1:A4,B1:B4,\"Missing\",1)");
                sheet.Cell(9, 5).SetFormula("XMATCH(25,A1:A4,2)");
                sheet.Cell(10, 5).SetFormula("MATCH(25,A1:B4,1)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(10, before.SupportedFormulas);
                Assert.Contains("MATCH", before.Capabilities.SupportedFunctions);
                Assert.Contains("XMATCH", before.Capabilities.SupportedFunctions);
                Assert.Contains("XLOOKUP", before.Capabilities.SupportedFunctions);
                Assert.All(before.Formulas, formula => Assert.True(formula.IsSupportedByOfficeIMO));

                Assert.Equal(10, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E5" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E6" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E7" && formula.CachedValue == "Mid");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E8" && formula.CachedValue == "High");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E9" && formula.CachedValue == "#N/A");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E10" && formula.CachedValue == "#N/A");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesConditionalAggregatesWithCriteriaCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ConditionalAggregates.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Quarter");
                sheet.Cell(1, 3).SetValue("Amount");
                sheet.Cell(2, 1).SetValue("East");
                sheet.Cell(2, 2).SetValue("Q1");
                sheet.Cell(2, 3).SetValue(10d);
                sheet.Cell(3, 1).SetValue("West");
                sheet.Cell(3, 2).SetValue("Q1");
                sheet.Cell(3, 3).SetValue(15d);
                sheet.Cell(4, 1).SetValue("East");
                sheet.Cell(4, 2).SetValue("Q2");
                sheet.Cell(4, 3).SetValue(20d);
                sheet.Cell(5, 1).SetValue("East");
                sheet.Cell(5, 2).SetValue("Q1");
                sheet.Cell(5, 3).SetValue(30d);

                sheet.Cell(1, 5).SetValue("East");
                sheet.Cell(2, 5).SetValue("Q1");
                sheet.Cell(3, 5).SetValue(">15");

                sheet.Cell(1, 6).SetFormula("COUNTIF(A2:A5,E1)");
                sheet.Cell(2, 6).SetFormula("SUMIF(A2:A5,E1,C2:C5)");
                sheet.Cell(3, 6).SetFormula("AVERAGEIF(A2:A5,E1,C2:C5)");
                sheet.Cell(4, 6).SetFormula("COUNTIFS(A2:A5,E1,B2:B5,E2)");
                sheet.Cell(5, 6).SetFormula("SUMIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.Cell(6, 6).SetFormula("AVERAGEIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.Cell(7, 6).SetFormula("SUMIFS(C2:C5,A2:A5,E1,C2:C5,E3)");
                sheet.Cell(8, 6).SetFormula("COUNTIF(A2:A5,\"E*\")");
                sheet.Cell(9, 6).SetFormula("MINIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.Cell(10, 6).SetFormula("MAXIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.Cell(11, 6).SetFormula("MINIFS(C2:C5,A2:A5,\"North\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(11, before.TotalFormulas);
                Assert.Equal(11, before.SupportedFormulas);
                Assert.Contains("SUMIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains("AVERAGEIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains("MINIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains("MAXIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "F11" && formula.IsSupportedByOfficeIMO);

                Assert.Equal(11, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F1" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F2" && formula.CachedValue == "60");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F3" && formula.CachedValue == "20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F4" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F5" && formula.CachedValue == "40");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F6" && formula.CachedValue == "20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F7" && formula.CachedValue == "50");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F8" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F9" && formula.CachedValue == "10");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F10" && formula.CachedValue == "30");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "F11" && formula.CachedValue == "0");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesXLookupFallbackAndSearchMode() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.XLookupFallback.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(1, 3).SetValue("Owner");
                sheet.Cell(2, 1).SetValue("East");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(2, 3).SetValue("Alice");
                sheet.Cell(3, 1).SetValue("West");
                sheet.Cell(3, 2).SetValue(20d);
                sheet.Cell(3, 3).SetValue("Bob");
                sheet.Cell(4, 1).SetValue("East");
                sheet.Cell(4, 2).SetValue(30d);
                sheet.Cell(4, 3).SetValue("Ann");

                sheet.Cell(1, 5).SetFormula("XLOOKUP(\"West\",A2:A4,B2:B4,\"Missing\",0)");
                sheet.Cell(2, 5).SetFormula("XLOOKUP(\"North\",A2:A4,C2:C4,\"Missing\",0)");
                sheet.Cell(3, 5).SetFormula("XLOOKUP(\"East\",A2:A4,C2:C4,\"Missing\",0,-1)");
                sheet.Cell(4, 5).SetFormula("XLOOKUP(\"North\",A2:A4,B2:B4,0,0)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(4, before.TotalFormulas);
                Assert.Equal(4, before.SupportedFormulas);

                Assert.Equal(4, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "Missing");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "Ann");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "0");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesValueReturningIfAndIfError() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IfTextResults.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(1, 3).SetValue("Owner");
                sheet.Cell(2, 1).SetValue("East");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(2, 3).SetValue("Alice");
                sheet.Cell(3, 1).SetValue("West");
                sheet.Cell(3, 2).SetValue(20d);
                sheet.Cell(3, 3).SetValue("Bob");
                sheet.Cell(4, 1).SetValue("East");
                sheet.Cell(4, 2).SetValue(30d);
                sheet.Cell(4, 3).SetValue("Ann");

                sheet.Cell(1, 5).SetFormula("IF(A2=\"East\",\"Priority\",\"Standard\")");
                sheet.Cell(2, 5).SetFormula("IF(A3<>\"East\",\"Other\",\"Priority\")");
                sheet.Cell(3, 5).SetFormula("IF(B2>15,\"High\",\"Low\")");
                sheet.Cell(4, 5).SetFormula("IFERROR(XLOOKUP(\"North\",A2:A4,C2:C4),\"Missing\")");
                sheet.Cell(5, 5).SetFormula("IFERROR(XLOOKUP(\"West\",A2:A4,C2:C4),\"Missing\")");
                sheet.Cell(6, 5).SetFormula("IF(A2=A4,\"Same\",\"Different\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(6, before.TotalFormulas);
                Assert.Equal(6, before.SupportedFormulas);

                Assert.Equal(6, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "Priority");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "Other");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "Low");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "Missing");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E5" && formula.CachedValue == "Bob");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E6" && formula.CachedValue == "Same");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextFormattingFunction() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextFunction.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Summary");
                sheet.Cell(1, 1).SetValue("East");
                sheet.Cell(1, 2).SetValue(1234.5d);
                sheet.Cell(1, 3).SetValue(0.257d);

                sheet.Cell(1, 5).SetFormula("TEXT(B1,\"#,##0.00\")");
                sheet.Cell(2, 5).SetFormula("TEXT(C1,\"0.0%\")");
                sheet.Cell(3, 5).SetFormula("TEXT(DATE(2026,5,28),\"yyyy-mm-dd\")");
                sheet.Cell(4, 5).SetFormula("TEXT(DATE(2026,5,28),\"mmm yyyy\")");
                sheet.Cell(5, 5).SetFormula("TEXT(TIME(9,5,0),\"hh:mm\")");
                sheet.Cell(6, 5).SetFormula("TEXT(A1,\"@\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(6, before.TotalFormulas);
                Assert.Equal(6, before.SupportedFormulas);
                Assert.Contains("TEXT", before.Capabilities.SupportedFunctions);

                Assert.Equal(6, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "1,234.50");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "25.7%");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "2026-05-28");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "May 2026");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E5" && formula.CachedValue == "09:05");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E6" && formula.CachedValue == "East");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextCaseReportLabels() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextCaseFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Labels");
                sheet.Cell(1, 1).SetValue("north region");
                sheet.Cell(2, 1).SetValue("Q1-sales review");
                sheet.Cell(3, 1).SetValue("  MIXED   spacing  ");

                sheet.Cell(1, 3).SetFormula("UPPER(A1)");
                sheet.Cell(2, 3).SetFormula("LOWER(\"READY\")");
                sheet.Cell(3, 3).SetFormula("PROPER(A2)");
                sheet.Cell(4, 3).SetFormula("PROPER(TRIM(A3))");
                sheet.Cell(5, 3).SetFormula("CONCAT(UPPER(LEFT(A1,1)),RIGHT(A1,5))");
                sheet.Cell(6, 3).SetFormula("PROPER(A1,A2)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(6, before.TotalFormulas);
                Assert.Equal(5, before.SupportedFormulas);
                Assert.Contains("UPPER", before.Capabilities.SupportedFunctions);
                Assert.Contains("LOWER", before.Capabilities.SupportedFunctions);
                Assert.Contains("PROPER", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "C6" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(5, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "NORTH REGION");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "ready");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C3" && formula.CachedValue == "Q1-Sales Review");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C4" && formula.CachedValue == "Mixed Spacing");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C5" && formula.CachedValue == "Negion");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextCleanupReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextCleanupFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Labels");
                sheet.Cell(1, 1).SetValue("North - Q1 - Draft");
                sheet.Cell(2, 1).SetValue("1,234.50");
                sheet.Cell(3, 1).SetValue("west region");

                sheet.Cell(1, 3).SetFormula("SUBSTITUTE(A1,\" - \",\" / \")");
                sheet.Cell(2, 3).SetFormula("SUBSTITUTE(A1,\" - \",\" | \",2)");
                sheet.Cell(3, 3).SetFormula("FIND(\"Q1\",A1)");
                sheet.Cell(4, 3).SetFormula("SEARCH(\"REGION\",A3)");
                sheet.Cell(5, 3).SetFormula("VALUE(A2)");
                sheet.Cell(6, 3).SetFormula("CONCATENATE(LEFT(A3,4),\":\",VALUE(\"42\"))");
                sheet.Cell(7, 3).SetFormula("VALUE(SUBSTITUTE(\"$1,234\",\"$\",\"\"))");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(7, before.TotalFormulas);
                Assert.Equal(7, before.SupportedFormulas);
                Assert.Contains("SUBSTITUTE", before.Capabilities.SupportedFunctions);
                Assert.Contains("FIND", before.Capabilities.SupportedFunctions);
                Assert.Contains("SEARCH", before.Capabilities.SupportedFunctions);
                Assert.Contains("VALUE", before.Capabilities.SupportedFunctions);
                Assert.Contains("CONCATENATE", before.Capabilities.SupportedFunctions);

                Assert.Equal(7, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "North / Q1 / Draft");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "North - Q1 | Draft");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C3" && formula.CachedValue == "9");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C4" && formula.CachedValue == "6");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C5" && formula.CachedValue == "1234.5");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C6" && formula.CachedValue == "west:42");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C7" && formula.CachedValue == "1234");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextBeforeAfterReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextBeforeAfterFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Labels");
                sheet.Cell(1, 1).SetValue("North - Q1 - Draft");
                sheet.Cell(2, 1).SetValue("Region=West|Status=Open");
                sheet.Cell(3, 1).SetValue("alpha-BETA-gamma");

                sheet.Cell(1, 3).SetFormula("TEXTBEFORE(A1,\" - \")");
                sheet.Cell(2, 3).SetFormula("TEXTAFTER(A1,\" - \")");
                sheet.Cell(3, 3).SetFormula("TEXTBEFORE(A1,\" - \",2)");
                sheet.Cell(4, 3).SetFormula("TEXTAFTER(A1,\" - \",2)");
                sheet.Cell(5, 3).SetFormula("TEXTBEFORE(A1,\" - \",-1)");
                sheet.Cell(6, 3).SetFormula("TEXTAFTER(A1,\" - \",-1)");
                sheet.Cell(7, 3).SetFormula("TEXTBEFORE(A3,\"beta\",1,1)");
                sheet.Cell(8, 3).SetFormula("TEXTAFTER(A3,\"beta\",1,1)");
                sheet.Cell(9, 3).SetFormula("TEXTBEFORE(A2,\";\",1,0,FALSE,\"Missing\")");
                sheet.Cell(10, 3).SetFormula("TEXTBEFORE(A2,\";\",1,0,TRUE)");
                sheet.Cell(11, 3).SetFormula("TEXTAFTER(A2,\";\",1,0,FALSE,\"Missing\")");
                sheet.Cell(12, 3).SetFormula("TEXTBEFORE(A1,\"\")");
                sheet.Cell(13, 3).SetFormula("TEXTAFTER(A1,\" - \",0)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(13, before.TotalFormulas);
                Assert.Equal(11, before.SupportedFormulas);
                Assert.Contains("TEXTBEFORE", before.Capabilities.SupportedFunctions);
                Assert.Contains("TEXTAFTER", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "C12" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "C13" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(11, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "North");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "Q1 - Draft");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C3" && formula.CachedValue == "North - Q1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C4" && formula.CachedValue == "Draft");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C5" && formula.CachedValue == "North - Q1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C6" && formula.CachedValue == "Draft");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C7" && formula.CachedValue == "alpha-");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C8" && formula.CachedValue == "-gamma");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C9" && formula.CachedValue == "Missing");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C10" && formula.CachedValue == "Region=West|Status=Open");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C11" && formula.CachedValue == "Missing");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextJoinExactReptAndSumProductReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextJoinExactReptSumProduct.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Report");
                sheet.Cell(1, 1).SetValue(1d);
                sheet.Cell(2, 1).SetValue(2d);
                sheet.Cell(3, 1).SetValue(3d);
                sheet.Cell(1, 2).SetValue(10d);
                sheet.Cell(2, 2).SetValue(20d);
                sheet.Cell(3, 2).SetValue(30d);
                sheet.Cell(1, 3).SetValue("North");
                sheet.Cell(2, 3).SetValue(string.Empty);
                sheet.Cell(3, 3).SetValue("West");
                sheet.Cell(4, 3).SetValue("East");

                sheet.Cell(1, 5).SetFormula("TEXTJOIN(\"/\",TRUE,C1:C4)");
                sheet.Cell(2, 5).SetFormula("TEXTJOIN(\"-\",FALSE,C1:C3)");
                sheet.Cell(3, 5).SetFormula("EXACT(C1,\"North\")");
                sheet.Cell(4, 5).SetFormula("EXACT(C1,\"north\")");
                sheet.Cell(5, 5).SetFormula("REPT(\"*\",3)");
                sheet.Cell(6, 5).SetFormula("CONCAT(REPT(\"0\",2),TEXT(A1,\"0\"))");
                sheet.Cell(7, 5).SetFormula("SUMPRODUCT(A1:A3,B1:B3)");
                sheet.Cell(8, 5).SetFormula("SUMPRODUCT(A1:A3)");
                sheet.Cell(9, 5).SetFormula("SUMPRODUCT(A1:A3,B1:B2)");
                sheet.Cell(10, 5).SetFormula("REPT(\"x\",-1)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(9, before.SupportedFormulas);
                Assert.Contains("TEXTJOIN", before.Capabilities.SupportedFunctions);
                Assert.Contains("EXACT", before.Capabilities.SupportedFunctions);
                Assert.Contains("REPT", before.Capabilities.SupportedFunctions);
                Assert.Contains("SUMPRODUCT", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "E9" && formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "E10" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(9, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "North/West/East");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "North--West");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E5" && formula.CachedValue == "***");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E6" && formula.CachedValue == "001");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E7" && formula.CachedValue == "140");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E8" && formula.CachedValue == "6");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E9" && formula.CachedValue == "#VALUE!");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesDateAndTimeReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.DateTimeReportFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Schedule");
                sheet.Cell(1, 1).SetValue(new DateTime(2026, 5, 28, 9, 5, 7));
                sheet.Cell(1, 2).SetValue(new DateTime(2026, 5, 27));

                sheet.Cell(1, 4).SetFormula("YEAR(A1)");
                sheet.Cell(2, 4).SetFormula("MONTH(A1)");
                sheet.Cell(3, 4).SetFormula("DAY(A1)");
                sheet.Cell(4, 4).SetFormula("HOUR(A1)");
                sheet.Cell(5, 4).SetFormula("MINUTE(A1)");
                sheet.Cell(6, 4).SetFormula("SECOND(A1)");
                sheet.Cell(7, 4).SetFormula("TEXT(EDATE(A1,1),\"yyyy-mm-dd\")");
                sheet.Cell(8, 4).SetFormula("TEXT(EOMONTH(A1,0),\"yyyy-mm-dd\")");
                sheet.Cell(9, 4).SetFormula("DAYS(DATE(2026,6,11),DATE(2026,5,28))");
                sheet.Cell(10, 4).SetFormula("WEEKDAY(DATE(2026,5,31),2)");
                sheet.Cell(11, 4).SetFormula("NETWORKDAYS(DATE(2026,5,25),DATE(2026,5,29),B1:B1)");
                sheet.Cell(12, 4).SetFormula("TEXT(TIME(23,75,0),\"hh:mm\")");
                sheet.Cell(13, 4).SetFormula("TEXT(DATEVALUE(\"2026-05-28\"),\"yyyy-mm-dd\")");
                sheet.Cell(14, 4).SetFormula("YEAR(DATEVALUE(\"May 28 2026\"))");
                sheet.Cell(15, 4).SetFormula("TEXT(DATEVALUE(\"5/28/2026\"),\"yyyy-mm-dd\")");
                sheet.Cell(16, 4).SetFormula("TEXT(TIMEVALUE(\"09:05:07\"),\"hh:mm\")");
                sheet.Cell(17, 4).SetFormula("SECOND(TIMEVALUE(\"09:05:07\"))");
                sheet.Cell(18, 4).SetFormula("TEXT(DATEVALUE(TEXT(A1,\"yyyy-mm-dd\")),\"yyyy-mm-dd\")");
                sheet.Cell(19, 4).SetFormula("DATEVALUE(\"not a date\")");
                sheet.Cell(20, 4).SetFormula("DATEVALUE(\"2026-05-28\")");
                sheet.Cell(21, 4).SetFormula("TIMEVALUE(\"09:05:07\")");
                sheet.Cell(22, 4).SetFormula("WEEKNUM(A1)");
                sheet.Cell(23, 4).SetFormula("WEEKNUM(A1,2)");
                sheet.Cell(24, 4).SetFormula("ISOWEEKNUM(DATE(2021,1,1))");
                sheet.Cell(25, 4).SetFormula("WEEKNUM(DATE(2021,1,1),21)");
                sheet.Cell(26, 4).SetFormula("DAYS360(DATE(2026,2,28),DATE(2026,3,31))");
                sheet.Cell(27, 4).SetFormula("DAYS360(DATE(2026,2,28),DATE(2026,3,31),TRUE)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(27, before.TotalFormulas);
                Assert.Equal(26, before.SupportedFormulas);
                Assert.Contains("EDATE", before.Capabilities.SupportedFunctions);
                Assert.Contains("EOMONTH", before.Capabilities.SupportedFunctions);
                Assert.Contains("NETWORKDAYS", before.Capabilities.SupportedFunctions);
                Assert.Contains("DATEVALUE", before.Capabilities.SupportedFunctions);
                Assert.Contains("TIMEVALUE", before.Capabilities.SupportedFunctions);
                Assert.Contains("WEEKNUM", before.Capabilities.SupportedFunctions);
                Assert.Contains("ISOWEEKNUM", before.Capabilities.SupportedFunctions);
                Assert.Contains("DAYS360", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D19" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(26, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D1" && formula.CachedValue == "2026");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D2" && formula.CachedValue == "5");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D3" && formula.CachedValue == "28");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D4" && formula.CachedValue == "9");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D5" && formula.CachedValue == "5");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D6" && formula.CachedValue == "7");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D7" && formula.CachedValue == "2026-06-28");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D8" && formula.CachedValue == "2026-05-31");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D9" && formula.CachedValue == "14");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D10" && formula.CachedValue == "7");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D11" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D12" && formula.CachedValue == "00:15");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D13" && formula.CachedValue == "2026-05-28");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D14" && formula.CachedValue == "2026");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D15" && formula.CachedValue == "2026-05-28");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D16" && formula.CachedValue == "09:05");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D17" && formula.CachedValue == "7");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D18" && formula.CachedValue == "2026-05-28");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D19" && string.IsNullOrEmpty(formula.CachedValue));
                AssertCachedNumber(after, "D20", new DateTime(2026, 5, 28).ToOADate());
                AssertCachedNumber(after, "D21", new TimeSpan(9, 5, 7).TotalDays);
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D22" && formula.CachedValue == "22");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D23" && formula.CachedValue == "22");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D24" && formula.CachedValue == "53");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D25" && formula.CachedValue == "53");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D26" && formula.CachedValue == "30");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D27" && formula.CachedValue == "32");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesDateDifReportUnits() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.DateDifFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Schedule");
                sheet.Cell(1, 1).SetValue(new DateTime(2020, 1, 15));
                sheet.Cell(1, 2).SetValue(new DateTime(2026, 5, 28));

                sheet.Cell(1, 4).SetFormula("DATEDIF(A1,B1,\"Y\")");
                sheet.Cell(2, 4).SetFormula("DATEDIF(A1,B1,\"M\")");
                sheet.Cell(3, 4).SetFormula("DATEDIF(A1,B1,\"D\")");
                sheet.Cell(4, 4).SetFormula("DATEDIF(A1,B1,\"YM\")");
                sheet.Cell(5, 4).SetFormula("DATEDIF(A1,B1,\"YD\")");
                sheet.Cell(6, 4).SetFormula("DATEDIF(A1,B1,\"MD\")");
                sheet.Cell(7, 4).SetFormula("DATEDIF(B1,A1,\"D\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(7, before.TotalFormulas);
                Assert.Equal(7, before.SupportedFormulas);
                Assert.Contains("DATEDIF", before.Capabilities.SupportedFunctions);

                Assert.Equal(7, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D1" && formula.CachedValue == "6");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D2" && formula.CachedValue == "76");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D3" && formula.CachedValue == "2325");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D4" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D5" && formula.CachedValue == "133");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D6" && formula.CachedValue == "13");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D7" && formula.CachedValue == "#NUM!");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesYearFracReportBases() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.YearFracFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Schedule");
                sheet.Cell(1, 1).SetValue(new DateTime(2024, 1, 1));
                sheet.Cell(1, 2).SetValue(new DateTime(2025, 1, 1));
                sheet.Cell(2, 1).SetValue(new DateTime(2024, 1, 31));
                sheet.Cell(2, 2).SetValue(new DateTime(2024, 2, 29));

                sheet.Cell(1, 4).SetFormula("YEARFRAC(A1,B1)");
                sheet.Cell(2, 4).SetFormula("YEARFRAC(A1,B1,1)");
                sheet.Cell(3, 4).SetFormula("YEARFRAC(A1,B1,2)");
                sheet.Cell(4, 4).SetFormula("YEARFRAC(A1,B1,3)");
                sheet.Cell(5, 4).SetFormula("YEARFRAC(A2,B2,0)");
                sheet.Cell(6, 4).SetFormula("YEARFRAC(A2,B2,4)");
                sheet.Cell(7, 4).SetFormula("YEARFRAC(B1,A1,1)");
                sheet.Cell(8, 4).SetFormula("YEARFRAC(A1,B1,5)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(8, before.TotalFormulas);
                Assert.Equal(7, before.SupportedFormulas);
                Assert.Contains("YEARFRAC", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D8" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(7, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                AssertCachedNumber(after, "D1", 1d);
                AssertCachedNumber(after, "D2", 1d);
                AssertCachedNumber(after, "D3", 366d / 360d);
                AssertCachedNumber(after, "D4", 366d / 365d);
                AssertCachedNumber(after, "D5", 29d / 360d);
                AssertCachedNumber(after, "D6", 29d / 360d);
                AssertCachedNumber(after, "D7", 1d);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesFinancialReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.FinancialFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Finance");
                sheet.Cell(1, 1).SetValue(0.05d / 12d);
                sheet.Cell(2, 1).SetValue(360d);
                sheet.Cell(3, 1).SetValue(250000d);
                sheet.Cell(6, 1).SetValue(100d);
                sheet.Cell(7, 1).SetValue(200d);
                sheet.Cell(8, 1).SetValue(300d);

                sheet.Cell(1, 4).SetFormula("PMT(A1,A2,A3)");
                sheet.Cell(2, 4).SetFormula("PV(A1,A2,PMT(A1,A2,A3))");
                sheet.Cell(3, 4).SetFormula("FV(A1,60,-200,0)");
                sheet.Cell(4, 4).SetFormula("NPER(A1,PMT(A1,A2,A3),A3)");
                sheet.Cell(5, 4).SetFormula("NPV(0.1,A6:A8)");
                sheet.Cell(6, 4).SetFormula("PMT(0,12,1200)");
                sheet.Cell(7, 4).SetFormula("PV(0,12,-100)");
                sheet.Cell(8, 4).SetFormula("FV(0,12,-100)");
                sheet.Cell(9, 4).SetFormula("PMT(A1,A2,A3,0,1)");
                sheet.Cell(10, 4).SetFormula("PMT(A1,A2,A3,0,2)");
                sheet.Cell(11, 4).SetFormula("NPER(A1,0,A3)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(11, before.TotalFormulas);
                Assert.Equal(10, before.SupportedFormulas);
                Assert.Contains("PMT", before.Capabilities.SupportedFunctions);
                Assert.Contains("PV", before.Capabilities.SupportedFunctions);
                Assert.Contains("FV", before.Capabilities.SupportedFunctions);
                Assert.Contains("NPER", before.Capabilities.SupportedFunctions);
                Assert.Contains("NPV", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D10" && formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D11" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(10, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                AssertCachedNumber(after, "D1", -1342.05405753035d, precision: 8);
                AssertCachedNumber(after, "D2", 250000d, precision: 8);
                AssertCachedNumber(after, "D3", 13601.2165681686d, precision: 8);
                AssertCachedNumber(after, "D4", 360d, precision: 8);
                AssertCachedNumber(after, "D5", 481.592787377911d, precision: 8);
                AssertCachedNumber(after, "D6", -100d);
                AssertCachedNumber(after, "D7", 1200d);
                AssertCachedNumber(after, "D8", 1200d);
                AssertCachedNumber(after, "D9", -1336.48536849495d, precision: 8);
                AssertCachedNumber(after, "D10", -1336.48536849495d, precision: 8);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_RejectsOversizedCovarianceRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.OversizedCovarianceRange.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Stats");
                document.AddSheet("Data");
                sheet.Cell(1, 1).SetFormula("COVARIANCE.P(Data!A1:XFD1048576,Data!A1:XFD1048576)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();
                Assert.Equal(1, inspection.TotalFormulas);
                Assert.Equal(0, inspection.SupportedFormulas);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && !formula.IsSupportedByOfficeIMO);
                Assert.Equal(0, document.Calculate());
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_RejectsCumulativeOversizedFormulaRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.CumulativeOversizedRanges.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Stats");
                document.AddSheet("Data");
                sheet.Cell(1, 1).SetFormula("COVARIANCE.P(Data!A1:A100000,Data!B1:B100000)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();
                Assert.Equal(1, inspection.TotalFormulas);
                Assert.Equal(0, inspection.SupportedFormulas);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && !formula.IsSupportedByOfficeIMO);
                Assert.Equal(0, document.Calculate());
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_RejectsOversizedDirectLookupRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.OversizedLookupRange.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Lookup");
                document.AddSheet("Data");
                sheet.Cell(1, 1).SetFormula("VLOOKUP(\"missing\",Data!A1:A100001,1,FALSE)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();
                Assert.Equal(1, inspection.TotalFormulas);
                Assert.Equal(0, inspection.SupportedFormulas);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && !formula.IsSupportedByOfficeIMO);
                Assert.Equal(0, document.Calculate());
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesStatisticalReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.StatisticalFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Stats");
                sheet.Cell(1, 1).SetValue(10d);
                sheet.Cell(2, 1).SetValue(12d);
                sheet.Cell(3, 1).SetValue(14d);
                sheet.Cell(4, 1).SetValue(18d);
                sheet.Cell(5, 1).SetValue(21d);
                sheet.Cell(1, 2).SetValue(10d);
                sheet.Cell(2, 2).SetValue(12d);
                sheet.Cell(3, 2).SetValue(14d);
                sheet.Cell(4, 2).SetValue(14d);
                sheet.Cell(5, 2).SetValue(21d);
                sheet.Cell(1, 3).SetValue(1d);
                sheet.Cell(2, 3).SetValue(2d);
                sheet.Cell(3, 3).SetValue(3d);
                sheet.Cell(4, 3).SetValue(4d);
                sheet.Cell(5, 3).SetValue(5d);
                sheet.Cell(6, 1).SetValue("Text");

                sheet.Cell(1, 4).SetFormula("VAR.P(A1:A5)");
                sheet.Cell(2, 4).SetFormula("VAR.S(A1:A5)");
                sheet.Cell(3, 4).SetFormula("STDEV.P(A1:A5)");
                sheet.Cell(4, 4).SetFormula("STDEV.S(A1:A5)");
                sheet.Cell(5, 4).SetFormula("PERCENTILE.INC(A1:A5,0)");
                sheet.Cell(6, 4).SetFormula("PERCENTILE.INC(A1:A5,0.25)");
                sheet.Cell(7, 4).SetFormula("PERCENTILE.INC(A1:A5,0.9)");
                sheet.Cell(8, 4).SetFormula("QUARTILE.INC(A1:A5,2)");
                sheet.Cell(9, 4).SetFormula("QUARTILE.INC(A1:A5,4)");
                sheet.Cell(10, 4).SetFormula("PERCENTILE.INC(A1:A5,1.5)");
                sheet.Cell(11, 4).SetFormula("QUARTILE.INC(A1:A5,5)");
                sheet.Cell(12, 4).SetFormula("STDEV.S(A1)");
                sheet.Cell(13, 4).SetFormula("SUMSQ(A1:A5)");
                sheet.Cell(14, 4).SetFormula("RANK.EQ(18,A1:A5,0)");
                sheet.Cell(15, 4).SetFormula("RANK.EQ(A3,A1:A5,1)");
                sheet.Cell(16, 4).SetFormula("RANK.EQ(15,A1:A5)");
                sheet.Cell(17, 4).SetFormula("RANK.AVG(14,B1:B5,0)");
                sheet.Cell(18, 4).SetFormula("RANK.AVG(14,B1:B5,1)");
                sheet.Cell(19, 4).SetFormula("PERCENTRANK.INC(A1:A5,18)");
                sheet.Cell(20, 4).SetFormula("PERCENTRANK.INC(A1:A5,16)");
                sheet.Cell(21, 4).SetFormula("PERCENTRANK.INC(A1:A5,5)");
                sheet.Cell(22, 4).SetFormula("PERCENTILE.EXC(A1:A5,0.25)");
                sheet.Cell(23, 4).SetFormula("QUARTILE.EXC(A1:A5,3)");
                sheet.Cell(24, 4).SetFormula("PERCENTRANK.EXC(A1:A5,18)");
                sheet.Cell(25, 4).SetFormula("PERCENTRANK.EXC(A1:A5,16)");
                sheet.Cell(26, 4).SetFormula("PERCENTILE.EXC(A1:A5,0.99)");
                sheet.Cell(27, 4).SetFormula("CORREL(A1:A5,C1:C5)");
                sheet.Cell(28, 4).SetFormula("SLOPE(A1:A5,C1:C5)");
                sheet.Cell(29, 4).SetFormula("INTERCEPT(A1:A5,C1:C5)");
                sheet.Cell(30, 4).SetFormula("RSQ(A1:A5,C1:C5)");
                sheet.Cell(31, 4).SetFormula("FORECAST.LINEAR(6,A1:A5,C1:C5)");
                sheet.Cell(32, 4).SetFormula("SLOPE(A1:A5,C1:C4)");
                sheet.Cell(33, 4).SetFormula("MEDIAN(A1:A5)");
                sheet.Cell(34, 4).SetFormula("LARGE(A1:A5,2)");
                sheet.Cell(35, 4).SetFormula("SMALL(A1:A5,3)");
                sheet.Cell(36, 4).SetFormula("MODE.SNGL(B1:B5)");
                sheet.Cell(37, 4).SetFormula("MODE(A1:A5)");
                sheet.Cell(38, 4).SetFormula("GEOMEAN(A1:A5)");
                sheet.Cell(39, 4).SetFormula("HARMEAN(A1:A5)");
                sheet.Cell(40, 4).SetFormula("AVERAGEA(A1:A6)");
                sheet.Cell(41, 4).SetFormula("MINA(A1:A6)");
                sheet.Cell(42, 4).SetFormula("MAXA(A1:A6)");
                sheet.Cell(43, 4).SetFormula("COUNTA(A1:A6)");
                sheet.Cell(44, 4).SetFormula("AVERAGEA(\"skip\",TRUE,FALSE,5)");
                sheet.Cell(45, 4).SetFormula("AVEDEV(A1:A5)");
                sheet.Cell(46, 4).SetFormula("DEVSQ(A1:A5)");
                sheet.Cell(47, 4).SetFormula("SUMXMY2(A1:A5,C1:C5)");
                sheet.Cell(48, 4).SetFormula("SUMX2MY2(A1:A5,C1:C5)");
                sheet.Cell(49, 4).SetFormula("SUMX2PY2(A1:A5,C1:C5)");
                sheet.Cell(50, 4).SetFormula("SUMXMY2(A1:A5,C1:C4)");
                sheet.Cell(51, 4).SetFormula("COVARIANCE.P(A1:A5,C1:C5)");
                sheet.Cell(52, 4).SetFormula("COVARIANCE.S(A1:A5,C1:C5)");
                sheet.Cell(53, 4).SetFormula("COVAR(A1:A5,C1:C5)");
                sheet.Cell(54, 4).SetFormula("COVARIANCE.P(A1:A5,C1:C4)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(54, before.TotalFormulas);
                Assert.Equal(44, before.SupportedFormulas);
                Assert.Contains("AVERAGEA", before.Capabilities.SupportedFunctions);
                Assert.Contains("MINA", before.Capabilities.SupportedFunctions);
                Assert.Contains("MAXA", before.Capabilities.SupportedFunctions);
                Assert.Contains("COUNTA", before.Capabilities.SupportedFunctions);
                Assert.Contains("MEDIAN", before.Capabilities.SupportedFunctions);
                Assert.Contains("LARGE", before.Capabilities.SupportedFunctions);
                Assert.Contains("SMALL", before.Capabilities.SupportedFunctions);
                Assert.Contains("MODE.SNGL", before.Capabilities.SupportedFunctions);
                Assert.Contains("MODE", before.Capabilities.SupportedFunctions);
                Assert.Contains("GEOMEAN", before.Capabilities.SupportedFunctions);
                Assert.Contains("HARMEAN", before.Capabilities.SupportedFunctions);
                Assert.Contains("AVEDEV", before.Capabilities.SupportedFunctions);
                Assert.Contains("DEVSQ", before.Capabilities.SupportedFunctions);
                Assert.Contains("SUMXMY2", before.Capabilities.SupportedFunctions);
                Assert.Contains("SUMX2MY2", before.Capabilities.SupportedFunctions);
                Assert.Contains("SUMX2PY2", before.Capabilities.SupportedFunctions);
                Assert.Contains("COVARIANCE.P", before.Capabilities.SupportedFunctions);
                Assert.Contains("COVARIANCE.S", before.Capabilities.SupportedFunctions);
                Assert.Contains("COVAR", before.Capabilities.SupportedFunctions);
                Assert.Contains("SUMSQ", before.Capabilities.SupportedFunctions);
                Assert.Contains("VAR.P", before.Capabilities.SupportedFunctions);
                Assert.Contains("VAR.S", before.Capabilities.SupportedFunctions);
                Assert.Contains("STDEV.P", before.Capabilities.SupportedFunctions);
                Assert.Contains("STDEV.S", before.Capabilities.SupportedFunctions);
                Assert.Contains("PERCENTILE.INC", before.Capabilities.SupportedFunctions);
                Assert.Contains("PERCENTILE.EXC", before.Capabilities.SupportedFunctions);
                Assert.Contains("QUARTILE.INC", before.Capabilities.SupportedFunctions);
                Assert.Contains("QUARTILE.EXC", before.Capabilities.SupportedFunctions);
                Assert.Contains("PERCENTRANK.INC", before.Capabilities.SupportedFunctions);
                Assert.Contains("PERCENTRANK.EXC", before.Capabilities.SupportedFunctions);
                Assert.Contains("RANK.EQ", before.Capabilities.SupportedFunctions);
                Assert.Contains("RANK.AVG", before.Capabilities.SupportedFunctions);
                Assert.Contains("CORREL", before.Capabilities.SupportedFunctions);
                Assert.Contains("SLOPE", before.Capabilities.SupportedFunctions);
                Assert.Contains("INTERCEPT", before.Capabilities.SupportedFunctions);
                Assert.Contains("RSQ", before.Capabilities.SupportedFunctions);
                Assert.Contains("FORECAST.LINEAR", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D10" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D11" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D12" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D16" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D21" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D26" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D32" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D37" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D50" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D54" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(44, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                AssertCachedNumber(after, "D1", 16d);
                AssertCachedNumber(after, "D2", 20d);
                AssertCachedNumber(after, "D3", 4d);
                AssertCachedNumber(after, "D4", Math.Sqrt(20d), precision: 8);
                AssertCachedNumber(after, "D5", 10d);
                AssertCachedNumber(after, "D6", 12d);
                AssertCachedNumber(after, "D7", 19.8d);
                AssertCachedNumber(after, "D8", 14d);
                AssertCachedNumber(after, "D9", 21d);
                AssertCachedNumber(after, "D13", 1205d);
                AssertCachedNumber(after, "D14", 2d);
                AssertCachedNumber(after, "D15", 3d);
                AssertCachedNumber(after, "D17", 2.5d);
                AssertCachedNumber(after, "D18", 3.5d);
                AssertCachedNumber(after, "D19", 0.75d);
                AssertCachedNumber(after, "D20", 0.625d);
                AssertCachedNumber(after, "D22", 11d);
                AssertCachedNumber(after, "D23", 19.5d);
                AssertCachedNumber(after, "D24", 4d / 6d);
                AssertCachedNumber(after, "D25", 3.5d / 6d);
                AssertCachedNumber(after, "D27", 0.9899494936611666d, precision: 8);
                AssertCachedNumber(after, "D28", 2.8d);
                AssertCachedNumber(after, "D29", 6.6d);
                AssertCachedNumber(after, "D30", 0.98d);
                AssertCachedNumber(after, "D31", 23.4d);
                AssertCachedNumber(after, "D33", 14d);
                AssertCachedNumber(after, "D34", 18d);
                AssertCachedNumber(after, "D35", 14d);
                AssertCachedNumber(after, "D36", 14d);
                AssertCachedNumber(after, "D38", Math.Pow(10d * 12d * 14d * 18d * 21d, 1d / 5d), precision: 8);
                AssertCachedNumber(after, "D39", 5d / ((1d / 10d) + (1d / 12d) + (1d / 14d) + (1d / 18d) + (1d / 21d)), precision: 8);
                AssertCachedNumber(after, "D40", 12.5d);
                AssertCachedNumber(after, "D41", 0d);
                AssertCachedNumber(after, "D42", 21d);
                AssertCachedNumber(after, "D43", 6d);
                AssertCachedNumber(after, "D44", 1.5d);
                AssertCachedNumber(after, "D45", 3.6d);
                AssertCachedNumber(after, "D46", 80d);
                AssertCachedNumber(after, "D47", 754d);
                AssertCachedNumber(after, "D48", 1150d);
                AssertCachedNumber(after, "D49", 1260d);
                AssertCachedNumber(after, "D51", 5.6d);
                AssertCachedNumber(after, "D52", 7d);
                AssertCachedNumber(after, "D53", 5.6d);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesRoundingReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.RoundingFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Rounding");
                sheet.Cell(1, 1).SetFormula("MROUND(23,5)");
                sheet.Cell(2, 1).SetFormula("MROUND(-23,-5)");
                sheet.Cell(3, 1).SetFormula("CEILING.MATH(23,5)");
                sheet.Cell(4, 1).SetFormula("FLOOR.MATH(23,5)");
                sheet.Cell(5, 1).SetFormula("CEILING.MATH(-23,5)");
                sheet.Cell(6, 1).SetFormula("CEILING.MATH(-23,5,1)");
                sheet.Cell(7, 1).SetFormula("FLOOR.MATH(-23,5)");
                sheet.Cell(8, 1).SetFormula("FLOOR.MATH(-23,5,1)");
                sheet.Cell(9, 1).SetFormula("MROUND(23,-5)");
                sheet.Cell(10, 1).SetFormula("ROUND(23.456,2)");
                sheet.Cell(11, 1).SetFormula("ROUND(1234,-2)");
                sheet.Cell(12, 1).SetFormula("ROUNDUP(23.451,2)");
                sheet.Cell(13, 1).SetFormula("ROUNDUP(1234,-2)");
                sheet.Cell(14, 1).SetFormula("ROUNDDOWN(23.459,2)");
                sheet.Cell(15, 1).SetFormula("ROUNDDOWN(1234,-2)");
                sheet.Cell(16, 1).SetFormula("TRUNC(23.459,2)");
                sheet.Cell(17, 1).SetFormula("TRUNC(1234,-2)");
                sheet.Cell(18, 1).SetFormula("INT(-23.1)");
                sheet.Cell(19, 1).SetFormula("CEILING(23,5)");
                sheet.Cell(20, 1).SetFormula("FLOOR(23,5)");
                sheet.Cell(21, 1).SetFormula("ROUND(23,16)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(21, before.TotalFormulas);
                Assert.Equal(19, before.SupportedFormulas);
                Assert.Contains("ROUND", before.Capabilities.SupportedFunctions);
                Assert.Contains("ROUNDUP", before.Capabilities.SupportedFunctions);
                Assert.Contains("ROUNDDOWN", before.Capabilities.SupportedFunctions);
                Assert.Contains("TRUNC", before.Capabilities.SupportedFunctions);
                Assert.Contains("INT", before.Capabilities.SupportedFunctions);
                Assert.Contains("CEILING", before.Capabilities.SupportedFunctions);
                Assert.Contains("FLOOR", before.Capabilities.SupportedFunctions);
                Assert.Contains("MROUND", before.Capabilities.SupportedFunctions);
                Assert.Contains("CEILING.MATH", before.Capabilities.SupportedFunctions);
                Assert.Contains("FLOOR.MATH", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "A9" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "A21" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(19, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A1" && formula.CachedValue == "25");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A2" && formula.CachedValue == "-25");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A3" && formula.CachedValue == "25");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A4" && formula.CachedValue == "20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A5" && formula.CachedValue == "-20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A6" && formula.CachedValue == "-25");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A7" && formula.CachedValue == "-25");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A8" && formula.CachedValue == "-20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A10" && formula.CachedValue == "23.46");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A11" && formula.CachedValue == "1200");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A12" && formula.CachedValue == "23.46");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A13" && formula.CachedValue == "1300");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A14" && formula.CachedValue == "23.45");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A15" && formula.CachedValue == "1200");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A16" && formula.CachedValue == "23.45");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A17" && formula.CachedValue == "1200");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A18" && formula.CachedValue == "-24");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A19" && formula.CachedValue == "25");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "A20" && formula.CachedValue == "20");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesChooseReportSelectors() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ChooseSelectors.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Selectors");
                sheet.Cell(1, 1).SetValue(2d);
                sheet.Cell(2, 1).SetValue("EU");
                sheet.Cell(3, 1).SetValue("US");
                sheet.Cell(4, 1).SetValue("APAC");

                sheet.Cell(1, 3).SetFormula("CHOOSE(2,10,20,30)");
                sheet.Cell(2, 3).SetFormula("CHOOSE(A1,\"Draft\",\"Final\",\"Archived\")");
                sheet.Cell(3, 3).SetFormula("CHOOSE(MATCH(\"US\",A2:A4,0),100,200,300)");
                sheet.Cell(4, 3).SetFormula("TEXT(CHOOSE(3,0.1,0.2,0.3),\"0%\")");
                sheet.Cell(5, 3).SetFormula("CHOOSE(4,10,20,30)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(5, before.TotalFormulas);
                Assert.Equal(4, before.SupportedFormulas);
                Assert.Contains("CHOOSE", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "C5" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(4, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "Final");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C3" && formula.CachedValue == "200");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C4" && formula.CachedValue == "30%");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesInfoReportGuards() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.InfoGuards.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Info");
                sheet.Cell(1, 1).SetValue(42d);
                sheet.Cell(3, 1).SetValue("Ready");

                sheet.Cell(1, 3).SetFormula("ISBLANK(A2)");
                sheet.Cell(2, 3).SetFormula("ISNUMBER(A1)");
                sheet.Cell(3, 3).SetFormula("ISTEXT(A3)");
                sheet.Cell(4, 3).SetFormula("IF(ISBLANK(A2),\"Missing\",\"Present\")");
                sheet.Cell(5, 3).SetFormula("IF(ISNUMBER(A1),A1*2,0)");
                sheet.Cell(6, 3).SetFormula("ISBLANK(\"\")");
                sheet.Cell(7, 3).SetFormula("ISNUMBER(A3)");
                sheet.Cell(8, 3).SetFormula("ISTEXT(A2)");
                sheet.Cell(9, 3).SetFormula("ISBLANK(A2,A3)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(9, before.TotalFormulas);
                Assert.Equal(8, before.SupportedFormulas);
                Assert.Contains("ISBLANK", before.Capabilities.SupportedFunctions);
                Assert.Contains("ISNUMBER", before.Capabilities.SupportedFunctions);
                Assert.Contains("ISTEXT", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "C9" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(8, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C3" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C4" && formula.CachedValue == "Missing");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C5" && formula.CachedValue == "84");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C6" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C7" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C8" && formula.CachedValue == "0");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesErrorInfoReportGuards() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ErrorInfoGuards.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Info");
                sheet.Cell(1, 1).SetValue("East");
                sheet.Cell(2, 1).SetValue("West");
                sheet.Cell(1, 2).SetValue(10d);
                sheet.Cell(2, 2).SetValue(0d);

                sheet.Cell(1, 4).SetFormula("ISERROR(XLOOKUP(\"North\",A1:A2,B1:B2))");
                sheet.Cell(2, 4).SetFormula("ISNA(XLOOKUP(\"North\",A1:A2,B1:B2))");
                sheet.Cell(3, 4).SetFormula("ISERR(FIND(\"z\",\"abc\"))");
                sheet.Cell(4, 4).SetFormula("ISERROR(FIND(\"a\",\"abc\"))");
                sheet.Cell(5, 4).SetFormula("IF(ISERROR(XLOOKUP(\"North\",A1:A2,B1:B2)),\"Missing\",\"Found\")");
                sheet.Cell(6, 4).SetFormula("IF(ISNA(#N/A),\"NA\",\"Other\")");
                sheet.Cell(7, 4).SetFormula("IF(ISERR(#DIV/0!),\"Err\",\"Other\")");
                sheet.Cell(8, 4).SetFormula("ISERROR(B1/B2)");
                sheet.Cell(9, 4).SetFormula("ISNA(FIND(\"z\",\"abc\"))");
                sheet.Cell(10, 4).SetFormula("ISERR(#N/A)");
                sheet.Cell(11, 4).SetFormula("ISERROR(\"text\",A1)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(11, before.TotalFormulas);
                Assert.Equal(10, before.SupportedFormulas);
                Assert.Contains("ISERROR", before.Capabilities.SupportedFunctions);
                Assert.Contains("ISERR", before.Capabilities.SupportedFunctions);
                Assert.Contains("ISNA", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D11" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(10, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D1" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D2" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D3" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D4" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D5" && formula.CachedValue == "Missing");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D6" && formula.CachedValue == "NA");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D7" && formula.CachedValue == "Err");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D8" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D9" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D10" && formula.CachedValue == "0");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesIfNaReportFallbacks() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IfNaFallbacks.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Info");
                sheet.Cell(1, 1).SetValue("East");
                sheet.Cell(2, 1).SetValue("West");
                sheet.Cell(1, 2).SetValue(10d);
                sheet.Cell(2, 2).SetValue(20d);

                sheet.Cell(1, 4).SetFormula("IFNA(XLOOKUP(\"North\",A1:A2,B1:B2),\"Missing\")");
                sheet.Cell(2, 4).SetFormula("IFNA(XLOOKUP(\"East\",A1:A2,B1:B2),\"Missing\")");
                sheet.Cell(3, 4).SetFormula("IFNA(#N/A,\"NA\")");
                sheet.Cell(4, 4).SetFormula("IFNA(#DIV/0!,\"NA\")");
                sheet.Cell(5, 4).SetFormula("IFERROR(#DIV/0!,\"Any\")");
                sheet.Cell(6, 4).SetFormula("IFERROR(#N/A,\"Any\")");
                sheet.Cell(7, 4).SetFormula("IFNA(FIND(\"z\",\"abc\"),\"NA\")");
                sheet.Cell(8, 4).SetFormula("IFNA(\"ready\",\"NA\")");
                sheet.Cell(9, 4).SetFormula("IFNA(XLOOKUP(\"North\",A1:A2,B1:B2),A1)");
                sheet.Cell(10, 4).SetFormula("IFNA(\"text\",A1,B1)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(9, before.SupportedFormulas);
                Assert.Contains("IFNA", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D10" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(9, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D1" && formula.CachedValue == "Missing");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D2" && formula.CachedValue == "10");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D3" && formula.CachedValue == "NA");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D4" && formula.CachedValue == "#DIV/0!");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D5" && formula.CachedValue == "Any");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D6" && formula.CachedValue == "Any");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D7" && formula.CachedValue == "#VALUE!");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D8" && formula.CachedValue == "ready");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D9" && formula.CachedValue == "East");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesFormulaInspectionReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.FormulaInspectionFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddSheet("Source Data");
                source.Cell(1, 1).SetValue(2d);
                source.Cell(1, 2).SetValue(3d);
                source.Cell(1, 3).SetFormula("A1+B1");

                ExcelSheet sheet = document.AddSheet("Info");
                sheet.Cell(1, 1).SetValue(10d);
                sheet.Cell(1, 2).SetValue(5d);
                sheet.Cell(1, 3).SetFormula("A1+B1");
                sheet.Cell(2, 3).SetFormula("IF(A1>3,\"High\",\"Low\")");

                sheet.Cell(1, 4).SetFormula("ISFORMULA(C1)");
                sheet.Cell(2, 4).SetFormula("ISFORMULA(A1)");
                sheet.Cell(3, 4).SetFormula("FORMULATEXT(C1)");
                sheet.Cell(4, 4).SetFormula("FORMULATEXT(C2)");
                sheet.Cell(5, 4).SetFormula("IF(ISFORMULA(C2),FORMULATEXT(C2),\"Missing\")");
                sheet.Cell(6, 4).SetFormula("LEN(FORMULATEXT(C1))");
                sheet.Cell(7, 4).SetFormula("ISFORMULA('Source Data'!C1)");
                sheet.Cell(8, 4).SetFormula("FORMULATEXT('Source Data'!C1)");
                sheet.Cell(9, 4).SetFormula("FORMULATEXT(A1)");
                sheet.Cell(10, 4).SetFormula("ISFORMULA(A1:B1)");

                ExcelFormulaInspection before = document.InspectFormulas();
                Assert.Equal(13, before.TotalFormulas);
                Assert.Equal(11, before.SupportedFormulas);
                Assert.Contains("ISFORMULA", before.Capabilities.SupportedFunctions);
                Assert.Contains("FORMULATEXT", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D9" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D10" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(11, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D1" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D2" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D3" && formula.CachedValue == "=A1+B1");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D4" && formula.CachedValue == "=IF(A1>3,\"High\",\"Low\")");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D5" && formula.CachedValue == "=IF(A1>3,\"High\",\"Low\")");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D6" && formula.CachedValue == "6");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D7" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Info" && formula.CellReference == "D8" && formula.CachedValue == "=A1+B1");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesReferenceShapeReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ReferenceShapeFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddSheet("Source Data");

                ExcelSheet sheet = document.AddSheet("Reference");
                sheet.Cell(1, 4).SetFormula("ROW(A5)");
                sheet.Cell(2, 4).SetFormula("COLUMN(C1)");
                sheet.Cell(3, 4).SetFormula("ROWS(A1:B3)");
                sheet.Cell(4, 4).SetFormula("COLUMNS(A1:B3)");
                sheet.Cell(5, 4).SetFormula("ROW('Source Data'!D5)");
                sheet.Cell(6, 4).SetFormula("COLUMN('Source Data'!D5)");
                sheet.Cell(7, 4).SetFormula("ROWS('Source Data'!B2:D5)");
                sheet.Cell(8, 4).SetFormula("COLUMNS('Source Data'!B2:D5)");
                sheet.Cell(9, 4).SetFormula("ROW()");
                sheet.Cell(10, 4).SetFormula("ROWS(A1,B1)");

                ExcelFormulaInspection before = document.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(9, before.SupportedFormulas);
                Assert.Contains("ROW", before.Capabilities.SupportedFunctions);
                Assert.Contains("COLUMN", before.Capabilities.SupportedFunctions);
                Assert.Contains("ROWS", before.Capabilities.SupportedFunctions);
                Assert.Contains("COLUMNS", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D9" && formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D10" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(9, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D1" && formula.CachedValue == "5");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D2" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D3" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D4" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D5" && formula.CachedValue == "5");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D6" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D7" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D8" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D9" && formula.CachedValue == "9");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesWorkdaySchedulingFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.WorkdayFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Schedule");
                sheet.Cell(1, 1).SetValue(new DateTime(2026, 5, 28));
                sheet.Cell(1, 2).SetValue(new DateTime(2026, 5, 29));
                sheet.Cell(2, 2).SetValue(new DateTime(2026, 6, 1));

                sheet.Cell(1, 4).SetFormula("TEXT(WORKDAY(A1,1),\"yyyy-mm-dd\")");
                sheet.Cell(2, 4).SetFormula("TEXT(WORKDAY(A1,1,B1:B1),\"yyyy-mm-dd\")");
                sheet.Cell(3, 4).SetFormula("TEXT(WORKDAY(A1,1,B1:B2),\"yyyy-mm-dd\")");
                sheet.Cell(4, 4).SetFormula("TEXT(WORKDAY(A1,-3),\"yyyy-mm-dd\")");
                sheet.Cell(5, 4).SetFormula("TEXT(WORKDAY.INTL(A1,2,\"0000110\"),\"yyyy-mm-dd\")");
                sheet.Cell(6, 4).SetFormula("TEXT(WORKDAY.INTL(A1,1,11),\"yyyy-mm-dd\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(6, before.TotalFormulas);
                Assert.Equal(6, before.SupportedFormulas);
                Assert.Contains("WORKDAY", before.Capabilities.SupportedFunctions);
                Assert.Contains("WORKDAY.INTL", before.Capabilities.SupportedFunctions);

                Assert.Equal(6, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D1" && formula.CachedValue == "2026-05-29");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D2" && formula.CachedValue == "2026-06-01");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D3" && formula.CachedValue == "2026-06-02");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D4" && formula.CachedValue == "2026-05-25");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D5" && formula.CachedValue == "2026-06-01");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D6" && formula.CachedValue == "2026-05-29");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesSubtotalAndCountBlankReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.SubtotalCountBlank.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Audit");
                sheet.Cell(1, 1).SetValue(10d);
                sheet.Cell(2, 1).SetValue(20d);
                sheet.Cell(4, 1).SetValue("North");
                sheet.Cell(5, 1).SetValue(string.Empty);
                sheet.Cell(6, 1).SetFormula("A1+A2");

                sheet.Cell(1, 3).SetFormula("COUNTBLANK(A1:A6)");
                sheet.Cell(2, 3).SetFormula("SUBTOTAL(9,A1:A6)");
                sheet.Cell(3, 3).SetFormula("SUBTOTAL(1,A1:A6)");
                sheet.Cell(4, 3).SetFormula("SUBTOTAL(2,A1:A6)");
                sheet.Cell(5, 3).SetFormula("SUBTOTAL(3,A1:A6)");
                sheet.Cell(6, 3).SetFormula("SUBTOTAL(4,A1:A6)");
                sheet.Cell(7, 3).SetFormula("SUBTOTAL(5,A1:A6)");
                sheet.Cell(8, 3).SetFormula("SUBTOTAL(109,A1:A2,A6:A6)");
                sheet.Cell(9, 3).SetFormula("SUBTOTAL(6,A1:A6)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(9, before.SupportedFormulas);
                Assert.Contains("COUNTBLANK", before.Capabilities.SupportedFunctions);
                Assert.Contains("SUBTOTAL", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "C9" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(9, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "60");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C3" && formula.CachedValue == "20");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C4" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C5" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C6" && formula.CachedValue == "30");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C7" && formula.CachedValue == "10");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "C8" && formula.CachedValue == "60");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        private static void AssertCachedNumber(ExcelFormulaInspection inspection, string cellReference, double expected, int precision = 10) {
            ExcelFormulaCellInfo formula = Assert.Single(inspection.Formulas, item => item.CellReference == cellReference);
            Assert.NotNull(formula.CachedValue);
            Assert.True(double.TryParse(formula.CachedValue, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double actual));
            Assert.Equal(expected, actual, precision);
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesIfsAndSwitchReportBranches() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IfsSwitch.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(1, 3).SetValue("Owner");
                sheet.Cell(2, 1).SetValue("East");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(2, 3).SetValue("Alice");
                sheet.Cell(3, 1).SetValue("West");
                sheet.Cell(3, 2).SetValue(20d);
                sheet.Cell(3, 3).SetValue("Bob");
                sheet.Cell(4, 1).SetValue("North");
                sheet.Cell(4, 2).SetValue(5d);
                sheet.Cell(4, 3).SetValue("Nina");

                sheet.Cell(1, 5).SetFormula("IFS(B2>=20,\"High\",B2>=10,\"Medium\",TRUE,\"Low\")");
                sheet.Cell(2, 5).SetFormula("IFS(B4>=20,\"High\",B4>=10,\"Medium\",TRUE,\"Low\")");
                sheet.Cell(3, 5).SetFormula("SWITCH(A3,\"East\",\"Priority\",\"West\",\"Standard\",\"Other\")");
                sheet.Cell(4, 5).SetFormula("SWITCH(XLOOKUP(\"East\",A2:A4,C2:C4),\"Alice\",\"Owner A\",\"Ann\",\"Owner Ann\",\"Other\")");
                sheet.Cell(5, 5).SetFormula("SWITCH(A4,\"East\",1,\"West\",2,0)");
                sheet.Cell(6, 5).SetFormula("IFS(A2=\"East\",TEXT(B2,\"0.0\"),TRUE,\"Missing\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(6, before.TotalFormulas);
                Assert.Equal(6, before.SupportedFormulas);
                Assert.Contains("IFS", before.Capabilities.SupportedFunctions);
                Assert.Contains("SWITCH", before.Capabilities.SupportedFunctions);

                Assert.Equal(6, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "Medium");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "Low");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "Standard");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "Owner A");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E5" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E6" && formula.CachedValue == "10.0");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaInspection_ReportsDependenciesAndIssues() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.DependencyDiagnostics.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue(10d);
                sheet.Cell(1, 2).SetFormula("A1+5");
                sheet.Cell(1, 3).SetFormula("B1+1");
                sheet.Cell(1, 4).SetFormula("D1+1");
                sheet.Cell(1, 5).SetFormula("UNSUPPORTED(B1)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();

                Assert.True(inspection.HasDependencyIssues);
                Assert.True(inspection.DependencyIssueCount >= 3);

                ExcelFormulaCellInfo dependent = Assert.Single(inspection.Formulas, formula => formula.CellReference == "C1");
                Assert.Contains("Sales!B1", dependent.Dependencies);
                Assert.Contains(dependent.DependencyIssues, issue => issue.Contains("Sales!B1", System.StringComparison.OrdinalIgnoreCase)
                    && issue.Contains("without a cached result", System.StringComparison.OrdinalIgnoreCase));

                ExcelFormulaCellInfo circular = Assert.Single(inspection.Formulas, formula => formula.CellReference == "D1");
                Assert.Contains("Sales!D1", circular.Dependencies);
                Assert.Contains(circular.DependencyIssues, issue => issue.Contains("own formula cell", System.StringComparison.OrdinalIgnoreCase));

                ExcelFormulaCellInfo unsupportedDependency = Assert.Single(inspection.Formulas, formula => formula.CellReference == "E1");
                Assert.Contains("Sales!B1", unsupportedDependency.Dependencies);

                var exception = Assert.Throws<System.InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
                Assert.Contains("Formula dependency issues", exception.Message);
            }
        }

        [Fact]
        public void Test_FormulaInspection_BuildsWorkbookDependencyGraph() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.DependencyGraph.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sales = document.AddSheet("Sales");
                sales.Cell(1, 1).SetValue(10d);
                sales.Cell(1, 2).SetFormula("A1+5");
                sales.Cell(1, 3).SetFormula("B1+1");

                ExcelSheet summary = document.AddSheet("Summary");
                summary.Cell(1, 1).SetFormula("Sales!C1*2");

                ExcelFormulaInspection inspection = document.InspectFormulas();
                ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;

                Assert.Equal(3, graph.NodeCount);

                ExcelFormulaDependencyNode? salesB1 = graph.FindNode("Sales", "B1");
                Assert.NotNull(salesB1);
                Assert.Contains("Sales!A1", salesB1!.Dependencies);
                Assert.Contains("Sales!C1", salesB1.Dependents);

                ExcelFormulaDependencyNode? salesC1 = graph.FindNode("Sales", "C1");
                Assert.NotNull(salesC1);
                Assert.Contains("Sales!B1", salesC1!.Dependencies);
                Assert.Contains("Summary!A1", salesC1.Dependents);

                ExcelFormulaDependencyNode? summaryA1 = graph.FindNode("Summary", "A1");
                Assert.NotNull(summaryA1);
                Assert.Contains("Sales!C1", summaryA1!.Dependencies);
                Assert.Empty(summaryA1.Dependents);

                string markdown = graph.ToMarkdown();
                Assert.Contains("Sales!B1", markdown);
                Assert.Contains("Summary!A1", markdown);
            }
        }
    }
}
