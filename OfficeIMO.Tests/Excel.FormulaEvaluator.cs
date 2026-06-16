using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaEvaluator_CalculatesIndexAndExactMatch() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IndexMatch.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, "EU");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "US");
                sheet.CellValue(3, 2, 20d);
                sheet.CellValue(4, 1, "APAC");
                sheet.CellValue(4, 2, 30d);
                sheet.CellFormula(6, 1, "MATCH(\"US\",A2:A4,0)");
                sheet.CellFormula(7, 1, "INDEX(B2:B4,MATCH(\"APAC\",A2:A4,0))");
                sheet.CellFormula(8, 1, "INDEX(A2:B4,2,1)");
                sheet.CellFormula(9, 1, "INDEX(A1:B1,2)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_ReportsUnsupportedIndexMatchShapes() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IndexMatchUnsupported.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "EU");
                sheet.CellValue(1, 2, "US");
                sheet.CellValue(2, 1, 10d);
                sheet.CellValue(2, 2, 20d);
                sheet.CellFormula(4, 1, "INDEX(A1:B2,2)");
                sheet.CellFormula(5, 1, "MATCH(\"US\",A1:B2,0)");
                sheet.CellFormula(6, 1, "MATCH(\"ZZ\",A1:B1,1)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();

                Assert.Equal(3, inspection.TotalFormulas);
                Assert.Equal(0, inspection.SupportedFormulas);
                Assert.All(inspection.Formulas, formula => Assert.Contains("supported function", formula.UnsupportedReason));
                Assert.Equal(0, document.Calculate());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesApproximateLookupReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ApproximateLookup.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Lookup");
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(2, 1, 20d);
                sheet.CellValue(3, 1, 30d);
                sheet.CellValue(4, 1, 40d);
                sheet.CellValue(1, 2, "Low");
                sheet.CellValue(2, 2, "Mid");
                sheet.CellValue(3, 2, "High");
                sheet.CellValue(4, 2, "Top");
                sheet.CellValue(1, 3, "East");
                sheet.CellValue(2, 3, "West");
                sheet.CellValue(3, 3, "East");
                sheet.CellValue(4, 3, "North");

                sheet.CellFormula(1, 5, "MATCH(25,A1:A4,1)");
                sheet.CellFormula(2, 5, "MATCH(25,A1:A4,-1)");
                sheet.CellFormula(3, 5, "MATCH(30,A1:A4)");
                sheet.CellFormula(4, 5, "XMATCH(\"East\",C1:C4,0,-1)");
                sheet.CellFormula(5, 5, "XMATCH(25,A1:A4,-1)");
                sheet.CellFormula(6, 5, "XMATCH(25,A1:A4,1)");
                sheet.CellFormula(7, 5, "XLOOKUP(25,A1:A4,B1:B4,\"Missing\",-1)");
                sheet.CellFormula(8, 5, "XLOOKUP(25,A1:A4,B1:B4,\"Missing\",1)");
                sheet.CellFormula(9, 5, "XMATCH(25,A1:A4,2)");
                sheet.CellFormula(10, 5, "MATCH(25,A1:B4,1)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(8, before.SupportedFormulas);
                Assert.Contains("MATCH", before.Capabilities.SupportedFunctions);
                Assert.Contains("XMATCH", before.Capabilities.SupportedFunctions);
                Assert.Contains("XLOOKUP", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "E9" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "E10" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(8, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E5" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E6" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E7" && formula.CachedValue == "Mid");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E8" && formula.CachedValue == "High");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesConditionalAggregatesWithCriteriaCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ConditionalAggregates.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Quarter");
                sheet.CellValue(1, 3, "Amount");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "Q1");
                sheet.CellValue(2, 3, 10d);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "Q1");
                sheet.CellValue(3, 3, 15d);
                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "Q2");
                sheet.CellValue(4, 3, 20d);
                sheet.CellValue(5, 1, "East");
                sheet.CellValue(5, 2, "Q1");
                sheet.CellValue(5, 3, 30d);

                sheet.CellValue(1, 5, "East");
                sheet.CellValue(2, 5, "Q1");
                sheet.CellValue(3, 5, ">15");

                sheet.CellFormula(1, 6, "COUNTIF(A2:A5,E1)");
                sheet.CellFormula(2, 6, "SUMIF(A2:A5,E1,C2:C5)");
                sheet.CellFormula(3, 6, "AVERAGEIF(A2:A5,E1,C2:C5)");
                sheet.CellFormula(4, 6, "COUNTIFS(A2:A5,E1,B2:B5,E2)");
                sheet.CellFormula(5, 6, "SUMIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.CellFormula(6, 6, "AVERAGEIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.CellFormula(7, 6, "SUMIFS(C2:C5,A2:A5,E1,C2:C5,E3)");
                sheet.CellFormula(8, 6, "COUNTIF(A2:A5,\"E*\")");
                sheet.CellFormula(9, 6, "MINIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.CellFormula(10, 6, "MAXIFS(C2:C5,A2:A5,E1,B2:B5,E2)");
                sheet.CellFormula(11, 6, "MINIFS(C2:C5,A2:A5,\"North\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(11, before.TotalFormulas);
                Assert.Equal(10, before.SupportedFormulas);
                Assert.Contains("SUMIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains("AVERAGEIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains("MINIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains("MAXIFS", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "F11" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(10, document.Calculate());
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
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesXLookupFallbackAndSearchMode() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.XLookupFallback.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Owner");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(2, 3, "Alice");
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, 20d);
                sheet.CellValue(3, 3, "Bob");
                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, 30d);
                sheet.CellValue(4, 3, "Ann");

                sheet.CellFormula(1, 5, "XLOOKUP(\"West\",A2:A4,B2:B4,\"Missing\",0)");
                sheet.CellFormula(2, 5, "XLOOKUP(\"North\",A2:A4,C2:C4,\"Missing\",0)");
                sheet.CellFormula(3, 5, "XLOOKUP(\"East\",A2:A4,C2:C4,\"Missing\",0,-1)");
                sheet.CellFormula(4, 5, "XLOOKUP(\"North\",A2:A4,B2:B4,0,0)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesValueReturningIfAndIfError() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IfTextResults.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Owner");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(2, 3, "Alice");
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, 20d);
                sheet.CellValue(3, 3, "Bob");
                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, 30d);
                sheet.CellValue(4, 3, "Ann");

                sheet.CellFormula(1, 5, "IF(A2=\"East\",\"Priority\",\"Standard\")");
                sheet.CellFormula(2, 5, "IF(A3<>\"East\",\"Other\",\"Priority\")");
                sheet.CellFormula(3, 5, "IF(B2>15,\"High\",\"Low\")");
                sheet.CellFormula(4, 5, "IFERROR(XLOOKUP(\"North\",A2:A4,C2:C4),\"Missing\")");
                sheet.CellFormula(5, 5, "IFERROR(XLOOKUP(\"West\",A2:A4,C2:C4),\"Missing\")");
                sheet.CellFormula(6, 5, "IF(A2=A4,\"Same\",\"Different\")");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextFormattingFunction() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextFunction.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Summary");
                sheet.CellValue(1, 1, "East");
                sheet.CellValue(1, 2, 1234.5d);
                sheet.CellValue(1, 3, 0.257d);

                sheet.CellFormula(1, 5, "TEXT(B1,\"#,##0.00\")");
                sheet.CellFormula(2, 5, "TEXT(C1,\"0.0%\")");
                sheet.CellFormula(3, 5, "TEXT(DATE(2026,5,28),\"yyyy-mm-dd\")");
                sheet.CellFormula(4, 5, "TEXT(DATE(2026,5,28),\"mmm yyyy\")");
                sheet.CellFormula(5, 5, "TEXT(TIME(9,5,0),\"hh:mm\")");
                sheet.CellFormula(6, 5, "TEXT(A1,\"@\")");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextCaseReportLabels() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextCaseFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Labels");
                sheet.CellValue(1, 1, "north region");
                sheet.CellValue(2, 1, "Q1-sales review");
                sheet.CellValue(3, 1, "  MIXED   spacing  ");

                sheet.CellFormula(1, 3, "UPPER(A1)");
                sheet.CellFormula(2, 3, "LOWER(\"READY\")");
                sheet.CellFormula(3, 3, "PROPER(A2)");
                sheet.CellFormula(4, 3, "PROPER(TRIM(A3))");
                sheet.CellFormula(5, 3, "CONCAT(UPPER(LEFT(A1,1)),RIGHT(A1,5))");
                sheet.CellFormula(6, 3, "PROPER(A1,A2)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextCleanupReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextCleanupFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Labels");
                sheet.CellValue(1, 1, "North - Q1 - Draft");
                sheet.CellValue(2, 1, "1,234.50");
                sheet.CellValue(3, 1, "west region");

                sheet.CellFormula(1, 3, "SUBSTITUTE(A1,\" - \",\" / \")");
                sheet.CellFormula(2, 3, "SUBSTITUTE(A1,\" - \",\" | \",2)");
                sheet.CellFormula(3, 3, "FIND(\"Q1\",A1)");
                sheet.CellFormula(4, 3, "SEARCH(\"REGION\",A3)");
                sheet.CellFormula(5, 3, "VALUE(A2)");
                sheet.CellFormula(6, 3, "CONCATENATE(LEFT(A3,4),\":\",VALUE(\"42\"))");
                sheet.CellFormula(7, 3, "VALUE(SUBSTITUTE(\"$1,234\",\"$\",\"\"))");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextBeforeAfterReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextBeforeAfterFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Labels");
                sheet.CellValue(1, 1, "North - Q1 - Draft");
                sheet.CellValue(2, 1, "Region=West|Status=Open");
                sheet.CellValue(3, 1, "alpha-BETA-gamma");

                sheet.CellFormula(1, 3, "TEXTBEFORE(A1,\" - \")");
                sheet.CellFormula(2, 3, "TEXTAFTER(A1,\" - \")");
                sheet.CellFormula(3, 3, "TEXTBEFORE(A1,\" - \",2)");
                sheet.CellFormula(4, 3, "TEXTAFTER(A1,\" - \",2)");
                sheet.CellFormula(5, 3, "TEXTBEFORE(A1,\" - \",-1)");
                sheet.CellFormula(6, 3, "TEXTAFTER(A1,\" - \",-1)");
                sheet.CellFormula(7, 3, "TEXTBEFORE(A3,\"beta\",1,1)");
                sheet.CellFormula(8, 3, "TEXTAFTER(A3,\"beta\",1,1)");
                sheet.CellFormula(9, 3, "TEXTBEFORE(A2,\";\",1,0,FALSE,\"Missing\")");
                sheet.CellFormula(10, 3, "TEXTBEFORE(A2,\";\",1,0,TRUE)");
                sheet.CellFormula(11, 3, "TEXTAFTER(A2,\";\",1,0,FALSE,\"Missing\")");
                sheet.CellFormula(12, 3, "TEXTBEFORE(A1,\"\")");
                sheet.CellFormula(13, 3, "TEXTAFTER(A1,\" - \",0)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesTextJoinExactReptAndSumProductReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.TextJoinExactReptSumProduct.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);
                sheet.CellValue(1, 2, 10d);
                sheet.CellValue(2, 2, 20d);
                sheet.CellValue(3, 2, 30d);
                sheet.CellValue(1, 3, "North");
                sheet.CellValue(2, 3, string.Empty);
                sheet.CellValue(3, 3, "West");
                sheet.CellValue(4, 3, "East");

                sheet.CellFormula(1, 5, "TEXTJOIN(\"/\",TRUE,C1:C4)");
                sheet.CellFormula(2, 5, "TEXTJOIN(\"-\",FALSE,C1:C3)");
                sheet.CellFormula(3, 5, "EXACT(C1,\"North\")");
                sheet.CellFormula(4, 5, "EXACT(C1,\"north\")");
                sheet.CellFormula(5, 5, "REPT(\"*\",3)");
                sheet.CellFormula(6, 5, "CONCAT(REPT(\"0\",2),TEXT(A1,\"0\"))");
                sheet.CellFormula(7, 5, "SUMPRODUCT(A1:A3,B1:B3)");
                sheet.CellFormula(8, 5, "SUMPRODUCT(A1:A3)");
                sheet.CellFormula(9, 5, "SUMPRODUCT(A1:A3,B1:B2)");
                sheet.CellFormula(10, 5, "REPT(\"x\",-1)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(8, before.SupportedFormulas);
                Assert.Contains("TEXTJOIN", before.Capabilities.SupportedFunctions);
                Assert.Contains("EXACT", before.Capabilities.SupportedFunctions);
                Assert.Contains("REPT", before.Capabilities.SupportedFunctions);
                Assert.Contains("SUMPRODUCT", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "E9" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "E10" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(8, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E1" && formula.CachedValue == "North/West/East");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E2" && formula.CachedValue == "North--West");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E3" && formula.CachedValue == "1");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E4" && formula.CachedValue == "0");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E5" && formula.CachedValue == "***");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E6" && formula.CachedValue == "001");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E7" && formula.CachedValue == "140");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "E8" && formula.CachedValue == "6");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesDateAndTimeReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.DateTimeReportFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Schedule");
                sheet.CellValue(1, 1, new DateTime(2026, 5, 28, 9, 5, 7));
                sheet.CellValue(1, 2, new DateTime(2026, 5, 27));

                sheet.CellFormula(1, 4, "YEAR(A1)");
                sheet.CellFormula(2, 4, "MONTH(A1)");
                sheet.CellFormula(3, 4, "DAY(A1)");
                sheet.CellFormula(4, 4, "HOUR(A1)");
                sheet.CellFormula(5, 4, "MINUTE(A1)");
                sheet.CellFormula(6, 4, "SECOND(A1)");
                sheet.CellFormula(7, 4, "TEXT(EDATE(A1,1),\"yyyy-mm-dd\")");
                sheet.CellFormula(8, 4, "TEXT(EOMONTH(A1,0),\"yyyy-mm-dd\")");
                sheet.CellFormula(9, 4, "DAYS(DATE(2026,6,11),DATE(2026,5,28))");
                sheet.CellFormula(10, 4, "WEEKDAY(DATE(2026,5,31),2)");
                sheet.CellFormula(11, 4, "NETWORKDAYS(DATE(2026,5,25),DATE(2026,5,29),B1:B1)");
                sheet.CellFormula(12, 4, "TEXT(TIME(23,75,0),\"hh:mm\")");
                sheet.CellFormula(13, 4, "TEXT(DATEVALUE(\"2026-05-28\"),\"yyyy-mm-dd\")");
                sheet.CellFormula(14, 4, "YEAR(DATEVALUE(\"May 28 2026\"))");
                sheet.CellFormula(15, 4, "TEXT(DATEVALUE(\"5/28/2026\"),\"yyyy-mm-dd\")");
                sheet.CellFormula(16, 4, "TEXT(TIMEVALUE(\"09:05:07\"),\"hh:mm\")");
                sheet.CellFormula(17, 4, "SECOND(TIMEVALUE(\"09:05:07\"))");
                sheet.CellFormula(18, 4, "TEXT(DATEVALUE(TEXT(A1,\"yyyy-mm-dd\")),\"yyyy-mm-dd\")");
                sheet.CellFormula(19, 4, "DATEVALUE(\"not a date\")");
                sheet.CellFormula(20, 4, "DATEVALUE(\"2026-05-28\")");
                sheet.CellFormula(21, 4, "TIMEVALUE(\"09:05:07\")");
                sheet.CellFormula(22, 4, "WEEKNUM(A1)");
                sheet.CellFormula(23, 4, "WEEKNUM(A1,2)");
                sheet.CellFormula(24, 4, "ISOWEEKNUM(DATE(2021,1,1))");
                sheet.CellFormula(25, 4, "WEEKNUM(DATE(2021,1,1),21)");
                sheet.CellFormula(26, 4, "DAYS360(DATE(2026,2,28),DATE(2026,3,31))");
                sheet.CellFormula(27, 4, "DAYS360(DATE(2026,2,28),DATE(2026,3,31),TRUE)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesDateDifReportUnits() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.DateDifFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Schedule");
                sheet.CellValue(1, 1, new DateTime(2020, 1, 15));
                sheet.CellValue(1, 2, new DateTime(2026, 5, 28));

                sheet.CellFormula(1, 4, "DATEDIF(A1,B1,\"Y\")");
                sheet.CellFormula(2, 4, "DATEDIF(A1,B1,\"M\")");
                sheet.CellFormula(3, 4, "DATEDIF(A1,B1,\"D\")");
                sheet.CellFormula(4, 4, "DATEDIF(A1,B1,\"YM\")");
                sheet.CellFormula(5, 4, "DATEDIF(A1,B1,\"YD\")");
                sheet.CellFormula(6, 4, "DATEDIF(A1,B1,\"MD\")");
                sheet.CellFormula(7, 4, "DATEDIF(B1,A1,\"D\")");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(7, before.TotalFormulas);
                Assert.Equal(6, before.SupportedFormulas);
                Assert.Contains("DATEDIF", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D7" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(6, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D1" && formula.CachedValue == "6");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D2" && formula.CachedValue == "76");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D3" && formula.CachedValue == "2325");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D4" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D5" && formula.CachedValue == "133");
                Assert.Contains(after.Formulas, formula => formula.CellReference == "D6" && formula.CachedValue == "13");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesYearFracReportBases() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.YearFracFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Schedule");
                sheet.CellValue(1, 1, new DateTime(2024, 1, 1));
                sheet.CellValue(1, 2, new DateTime(2025, 1, 1));
                sheet.CellValue(2, 1, new DateTime(2024, 1, 31));
                sheet.CellValue(2, 2, new DateTime(2024, 2, 29));

                sheet.CellFormula(1, 4, "YEARFRAC(A1,B1)");
                sheet.CellFormula(2, 4, "YEARFRAC(A1,B1,1)");
                sheet.CellFormula(3, 4, "YEARFRAC(A1,B1,2)");
                sheet.CellFormula(4, 4, "YEARFRAC(A1,B1,3)");
                sheet.CellFormula(5, 4, "YEARFRAC(A2,B2,0)");
                sheet.CellFormula(6, 4, "YEARFRAC(A2,B2,4)");
                sheet.CellFormula(7, 4, "YEARFRAC(B1,A1,1)");
                sheet.CellFormula(8, 4, "YEARFRAC(A1,B1,5)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(8, before.TotalFormulas);
                Assert.Equal(6, before.SupportedFormulas);
                Assert.Contains("YEARFRAC", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D7" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D8" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(6, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                AssertCachedNumber(after, "D1", 1d);
                AssertCachedNumber(after, "D2", 1d);
                AssertCachedNumber(after, "D3", 366d / 360d);
                AssertCachedNumber(after, "D4", 366d / 365d);
                AssertCachedNumber(after, "D5", 29d / 360d);
                AssertCachedNumber(after, "D6", 29d / 360d);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesFinancialReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.FinancialFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Finance");
                sheet.CellValue(1, 1, 0.05d / 12d);
                sheet.CellValue(2, 1, 360d);
                sheet.CellValue(3, 1, 250000d);
                sheet.CellValue(6, 1, 100d);
                sheet.CellValue(7, 1, 200d);
                sheet.CellValue(8, 1, 300d);

                sheet.CellFormula(1, 4, "PMT(A1,A2,A3)");
                sheet.CellFormula(2, 4, "PV(A1,A2,PMT(A1,A2,A3))");
                sheet.CellFormula(3, 4, "FV(A1,60,-200,0)");
                sheet.CellFormula(4, 4, "NPER(A1,PMT(A1,A2,A3),A3)");
                sheet.CellFormula(5, 4, "NPV(0.1,A6:A8)");
                sheet.CellFormula(6, 4, "PMT(0,12,1200)");
                sheet.CellFormula(7, 4, "PV(0,12,-100)");
                sheet.CellFormula(8, 4, "FV(0,12,-100)");
                sheet.CellFormula(9, 4, "PMT(A1,A2,A3,0,1)");
                sheet.CellFormula(10, 4, "PMT(A1,A2,A3,0,2)");
                sheet.CellFormula(11, 4, "NPER(A1,0,A3)");

                ExcelFormulaInspection before = sheet.InspectFormulas();
                Assert.Equal(11, before.TotalFormulas);
                Assert.Equal(9, before.SupportedFormulas);
                Assert.Contains("PMT", before.Capabilities.SupportedFunctions);
                Assert.Contains("PV", before.Capabilities.SupportedFunctions);
                Assert.Contains("FV", before.Capabilities.SupportedFunctions);
                Assert.Contains("NPER", before.Capabilities.SupportedFunctions);
                Assert.Contains("NPV", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D10" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.CellReference == "D11" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(9, document.Calculate());
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
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_RejectsOversizedCovarianceRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.OversizedCovarianceRange.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Stats");
                document.AddWorkSheet("Data");
                sheet.CellFormula(1, 1, "COVARIANCE.P(Data!A1:XFD1048576,Data!A1:XFD1048576)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();
                Assert.Equal(1, inspection.TotalFormulas);
                Assert.Equal(0, inspection.SupportedFormulas);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && !formula.IsSupportedByOfficeIMO);
                Assert.Equal(0, document.Calculate());
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_RejectsCumulativeOversizedFormulaRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.CumulativeOversizedRanges.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Stats");
                document.AddWorkSheet("Data");
                sheet.CellFormula(1, 1, "COVARIANCE.P(Data!A1:A100000,Data!B1:B100000)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();
                Assert.Equal(1, inspection.TotalFormulas);
                Assert.Equal(0, inspection.SupportedFormulas);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && !formula.IsSupportedByOfficeIMO);
                Assert.Equal(0, document.Calculate());
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_RejectsOversizedDirectLookupRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.OversizedLookupRange.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Lookup");
                document.AddWorkSheet("Data");
                sheet.CellFormula(1, 1, "VLOOKUP(\"missing\",Data!A1:A100001,1,FALSE)");

                ExcelFormulaInspection inspection = sheet.InspectFormulas();
                Assert.Equal(1, inspection.TotalFormulas);
                Assert.Equal(0, inspection.SupportedFormulas);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && !formula.IsSupportedByOfficeIMO);
                Assert.Equal(0, document.Calculate());
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesStatisticalReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.StatisticalFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Stats");
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(2, 1, 12d);
                sheet.CellValue(3, 1, 14d);
                sheet.CellValue(4, 1, 18d);
                sheet.CellValue(5, 1, 21d);
                sheet.CellValue(1, 2, 10d);
                sheet.CellValue(2, 2, 12d);
                sheet.CellValue(3, 2, 14d);
                sheet.CellValue(4, 2, 14d);
                sheet.CellValue(5, 2, 21d);
                sheet.CellValue(1, 3, 1d);
                sheet.CellValue(2, 3, 2d);
                sheet.CellValue(3, 3, 3d);
                sheet.CellValue(4, 3, 4d);
                sheet.CellValue(5, 3, 5d);
                sheet.CellValue(6, 1, "Text");

                sheet.CellFormula(1, 4, "VAR.P(A1:A5)");
                sheet.CellFormula(2, 4, "VAR.S(A1:A5)");
                sheet.CellFormula(3, 4, "STDEV.P(A1:A5)");
                sheet.CellFormula(4, 4, "STDEV.S(A1:A5)");
                sheet.CellFormula(5, 4, "PERCENTILE.INC(A1:A5,0)");
                sheet.CellFormula(6, 4, "PERCENTILE.INC(A1:A5,0.25)");
                sheet.CellFormula(7, 4, "PERCENTILE.INC(A1:A5,0.9)");
                sheet.CellFormula(8, 4, "QUARTILE.INC(A1:A5,2)");
                sheet.CellFormula(9, 4, "QUARTILE.INC(A1:A5,4)");
                sheet.CellFormula(10, 4, "PERCENTILE.INC(A1:A5,1.5)");
                sheet.CellFormula(11, 4, "QUARTILE.INC(A1:A5,5)");
                sheet.CellFormula(12, 4, "STDEV.S(A1)");
                sheet.CellFormula(13, 4, "SUMSQ(A1:A5)");
                sheet.CellFormula(14, 4, "RANK.EQ(18,A1:A5,0)");
                sheet.CellFormula(15, 4, "RANK.EQ(A3,A1:A5,1)");
                sheet.CellFormula(16, 4, "RANK.EQ(15,A1:A5)");
                sheet.CellFormula(17, 4, "RANK.AVG(14,B1:B5,0)");
                sheet.CellFormula(18, 4, "RANK.AVG(14,B1:B5,1)");
                sheet.CellFormula(19, 4, "PERCENTRANK.INC(A1:A5,18)");
                sheet.CellFormula(20, 4, "PERCENTRANK.INC(A1:A5,16)");
                sheet.CellFormula(21, 4, "PERCENTRANK.INC(A1:A5,5)");
                sheet.CellFormula(22, 4, "PERCENTILE.EXC(A1:A5,0.25)");
                sheet.CellFormula(23, 4, "QUARTILE.EXC(A1:A5,3)");
                sheet.CellFormula(24, 4, "PERCENTRANK.EXC(A1:A5,18)");
                sheet.CellFormula(25, 4, "PERCENTRANK.EXC(A1:A5,16)");
                sheet.CellFormula(26, 4, "PERCENTILE.EXC(A1:A5,0.99)");
                sheet.CellFormula(27, 4, "CORREL(A1:A5,C1:C5)");
                sheet.CellFormula(28, 4, "SLOPE(A1:A5,C1:C5)");
                sheet.CellFormula(29, 4, "INTERCEPT(A1:A5,C1:C5)");
                sheet.CellFormula(30, 4, "RSQ(A1:A5,C1:C5)");
                sheet.CellFormula(31, 4, "FORECAST.LINEAR(6,A1:A5,C1:C5)");
                sheet.CellFormula(32, 4, "SLOPE(A1:A5,C1:C4)");
                sheet.CellFormula(33, 4, "MEDIAN(A1:A5)");
                sheet.CellFormula(34, 4, "LARGE(A1:A5,2)");
                sheet.CellFormula(35, 4, "SMALL(A1:A5,3)");
                sheet.CellFormula(36, 4, "MODE.SNGL(B1:B5)");
                sheet.CellFormula(37, 4, "MODE(A1:A5)");
                sheet.CellFormula(38, 4, "GEOMEAN(A1:A5)");
                sheet.CellFormula(39, 4, "HARMEAN(A1:A5)");
                sheet.CellFormula(40, 4, "AVERAGEA(A1:A6)");
                sheet.CellFormula(41, 4, "MINA(A1:A6)");
                sheet.CellFormula(42, 4, "MAXA(A1:A6)");
                sheet.CellFormula(43, 4, "COUNTA(A1:A6)");
                sheet.CellFormula(44, 4, "AVERAGEA(\"skip\",TRUE,FALSE,5)");
                sheet.CellFormula(45, 4, "AVEDEV(A1:A5)");
                sheet.CellFormula(46, 4, "DEVSQ(A1:A5)");
                sheet.CellFormula(47, 4, "SUMXMY2(A1:A5,C1:C5)");
                sheet.CellFormula(48, 4, "SUMX2MY2(A1:A5,C1:C5)");
                sheet.CellFormula(49, 4, "SUMX2PY2(A1:A5,C1:C5)");
                sheet.CellFormula(50, 4, "SUMXMY2(A1:A5,C1:C4)");
                sheet.CellFormula(51, 4, "COVARIANCE.P(A1:A5,C1:C5)");
                sheet.CellFormula(52, 4, "COVARIANCE.S(A1:A5,C1:C5)");
                sheet.CellFormula(53, 4, "COVAR(A1:A5,C1:C5)");
                sheet.CellFormula(54, 4, "COVARIANCE.P(A1:A5,C1:C4)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesRoundingReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.RoundingFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Rounding");
                sheet.CellFormula(1, 1, "MROUND(23,5)");
                sheet.CellFormula(2, 1, "MROUND(-23,-5)");
                sheet.CellFormula(3, 1, "CEILING.MATH(23,5)");
                sheet.CellFormula(4, 1, "FLOOR.MATH(23,5)");
                sheet.CellFormula(5, 1, "CEILING.MATH(-23,5)");
                sheet.CellFormula(6, 1, "CEILING.MATH(-23,5,1)");
                sheet.CellFormula(7, 1, "FLOOR.MATH(-23,5)");
                sheet.CellFormula(8, 1, "FLOOR.MATH(-23,5,1)");
                sheet.CellFormula(9, 1, "MROUND(23,-5)");
                sheet.CellFormula(10, 1, "ROUND(23.456,2)");
                sheet.CellFormula(11, 1, "ROUND(1234,-2)");
                sheet.CellFormula(12, 1, "ROUNDUP(23.451,2)");
                sheet.CellFormula(13, 1, "ROUNDUP(1234,-2)");
                sheet.CellFormula(14, 1, "ROUNDDOWN(23.459,2)");
                sheet.CellFormula(15, 1, "ROUNDDOWN(1234,-2)");
                sheet.CellFormula(16, 1, "TRUNC(23.459,2)");
                sheet.CellFormula(17, 1, "TRUNC(1234,-2)");
                sheet.CellFormula(18, 1, "INT(-23.1)");
                sheet.CellFormula(19, 1, "CEILING(23,5)");
                sheet.CellFormula(20, 1, "FLOOR(23,5)");
                sheet.CellFormula(21, 1, "ROUND(23,16)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesChooseReportSelectors() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ChooseSelectors.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Selectors");
                sheet.CellValue(1, 1, 2d);
                sheet.CellValue(2, 1, "EU");
                sheet.CellValue(3, 1, "US");
                sheet.CellValue(4, 1, "APAC");

                sheet.CellFormula(1, 3, "CHOOSE(2,10,20,30)");
                sheet.CellFormula(2, 3, "CHOOSE(A1,\"Draft\",\"Final\",\"Archived\")");
                sheet.CellFormula(3, 3, "CHOOSE(MATCH(\"US\",A2:A4,0),100,200,300)");
                sheet.CellFormula(4, 3, "TEXT(CHOOSE(3,0.1,0.2,0.3),\"0%\")");
                sheet.CellFormula(5, 3, "CHOOSE(4,10,20,30)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesInfoReportGuards() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.InfoGuards.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Info");
                sheet.CellValue(1, 1, 42d);
                sheet.CellValue(3, 1, "Ready");

                sheet.CellFormula(1, 3, "ISBLANK(A2)");
                sheet.CellFormula(2, 3, "ISNUMBER(A1)");
                sheet.CellFormula(3, 3, "ISTEXT(A3)");
                sheet.CellFormula(4, 3, "IF(ISBLANK(A2),\"Missing\",\"Present\")");
                sheet.CellFormula(5, 3, "IF(ISNUMBER(A1),A1*2,0)");
                sheet.CellFormula(6, 3, "ISBLANK(\"\")");
                sheet.CellFormula(7, 3, "ISNUMBER(A3)");
                sheet.CellFormula(8, 3, "ISTEXT(A2)");
                sheet.CellFormula(9, 3, "ISBLANK(A2,A3)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesErrorInfoReportGuards() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ErrorInfoGuards.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Info");
                sheet.CellValue(1, 1, "East");
                sheet.CellValue(2, 1, "West");
                sheet.CellValue(1, 2, 10d);
                sheet.CellValue(2, 2, 0d);

                sheet.CellFormula(1, 4, "ISERROR(XLOOKUP(\"North\",A1:A2,B1:B2))");
                sheet.CellFormula(2, 4, "ISNA(XLOOKUP(\"North\",A1:A2,B1:B2))");
                sheet.CellFormula(3, 4, "ISERR(FIND(\"z\",\"abc\"))");
                sheet.CellFormula(4, 4, "ISERROR(FIND(\"a\",\"abc\"))");
                sheet.CellFormula(5, 4, "IF(ISERROR(XLOOKUP(\"North\",A1:A2,B1:B2)),\"Missing\",\"Found\")");
                sheet.CellFormula(6, 4, "IF(ISNA(#N/A),\"NA\",\"Other\")");
                sheet.CellFormula(7, 4, "IF(ISERR(#DIV/0!),\"Err\",\"Other\")");
                sheet.CellFormula(8, 4, "ISERROR(B1/B2)");
                sheet.CellFormula(9, 4, "ISNA(FIND(\"z\",\"abc\"))");
                sheet.CellFormula(10, 4, "ISERR(#N/A)");
                sheet.CellFormula(11, 4, "ISERROR(\"text\",A1)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesIfNaReportFallbacks() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.IfNaFallbacks.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Info");
                sheet.CellValue(1, 1, "East");
                sheet.CellValue(2, 1, "West");
                sheet.CellValue(1, 2, 10d);
                sheet.CellValue(2, 2, 20d);

                sheet.CellFormula(1, 4, "IFNA(XLOOKUP(\"North\",A1:A2,B1:B2),\"Missing\")");
                sheet.CellFormula(2, 4, "IFNA(XLOOKUP(\"East\",A1:A2,B1:B2),\"Missing\")");
                sheet.CellFormula(3, 4, "IFNA(#N/A,\"NA\")");
                sheet.CellFormula(4, 4, "IFNA(#DIV/0!,\"NA\")");
                sheet.CellFormula(5, 4, "IFERROR(#DIV/0!,\"Any\")");
                sheet.CellFormula(6, 4, "IFERROR(#N/A,\"Any\")");
                sheet.CellFormula(7, 4, "IFNA(FIND(\"z\",\"abc\"),\"NA\")");
                sheet.CellFormula(8, 4, "IFNA(\"ready\",\"NA\")");
                sheet.CellFormula(9, 4, "IFNA(XLOOKUP(\"North\",A1:A2,B1:B2),A1)");
                sheet.CellFormula(10, 4, "IFNA(\"text\",A1,B1)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesFormulaInspectionReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.FormulaInspectionFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet source = document.AddWorkSheet("Source Data");
                source.CellValue(1, 1, 2d);
                source.CellValue(1, 2, 3d);
                source.CellFormula(1, 3, "A1+B1");

                ExcelSheet sheet = document.AddWorkSheet("Info");
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(1, 2, 5d);
                sheet.CellFormula(1, 3, "A1+B1");
                sheet.CellFormula(2, 3, "IF(A1>3,\"High\",\"Low\")");

                sheet.CellFormula(1, 4, "ISFORMULA(C1)");
                sheet.CellFormula(2, 4, "ISFORMULA(A1)");
                sheet.CellFormula(3, 4, "FORMULATEXT(C1)");
                sheet.CellFormula(4, 4, "FORMULATEXT(C2)");
                sheet.CellFormula(5, 4, "IF(ISFORMULA(C2),FORMULATEXT(C2),\"Missing\")");
                sheet.CellFormula(6, 4, "LEN(FORMULATEXT(C1))");
                sheet.CellFormula(7, 4, "ISFORMULA('Source Data'!C1)");
                sheet.CellFormula(8, 4, "FORMULATEXT('Source Data'!C1)");
                sheet.CellFormula(9, 4, "FORMULATEXT(A1)");
                sheet.CellFormula(10, 4, "ISFORMULA(A1:B1)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesReferenceShapeReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.ReferenceShapeFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Source Data");

                ExcelSheet sheet = document.AddWorkSheet("Reference");
                sheet.CellFormula(1, 4, "ROW(A5)");
                sheet.CellFormula(2, 4, "COLUMN(C1)");
                sheet.CellFormula(3, 4, "ROWS(A1:B3)");
                sheet.CellFormula(4, 4, "COLUMNS(A1:B3)");
                sheet.CellFormula(5, 4, "ROW('Source Data'!D5)");
                sheet.CellFormula(6, 4, "COLUMN('Source Data'!D5)");
                sheet.CellFormula(7, 4, "ROWS('Source Data'!B2:D5)");
                sheet.CellFormula(8, 4, "COLUMNS('Source Data'!B2:D5)");
                sheet.CellFormula(9, 4, "ROW()");
                sheet.CellFormula(10, 4, "ROWS(A1,B1)");

                ExcelFormulaInspection before = document.InspectFormulas();
                Assert.Equal(10, before.TotalFormulas);
                Assert.Equal(8, before.SupportedFormulas);
                Assert.Contains("ROW", before.Capabilities.SupportedFunctions);
                Assert.Contains("COLUMN", before.Capabilities.SupportedFunctions);
                Assert.Contains("ROWS", before.Capabilities.SupportedFunctions);
                Assert.Contains("COLUMNS", before.Capabilities.SupportedFunctions);
                Assert.Contains(before.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D9" && !formula.IsSupportedByOfficeIMO);
                Assert.Contains(before.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D10" && !formula.IsSupportedByOfficeIMO);

                Assert.Equal(8, document.Calculate());
                ExcelFormulaInspection after = document.InspectFormulas();
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D1" && formula.CachedValue == "5");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D2" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D3" && formula.CachedValue == "3");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D4" && formula.CachedValue == "2");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D5" && formula.CachedValue == "5");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D6" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D7" && formula.CachedValue == "4");
                Assert.Contains(after.Formulas, formula => formula.SheetName == "Reference" && formula.CellReference == "D8" && formula.CachedValue == "3");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesWorkdaySchedulingFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.WorkdayFunctions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Schedule");
                sheet.CellValue(1, 1, new DateTime(2026, 5, 28));
                sheet.CellValue(1, 2, new DateTime(2026, 5, 29));
                sheet.CellValue(2, 2, new DateTime(2026, 6, 1));

                sheet.CellFormula(1, 4, "TEXT(WORKDAY(A1,1),\"yyyy-mm-dd\")");
                sheet.CellFormula(2, 4, "TEXT(WORKDAY(A1,1,B1:B1),\"yyyy-mm-dd\")");
                sheet.CellFormula(3, 4, "TEXT(WORKDAY(A1,1,B1:B2),\"yyyy-mm-dd\")");
                sheet.CellFormula(4, 4, "TEXT(WORKDAY(A1,-3),\"yyyy-mm-dd\")");
                sheet.CellFormula(5, 4, "TEXT(WORKDAY.INTL(A1,2,\"0000110\"),\"yyyy-mm-dd\")");
                sheet.CellFormula(6, 4, "TEXT(WORKDAY.INTL(A1,1,11),\"yyyy-mm-dd\")");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_CalculatesSubtotalAndCountBlankReportFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.SubtotalCountBlank.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Audit");
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(2, 1, 20d);
                sheet.CellValue(4, 1, "North");
                sheet.CellValue(5, 1, string.Empty);
                sheet.CellFormula(6, 1, "A1+A2");

                sheet.CellFormula(1, 3, "COUNTBLANK(A1:A6)");
                sheet.CellFormula(2, 3, "SUBTOTAL(9,A1:A6)");
                sheet.CellFormula(3, 3, "SUBTOTAL(1,A1:A6)");
                sheet.CellFormula(4, 3, "SUBTOTAL(2,A1:A6)");
                sheet.CellFormula(5, 3, "SUBTOTAL(3,A1:A6)");
                sheet.CellFormula(6, 3, "SUBTOTAL(4,A1:A6)");
                sheet.CellFormula(7, 3, "SUBTOTAL(5,A1:A6)");
                sheet.CellFormula(8, 3, "SUBTOTAL(109,A1:A2,A6:A6)");
                sheet.CellFormula(9, 3, "SUBTOTAL(6,A1:A6)");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
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
                ExcelSheet sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Owner");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(2, 3, "Alice");
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, 20d);
                sheet.CellValue(3, 3, "Bob");
                sheet.CellValue(4, 1, "North");
                sheet.CellValue(4, 2, 5d);
                sheet.CellValue(4, 3, "Nina");

                sheet.CellFormula(1, 5, "IFS(B2>=20,\"High\",B2>=10,\"Medium\",TRUE,\"Low\")");
                sheet.CellFormula(2, 5, "IFS(B4>=20,\"High\",B4>=10,\"Medium\",TRUE,\"Low\")");
                sheet.CellFormula(3, 5, "SWITCH(A3,\"East\",\"Priority\",\"West\",\"Standard\",\"Other\")");
                sheet.CellFormula(4, 5, "SWITCH(XLOOKUP(\"East\",A2:A4,C2:C4),\"Alice\",\"Owner A\",\"Ann\",\"Owner Ann\",\"Other\")");
                sheet.CellFormula(5, 5, "SWITCH(A4,\"East\",1,\"West\",2,0)");
                sheet.CellFormula(6, 5, "IFS(A2=\"East\",TEXT(B2,\"0.0\"),TRUE,\"Missing\")");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaInspection_ReportsDependenciesAndIssues() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaEvaluator.DependencyDiagnostics.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, 10d);
                sheet.CellFormula(1, 2, "A1+5");
                sheet.CellFormula(1, 3, "B1+1");
                sheet.CellFormula(1, 4, "D1+1");
                sheet.CellFormula(1, 5, "UNSUPPORTED(B1)");

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
                ExcelSheet sales = document.AddWorkSheet("Sales");
                sales.CellValue(1, 1, 10d);
                sales.CellFormula(1, 2, "A1+5");
                sales.CellFormula(1, 3, "B1+1");

                ExcelSheet summary = document.AddWorkSheet("Summary");
                summary.CellFormula(1, 1, "Sales!C1*2");

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
