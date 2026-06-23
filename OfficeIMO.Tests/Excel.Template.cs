using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelTemplate_ReplacesTextMarkersAcrossWorkbook() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.Markers.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var invoice = document.AddWorkSheet("Invoice");
                invoice.CellAt(1, 1).SetValue("Invoice {{Invoice.Number}}").HeaderStyle();
                invoice.CellAt(2, 1).SetValue("Customer: {{Customer.Name}}");

                var summary = document.AddWorkSheet("Summary");
                summary.CellAt(1, 1).SetValue("Total: {{Total}}");

                int replacements = document.ApplyTemplate(new Dictionary<string, object?> {
                    ["Invoice.Number"] = "INV-001",
                    ["Customer.Name"] = "Adatum",
                    ["Total"] = 123.45
                }, System.Globalization.CultureInfo.InvariantCulture);

                Assert.Equal(3, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Equal("Invoice INV-001", document["Invoice"].CellAt(1, 1).GetValue<string>());
                Assert.Equal("Customer: Adatum", document["Invoice"].CellAt(2, 1).GetValue<string>());
                Assert.Equal("Total: 123.45", document["Summary"].CellAt(1, 1).GetValue<string>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_CreateFromTemplatePreservesStylesAndThemeParts() {
            string templatePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.SourceTemplate.xlsx");
            string outputPath = Path.Combine(_directoryWithFiles, "ExcelTemplate.FromTemplate.xlsx");

            using (var document = ExcelDocument.Create(templatePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("{{Name}}").SetBold().SetFillColor("FFF2CC");
                sheet.CellAt(2, 1).SetValue("Footer");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) });
                sheet.AddChart(data, row: 1, column: 4, widthPixels: 320, heightPixels: 200, type: ExcelChartType.ColumnClustered, title: "Template Chart");
                document.Save(false);
            }

            using (var document = ExcelDocument.CreateFromTemplate(templatePath, outputPath)) {
                int replacements = document.ApplyTemplate(new Dictionary<string, object?> {
                    ["Name"] = "Adatum"
                });

                Assert.Equal(1, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(outputPath, readOnly: true)) {
                var sheet = document["Template"];
                Assert.Equal("Adatum", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Footer", sheet.CellAt(2, 1).GetValue<string>());

                var style = sheet.CellAt(1, 1).GetStyle();
                Assert.True(style.Bold);
                Assert.Equal("FFFFF2CC", style.FillColorArgb);

                Assert.NotNull(document.WorkbookPartRoot.GetPartsOfType<ThemePart>().FirstOrDefault());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_CanThrowOnMissingMarker() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.MissingMarker.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("Hello {{Name}}");

                Assert.Throws<InvalidOperationException>(() =>
                    sheet.ApplyTemplate(new Dictionary<string, object?>(), throwOnMissing: true));

                Assert.Throws<InvalidOperationException>(() =>
                    sheet.ApplyTemplate(new Dictionary<string, object?>(), new ExcelTemplateOptions {
                        MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
                    }));
            }
        }

        [Fact]
        public void Test_ExcelTemplate_MissingValuePolicyCanReplaceMarkersWithEmptyStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.MissingValuePolicy.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("Hello {{Name}} {{OptionalSuffix}}");
                sheet.CellAt(2, 1).SetValue("{{OptionalTotal}}");
                sheet.CellAt(3, 1).SetValue("Still {{Unknown}}");

                int replacements = sheet.ApplyTemplate(new Dictionary<string, object?> {
                    ["Name"] = "Adatum"
                }, new ExcelTemplateOptions {
                    MissingValueBehavior = ExcelTemplateMissingValueBehavior.EmptyString
                });

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Template"];
                Assert.Equal("Hello Adatum ", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal(string.Empty, sheet.CellAt(2, 1).GetValue<string>());
                Assert.Equal("Still ", sheet.CellAt(3, 1).GetValue<string>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatsTemplateRowAndShiftsFollowingRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRows.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Item").HeaderStyle();
                sheet.CellAt(1, 2).SetValue("Amount").HeaderStyle();
                sheet.CellAt(2, 1).SetValue("{{Name}}");
                sheet.CellAt(2, 2).SetValue("{{Amount:currency}}");
                sheet.CellAt(3, 1).SetValue("Footer");

                int replacements = sheet.ApplyTemplateRows(2, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting",
                        ["Amount"] = 1200m
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support",
                        ["Amount"] = 300m
                    }
                }, new ExcelTemplateOptions {
                    FormatProvider = System.Globalization.CultureInfo.GetCultureInfo("en-US"),
                    MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
                });

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Invoice"];
                Assert.Equal("Consulting", sheet.CellAt(2, 1).GetValue<string>());
                Assert.Equal(1200d, sheet.CellAt(2, 2).GetValue<double>());
                Assert.Equal("Support", sheet.CellAt(3, 1).GetValue<string>());
                Assert.Equal(300d, sheet.CellAt(3, 2).GetValue<double>());
                Assert.Equal("Footer", sheet.CellAt(4, 1).GetValue<string>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRejectsExpansionBeyondWorksheetLimit() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsBounds.xlsx");

            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorkSheet("Invoice");
            sheet.CellAt(A1.MaxRows, 1).SetValue("{{Name}}");

            ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                sheet.ApplyTemplateRows(A1.MaxRows, new[] {
                    new Dictionary<string, object?> { ["Name"] = "One" },
                    new Dictionary<string, object?> { ["Name"] = "Two" }
                }));
            Assert.Equal("rowBindings", exception.ParamName);
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRebasesFormulasAndShiftedMerges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsFormulas.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Item");
                sheet.CellAt(1, 2).SetValue("Amount");
                sheet.CellAt(1, 3).SetValue("Double");
                sheet.CellAt(2, 1).SetValue("{{Name}}");
                sheet.CellAt(2, 2).SetValue("{{Amount}}");
                sheet.CellFormula(2, 3, "B2*2");
                sheet.CellAt(3, 1).SetValue("Footer");
                sheet.CellFormula(3, 3, "B3");
                sheet.MergeRange("A3:B3");

                int replacements = sheet.ApplyTemplateRows(2, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting",
                        ["Amount"] = 1200
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support",
                        ["Amount"] = 300
                    }
                });

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                var merge = Assert.Single(wsPart.Worksheet.Descendants<MergeCell>());

                Assert.Equal("B2*2", cells["C2"].CellFormula!.Text);
                Assert.Equal("B3*2", cells["C3"].CellFormula!.Text);
                Assert.Equal("B4", cells["C4"].CellFormula!.Text);
                Assert.Equal("A4:B4", merge.Reference!.Value);
                Assert.DoesNotContain("{{", wsPart.Worksheet.OuterXml, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRebasesStationaryFormulasAndRowMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsStationaryFormulas.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellFormula(1, 4, "B4+$B$4");
                sheet.CellFormula(1, 5, "'Invoice'!$B$4");
                sheet.CellAt(2, 1).SetValue("Intro");
                sheet.CellAt(3, 1).SetValue("{{Name}}");
                sheet.CellAt(3, 2).SetValue("{{Amount}}");
                sheet.CellFormula(3, 3, "B3*2");
                sheet.CellFormula(3, 4, "B2+B3+\"A3\"");
                sheet.CellFormula(3, 6, "$B$3+B3");
                sheet.CellAt(4, 1).SetValue("Footer");
                sheet.CellAt(4, 2).SetValue(25);
                sheet.CellFormula(4, 3, "B4");
                sheet.MergeRange("A2:A4");
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Row templateRow = worksheetPart.Worksheet.Descendants<Row>().Single(row => row.RowIndex?.Value == 3U);
                templateRow.Height = 30D;
                templateRow.CustomHeight = true;
                templateRow.Hidden = true;
                templateRow.OutlineLevel = 1;
                worksheetPart.Worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document["Invoice"];
                int replacements = sheet.ApplyTemplateRows(3, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting",
                        ["Amount"] = 1200
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support",
                        ["Amount"] = 300
                    }
                });

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                var rows = worksheetPart.Worksheet.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
                MergeCell merge = Assert.Single(worksheetPart.Worksheet.Descendants<MergeCell>());

                Assert.Equal("B5+$B$5", cells["D1"].CellFormula!.Text);
                Assert.Equal("'Invoice'!$B$5", cells["E1"].CellFormula!.Text);
                Assert.Equal("B3*2", cells["C3"].CellFormula!.Text);
                Assert.Equal("B4*2", cells["C4"].CellFormula!.Text);
                Assert.Equal("B3+B4+\"A3\"", cells["D4"].CellFormula!.Text);
                Assert.Equal("$B$3+B3", cells["F3"].CellFormula!.Text);
                Assert.Equal("$B$3+B4", cells["F4"].CellFormula!.Text);
                Assert.Equal("B5", cells["C5"].CellFormula!.Text);
                Assert.Equal("A2:A5", merge.Reference!.Value);
                Assert.True(rows[4U].Hidden!.Value);
                Assert.True(rows[4U].CustomHeight!.Value);
                Assert.Equal(30D, rows[4U].Height!.Value);
                Assert.Equal(1, rows[4U].OutlineLevel!.Value);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRebasesSparklinesBelowTemplate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsSparklines.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Item");
                sheet.CellAt(2, 1).SetValue("{{Name}}");
                sheet.CellAt(3, 1).SetValue(1);
                sheet.CellAt(3, 2).SetValue(2);
                sheet.CellAt(3, 3).SetValue(3);
                sheet.AddSparklines("A3:C3", "D3");

                int replacements = sheet.ApplyTemplateRows(2, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting"
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support"
                    }
                });

                Assert.Equal(2, replacements);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                var sparkline = worksheet.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>().Single();

                Assert.Equal("A4:C4", sparkline.Formula!.Text);
                Assert.Equal("D4", sparkline.ReferenceSequence!.Text);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRebasesChartsAndKeepsPageSetup() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsCharts.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Item");
                sheet.CellAt(2, 1).SetValue("{{Name}}");
                sheet.CellAt(4, 1).SetValue("Region");
                sheet.CellAt(4, 2).SetValue("Amount");
                sheet.CellAt(5, 1).SetValue("EU");
                sheet.CellAt(5, 2).SetValue(10);
                sheet.CellAt(6, 1).SetValue("US");
                sheet.CellAt(6, 2).SetValue(20);
                sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 0U);
                sheet.AddChartFromRange("A4:B6", row: 8, column: 4, widthPixels: 320, heightPixels: 180, type: ExcelChartType.ColumnClustered, hasHeaders: true, title: "Sales");

                int replacements = sheet.ApplyTemplateRows(2, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting"
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support"
                    }
                });

                Assert.Equal(2, replacements);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Xdr.OneCellAnchor anchor = Assert.Single(worksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
                var formulas = worksheetPart.DrawingsPart.ChartParts.Single().ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()
                    .Select(formula => formula.Text)
                    .ToArray();
                PageSetup setup = worksheetPart.Worksheet.GetFirstChild<PageSetup>()!;

                Assert.Equal("8", anchor.FromMarker!.RowId!.Text);
                Assert.Contains(formulas, formula => formula == "'Invoice'!$B$5");
                Assert.Contains(formulas, formula => formula == "'Invoice'!$A$6:$A$7");
                Assert.Contains(formulas, formula => formula == "'Invoice'!$B$6:$B$7");
                Assert.Equal(1U, setup.FitToWidth!.Value);
                Assert.Equal(0U, setup.FitToHeight!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRebasesCrossSheetRelativeFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsCrossSheetFormulas.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var invoice = document.AddWorkSheet("Invoice");
                document.AddWorkSheet("Inputs");
                invoice.CellAt(2, 1).SetValue("{{Name}}");
                invoice.CellFormula(2, 2, "'Inputs'!B2+'Inputs'!$C$2");

                invoice.ApplyTemplateRows(2, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting"
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support"
                    }
                });
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, "Invoice");
                var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

                Assert.Equal("'Inputs'!B2+'Inputs'!$C$2", cells["B2"].CellFormula!.Text);
                Assert.Equal("'Inputs'!B3+'Inputs'!$C$2", cells["B3"].CellFormula!.Text);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_OptionalRowsCanBeIncludedAndBound() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.OptionalRowsIncluded.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellAt(2, 1).SetValue("Discount");
                sheet.CellAt(2, 2).SetValue("{{Discount:currency}}");
                sheet.CellAt(3, 1).SetValue("Reason");
                sheet.CellAt(3, 2).SetValue("{{Reason}}");
                sheet.CellAt(4, 1).SetValue("Footer");

                int replacements = sheet.ApplyTemplateOptionalRows(2, 2, include: true, new {
                    Discount = 25m,
                    Reason = "Loyalty"
                }, new ExcelTemplateOptions {
                    FormatProvider = System.Globalization.CultureInfo.GetCultureInfo("en-US"),
                    MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
                });

                Assert.Equal(2, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Invoice"];
                Assert.Equal("Header", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Discount", sheet.CellAt(2, 1).GetValue<string>());
                Assert.Equal(25d, sheet.CellAt(2, 2).GetValue<double>());
                Assert.Equal("Reason", sheet.CellAt(3, 1).GetValue<string>());
                Assert.Equal("Loyalty", sheet.CellAt(3, 2).GetValue<string>());
                Assert.Equal("Footer", sheet.CellAt(4, 1).GetValue<string>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_OptionalRowsCanBeRemovedAndShiftFollowingRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.OptionalRowsRemoved.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellAt(2, 1).SetValue("Optional {{Note}}");
                sheet.CellAt(3, 1).SetValue("Optional {{Amount:currency}}");
                sheet.CellAt(4, 1).SetValue("Footer");

                int replacements = sheet.RemoveTemplateOptionalRows(2, 2);

                Assert.Equal(0, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Invoice"];
                Assert.Equal("Header", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Footer", sheet.CellAt(2, 1).GetValue<string>());
                Assert.Null(sheet.CellAt(3, 1).GetValue<string>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RemovingOptionalRowsRebasesChartsAndKeepsPageSetup() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.OptionalRowsCharts.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellAt(2, 1).SetValue("Optional {{Note}}");
                sheet.CellAt(3, 1).SetValue("Optional detail");
                sheet.CellAt(5, 1).SetValue("Region");
                sheet.CellAt(5, 2).SetValue("Amount");
                sheet.CellAt(6, 1).SetValue("EU");
                sheet.CellAt(6, 2).SetValue(10);
                sheet.CellAt(7, 1).SetValue("US");
                sheet.CellAt(7, 2).SetValue(20);
                sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 0U);
                sheet.AddChartFromRange("A5:B7", row: 9, column: 4, widthPixels: 320, heightPixels: 180, type: ExcelChartType.ColumnClustered, hasHeaders: true, title: "Sales");

                int replacements = sheet.RemoveTemplateOptionalRows(2, 2);

                Assert.Equal(0, replacements);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Xdr.OneCellAnchor anchor = Assert.Single(worksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
                var formulas = worksheetPart.DrawingsPart.ChartParts.Single().ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()
                    .Select(formula => formula.Text)
                    .ToArray();
                PageSetup setup = worksheetPart.Worksheet.GetFirstChild<PageSetup>()!;

                Assert.Equal("6", anchor.FromMarker!.RowId!.Text);
                Assert.Contains(formulas, formula => formula == "'Invoice'!$B$3");
                Assert.Contains(formulas, formula => formula == "'Invoice'!$A$4:$A$5");
                Assert.Contains(formulas, formula => formula == "'Invoice'!$B$4:$B$5");
                Assert.Equal(1U, setup.FitToWidth!.Value);
                Assert.Equal(0U, setup.FitToHeight!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet).ToList());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RemovingOptionalRowsRewritesMovedFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.OptionalRowsRemovedFormulas.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellAt(2, 1).SetValue("Optional {{Note}}");
                sheet.CellAt(3, 1).SetValue("Optional {{Amount}}");
                sheet.CellFormula(4, 1, "B4");

                sheet.RemoveTemplateOptionalRows(2, 2);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "A2");

                Assert.Equal("B2", cell.CellFormula!.Text);
                Assert.DoesNotContain(wsPart.Worksheet.Descendants<Cell>(), item => item.CellReference?.Value == "A4");
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RemovingOptionalRowsRebasesStationaryFormulasAndBoundaryMerges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.OptionalRowsRemovedStationaryFormulas.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellFormula(1, 4, "A4+$A$4");
                sheet.CellFormula(1, 5, "'Invoice'!$A$4");
                sheet.CellFormula(1, 6, "$A$2");
                sheet.CellFormula(1, 7, "IF(\"A4\"=\"A4\",A4,0)");
                sheet.CellFormula(1, 8, "SUM(A2:A5)");
                sheet.CellFormula(1, 9, "SUM(A1:A3)");
                sheet.CellFormula(1, 10, "SUM(A3:A4)");
                sheet.CellAt(2, 1).SetValue("Optional {{Note}}");
                sheet.CellAt(3, 1).SetValue("Optional {{Amount}}");
                sheet.CellAt(4, 1).SetValue("Footer");
                sheet.MergeRange("B1:B4");
                sheet.MergeRange("C2:C3");

                sheet.RemoveTemplateOptionalRows(2, 2);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                MergeCell merge = Assert.Single(worksheetPart.Worksheet.Descendants<MergeCell>());

                Assert.Equal("A2+$A$2", cells["D1"].CellFormula!.Text);
                Assert.Equal("'Invoice'!$A$2", cells["E1"].CellFormula!.Text);
                Assert.Equal("#REF!", cells["F1"].CellFormula!.Text);
                Assert.Equal("IF(\"A4\"=\"A4\",A2,0)", cells["G1"].CellFormula!.Text);
                Assert.Equal("SUM(A2:A3)", cells["H1"].CellFormula!.Text);
                Assert.Equal("SUM(A1:A1)", cells["I1"].CellFormula!.Text);
                Assert.Equal("SUM(A2:A2)", cells["J1"].CellFormula!.Text);
                Assert.Equal("B1:B2", merge.Reference!.Value);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRemapsRowBoundMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsMetadata.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellAt(2, 1).SetValue("{{Name}}");
                sheet.CellAt(3, 1).SetValue("Footer");
                sheet.CellAt(3, 2).SetValue("Choice");
                sheet.CellAt(3, 3).SetValue(1);
                sheet.SetComment(3, 1, "Footer note");
                sheet.CellAt(4, 1).SetValue("Trailing");
                sheet.SetComment(4, 1, "Trailing note");
                sheet.SetHyperlink(3, 1, "https://example.org");
                sheet.AddConditionalFormulaRule("C3", "C3>0");
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var worksheet = worksheetPart.Worksheet;
                var validations = worksheet.GetFirstChild<DataValidations>() ?? worksheet.InsertAfter(new DataValidations(), worksheet.GetFirstChild<SheetData>());
                validations.Append(new DataValidation {
                    Type = DataValidationValues.List,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "B3" }
                });
                validations.Count = 1U;
                worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document["Invoice"];
                sheet.ApplyTemplateRows(2, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting"
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support"
                    }
                });
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var worksheet = worksheetPart.Worksheet;
                var comments = worksheetPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>()
                    .OrderBy(comment => comment.Reference!.Value)
                    .ToList();
                Hyperlink hyperlink = Assert.Single(worksheet.Elements<Hyperlinks>().Single().Elements<Hyperlink>());
                DataValidation validation = Assert.Single(worksheet.GetFirstChild<DataValidations>()!.Elements<DataValidation>());
                ConditionalFormatting conditional = Assert.Single(worksheet.Elements<ConditionalFormatting>());
                VmlDrawingPart vmlPart = Assert.Single(worksheetPart.VmlDrawingParts);
                XDocument vml = XDocument.Load(vmlPart.GetStream());
                XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
                string[] vmlCoordinates = vml.Root!.Descendants(excelNamespace + "ClientData")
                    .Select(clientData => string.Join(
                        ",",
                        clientData.Element(excelNamespace + "Row")?.Value.Trim(),
                        clientData.Element(excelNamespace + "Column")?.Value.Trim()))
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToArray();

                Assert.Equal(new[] { "A4", "A5" }, comments.Select(comment => comment.Reference!.Value).ToArray());
                Assert.Equal(new[] { "3,0", "4,0" }, vmlCoordinates);
                Assert.Equal("A4", hyperlink.Reference!.Value);
                Assert.Equal("B4", validation.SequenceOfReferences!.InnerText);
                Assert.Equal("C4", conditional.SequenceOfReferences!.InnerText);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatingRowsRebasesNamedRangesAndTables() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingRowsNamedRangesTables.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellAt(2, 1).SetValue("{{Name}}");
                sheet.CellAt(4, 1).SetValue("Code");
                sheet.CellAt(4, 2).SetValue("Amount");
                sheet.CellAt(5, 1).SetValue("A");
                sheet.CellAt(5, 2).SetValue(10);
                sheet.AddTable("A4:B5", hasHeader: true, name: "SalesTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                document.SetNamedRange("GlobalSales", "'Invoice'!A4:B5", save: false);
                sheet.SetNamedRange("LocalSales", "A4:B5", save: false);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document["Invoice"];
                int replacements = sheet.ApplyTemplateRows(2, new[] {
                    new Dictionary<string, object?> {
                        ["Name"] = "Consulting"
                    },
                    new Dictionary<string, object?> {
                        ["Name"] = "Support"
                    }
                });

                Assert.Equal(2, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Invoice"];
                Assert.Equal("'Invoice'!$A$5:$B$6", document.GetNamedRange("GlobalSales"));
                Assert.Equal("$A$5:$B$6", sheet.GetNamedRange("LocalSales"));
                Assert.Empty(document.ValidateOpenXml());
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, "Invoice");
                TableDefinitionPart tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
                Assert.Equal("A5:B6", tablePart.Table.Reference!.Value);
                Assert.Equal("A5:B6", tablePart.Table.GetFirstChild<AutoFilter>()!.Reference!.Value);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RemovingOptionalRowsRebasesNamedRangesAndTables() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.OptionalRowsNamedRangesTables.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Header");
                sheet.CellAt(2, 1).SetValue("Optional {{Note}}");
                sheet.CellAt(3, 1).SetValue("Optional detail");
                sheet.CellAt(5, 1).SetValue("Code");
                sheet.CellAt(5, 2).SetValue("Amount");
                sheet.CellAt(6, 1).SetValue("A");
                sheet.CellAt(6, 2).SetValue(10);
                sheet.AddTable("A5:B6", hasHeader: true, name: "SalesTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                document.SetNamedRange("GlobalSales", "'Invoice'!A5:B6", save: false);
                sheet.SetNamedRange("LocalSales", "A5:B6", save: false);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document["Invoice"];
                int replacements = sheet.RemoveTemplateOptionalRows(2, 2);

                Assert.Equal(0, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Invoice"];
                Assert.Equal("'Invoice'!$A$3:$B$4", document.GetNamedRange("GlobalSales"));
                Assert.Equal("$A$3:$B$4", sheet.GetNamedRange("LocalSales"));
                Assert.Empty(document.ValidateOpenXml());
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, "Invoice");
                TableDefinitionPart tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
                Assert.Equal("A3:B4", tablePart.Table.Reference!.Value);
                Assert.Equal("A3:B4", tablePart.Table.GetFirstChild<AutoFilter>()!.Reference!.Value);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatsTemplateSheetForModels() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingSheets.xlsx");
            byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Region Template");
                sheet.CellAt(1, 1).SetValue("Region {{Region}}").HeaderStyle();
                sheet.CellAt(2, 1).SetValue("Owner");
                sheet.CellAt(2, 2).SetValue("{{Owner}}");
                sheet.CellAt(3, 1).SetValue("Amount");
                sheet.CellAt(3, 2).SetValue("{{Amount:currency}}");
                sheet.Range("B3").Validation.DecimalBetween(0, 1000, allowBlank: false, errorTitle: "Amount", errorMessage: "Use amount 0-1000");
                sheet.Range("C3").Validation.CustomFormula("'Region Template'!B3>0", allowBlank: false, errorTitle: "Amount", errorMessage: "Use positive amount");
                sheet.Range("B3").ConditionalFormatting.GreaterThan("'Region Template'!B3");
                sheet.CellFormula(4, 2, "'Region Template'!B3*2");
                sheet.CellFormula(5, 2, "SUM(TemplateWorkbookArea)");
                sheet.MergeRange("A1:B1");
                sheet.SetColumnWidth(1, 18);
                sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 0U);
                document.SetNamedRange("TemplateWorkbookArea", "'Region Template'!A1:B3", save: false);
                sheet.SetNamedRange("TemplateArea", "A1:B3", save: false);
                document.SetPrintArea(sheet, "A1:B7", save: false);
                document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
                sheet.CellAt(6, 1).SetValue("Owner");
                sheet.CellAt(6, 2).SetValue("Amount");
                sheet.CellAt(7, 1).SetValue("{{Owner}}");
                sheet.CellAt(7, 2).SetValue("{{Amount}}");
                sheet.AddTable("A6:B7", hasHeader: true, name: "RegionTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                sheet.SetHyperlink("A9", "https://example.com/help", "Help");
                sheet.SetInternalLink(11, 1, "'Region Template'!A1", "Back to top");
                sheet.SetComment("A10", "Template note", author: "Template");
                sheet.AddImage(10, 3, png, "image/png", widthPixels: 18, heightPixels: 14, name: "Static Logo", altText: "Static template logo");
                var chart = sheet.AddChartFromRange("A6:B7", row: 12, column: 4, widthPixels: 320, heightPixels: 180, type: ExcelChartType.ColumnClustered, hasHeaders: true, title: "Regional total");
                chart.ApplyStylePreset(ExcelChartStylePreset.Default);

                int replacements = document.ApplyTemplateSheets(
                    "Region Template",
                    new[] {
                        new RegionSheetTemplateModel {
                            Region = "North",
                            Owner = "Alice",
                            Amount = 125m
                        },
                        new RegionSheetTemplateModel {
                            Region = "South",
                            Owner = "Bob",
                            Amount = 250m
                        }
                    },
                    (model, index) => model.Region,
                    new ExcelTemplateOptions {
                        FormatProvider = System.Globalization.CultureInfo.GetCultureInfo("en-US"),
                        MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
                    });

                Assert.Equal(10, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Contains(document.Sheets, sheet => sheet.Name == "North");
                Assert.Contains(document.Sheets, sheet => sheet.Name == "South");

                var north = document["North"];
                var south = document["South"];
                Assert.Equal("Region North", north.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Alice", north.CellAt(2, 2).GetValue<string>());
                Assert.Equal(125d, north.CellAt(3, 2).GetValue<double>());
                Assert.Equal("Alice", north.CellAt(7, 1).GetValue<string>());
                Assert.Equal("$A$1:$B$3", north.GetNamedRange("TemplateArea"));
                Assert.Equal("Region South", south.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Bob", south.CellAt(2, 2).GetValue<string>());
                Assert.Equal(250d, south.CellAt(3, 2).GetValue<double>());
                Assert.Equal("Bob", south.CellAt(7, 1).GetValue<string>());
                Assert.Equal("$A$1:$B$3", south.GetNamedRange("TemplateArea"));
                Assert.Empty(document.ValidateOpenXml());
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart northPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, "North");
                WorksheetPart southPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, "South");

                foreach (var generatedSheet in new[] { new { Name = "North", Part = northPart }, new { Name = "South", Part = southPart } }) {
                    WorksheetPart worksheetPart = generatedSheet.Part;
                    var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                    MergeCell merge = Assert.Single(worksheetPart.Worksheet.Descendants<MergeCell>());
                    PageSetup setup = worksheetPart.Worksheet.GetFirstChild<PageSetup>()!;
                    var validations = worksheetPart.Worksheet.GetFirstChild<DataValidations>()!.Elements<DataValidation>().ToArray();
                    DataValidation validation = Assert.Single(validations, item => item.Type?.Value == DataValidationValues.Decimal);
                    DataValidation customValidation = Assert.Single(validations, item => item.Type?.Value == DataValidationValues.Custom);
                    ConditionalFormatting conditionalFormatting = Assert.Single(worksheetPart.Worksheet.Elements<ConditionalFormatting>());
                    ConditionalFormattingRule conditionalRule = Assert.Single(conditionalFormatting.Elements<ConditionalFormattingRule>());
                    TableParts tableParts = worksheetPart.Worksheet.GetFirstChild<TableParts>()!;
                    TableDefinitionPart tablePart = Assert.Single(worksheetPart.TableDefinitionParts);
                    var hyperlinks = worksheetPart.Worksheet.Descendants<Hyperlink>().ToArray();
                    Hyperlink hyperlink = Assert.Single(hyperlinks, item => item.Id != null);
                    Hyperlink internalHyperlink = Assert.Single(hyperlinks, item => item.Location != null);
                    HyperlinkRelationship hyperlinkRelationship = Assert.Single(worksheetPart.HyperlinkRelationships);
                    Comment comment = Assert.Single(worksheetPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>());
                    LegacyDrawing legacyDrawing = Assert.Single(worksheetPart.Worksheet.Descendants<LegacyDrawing>());
                    VmlDrawingPart vmlDrawingPart = Assert.Single(worksheetPart.VmlDrawingParts);
                    DocumentFormat.OpenXml.Spreadsheet.Drawing drawing = Assert.Single(worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Drawing>());
                    DrawingsPart drawingsPart = worksheetPart.DrawingsPart!;
                    Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
                    ushort localSheetId = (ushort)workbook.Sheets!.Elements<Sheet>()
                        .Select((sheet, index) => new { Sheet = sheet, Index = index })
                        .Single(item => string.Equals(item.Sheet.Name?.Value, generatedSheet.Name, StringComparison.Ordinal))
                        .Index;
                    var localDefinedNames = workbook.DefinedNames!.Elements<DefinedName>()
                        .Where(name => name.LocalSheetId != null && name.LocalSheetId.Value == localSheetId)
                        .ToArray();
                    var workbookDefinedNames = workbook.DefinedNames!.Elements<DefinedName>()
                        .Where(name => name.LocalSheetId == null)
                        .ToArray();
                    ChartPart chartPart = Assert.Single(drawingsPart.ChartParts);
                    ChartStylePart chartStylePart = Assert.Single(chartPart.GetPartsOfType<ChartStylePart>());
                    ChartColorStylePart chartColorStylePart = Assert.Single(chartPart.GetPartsOfType<ChartColorStylePart>());
                    DocumentFormat.OpenXml.Drawing.Charts.ChartReference chartReference =
                        Assert.Single(drawingsPart.WorksheetDrawing!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>());
                    Xdr.Picture picture = Assert.Single(drawingsPart.WorksheetDrawing.Descendants<Xdr.Picture>());
                    DocumentFormat.OpenXml.Drawing.Blip blip = Assert.Single(picture.Descendants<DocumentFormat.OpenXml.Drawing.Blip>());
                    ImagePart imagePart = Assert.Single(drawingsPart.ImageParts);
                    var chartFormulas = chartPart.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()
                        .Select(formula => formula.Text)
                        .ToArray();

                    Assert.Equal($"'{generatedSheet.Name}'!B3*2", cells["B4"].CellFormula!.Text);
                    string workbookDefinedName = generatedSheet.Name == "North" ? "TemplateWorkbookArea" : $"TemplateWorkbookArea_{generatedSheet.Name}";
                    Assert.Equal($"SUM({workbookDefinedName})", cells["B5"].CellFormula!.Text);
                    Assert.Equal("A1:B1", merge.Reference!.Value);
                    Assert.Equal(1U, setup.FitToWidth!.Value);
                    Assert.Equal(0U, setup.FitToHeight!.Value);
                    Assert.Equal("B3:B3", validation.SequenceOfReferences!.InnerText);
                    Assert.Equal(DataValidationValues.Decimal, validation.Type!.Value);
                    Assert.Equal(DataValidationOperatorValues.Between, validation.Operator!.Value);
                    Assert.False(validation.AllowBlank!.Value);
                    Assert.Equal("0", validation.Formula1!.Text);
                    Assert.Equal("1000", validation.Formula2!.Text);
                    Assert.Equal("C3:C3", customValidation.SequenceOfReferences!.InnerText);
                    Assert.False(customValidation.AllowBlank!.Value);
                    Assert.Equal($"'{generatedSheet.Name}'!B3>0", customValidation.Formula1!.Text);
                    Assert.Equal("B3:B3", conditionalFormatting.SequenceOfReferences!.InnerText);
                    Assert.Equal(ConditionalFormatValues.CellIs, conditionalRule.Type!.Value);
                    Assert.Equal(ConditionalFormattingOperatorValues.GreaterThan, conditionalRule.Operator!.Value);
                    Assert.Equal($"'{generatedSheet.Name}'!B3", Assert.Single(conditionalRule.Elements<Formula>()).Text);
                    Assert.Contains(localDefinedNames, name => name.Name == "TemplateArea" && name.Text == $"'{generatedSheet.Name}'!$A$1:$B$3");
                    Assert.Contains(localDefinedNames, name => name.Name == "_xlnm.Print_Area" && name.Text == $"'{generatedSheet.Name}'!$A$1:$B$7");
                    Assert.Contains(localDefinedNames, name => name.Name == "_xlnm.Print_Titles" && name.Text == $"'{generatedSheet.Name}'!$1:$1");
                    Assert.Contains(workbookDefinedNames, name => name.Name == workbookDefinedName && name.Text == $"'{generatedSheet.Name}'!$A$1:$B$3");
                    Assert.Equal("A6:B7", tablePart.Table.Reference!.Value);
                    Assert.Equal("A6:B7", tablePart.Table.GetFirstChild<AutoFilter>()!.Reference!.Value);
                    Assert.Equal(1U, tableParts.Count!.Value);
                    Assert.Equal("A9", hyperlink.Reference!.Value);
                    Assert.Equal(hyperlinkRelationship.Id, hyperlink.Id!.Value);
                    Assert.Equal("https://example.com/help", hyperlinkRelationship.Uri.ToString());
                    Assert.Equal("A11", internalHyperlink.Reference!.Value);
                    Assert.Equal($"'{generatedSheet.Name}'!A1", internalHyperlink.Location!.Value);
                    Assert.Equal("A10", comment.Reference!.Value);
                    Assert.Equal("Template note", comment.InnerText);
                    Assert.Equal(worksheetPart.GetIdOfPart(vmlDrawingPart), legacyDrawing.Id!.Value);
                    Assert.Equal(worksheetPart.GetIdOfPart(drawingsPart), drawing.Id!.Value);
                    Assert.Equal(drawingsPart.GetIdOfPart(chartPart), chartReference.Id!.Value);
                    Assert.False(string.IsNullOrWhiteSpace(chartPart.GetIdOfPart(chartStylePart)));
                    Assert.False(string.IsNullOrWhiteSpace(chartPart.GetIdOfPart(chartColorStylePart)));
                    Assert.Equal(drawingsPart.GetIdOfPart(imagePart), blip.Embed!.Value);
                    Assert.Equal("Static Logo", picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name);
                    Assert.Equal("Static template logo", picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description);
                    Assert.Contains(chartFormulas, formula => formula == $"'{generatedSheet.Name}'!$B$6");
                    Assert.Contains(chartFormulas, formula => formula == $"'{generatedSheet.Name}'!$A$7:$A$7");
                    Assert.Contains(chartFormulas, formula => formula == $"'{generatedSheet.Name}'!$B$7:$B$7");
                    Assert.DoesNotContain(chartFormulas, formula => formula != null && formula.Contains("Region Template", StringComparison.Ordinal));
                    Assert.DoesNotContain("{{", worksheetPart.Worksheet.OuterXml, StringComparison.Ordinal);
                }

                string[] tableNames = spreadsheet.WorkbookPart!.WorksheetParts
                    .SelectMany(part => part.TableDefinitionParts)
                    .Select(part => part.Table!.Name!.Value!)
                    .OrderBy(name => name, StringComparer.Ordinal)
                    .ToArray();
                Assert.Equal(2, tableNames.Length);
                Assert.Equal(tableNames.Length, tableNames.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatsTemplateSheetPreservesChartEmbeddedPackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingSheets.ChartEmbeddedPackage.xlsx");
            byte[] embeddedBytes = { 0x50, 0x4B, 0x03, 0x04, 0x4F, 0x49, 0x4D, 0x4F };

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Chart Template");
                sheet.CellAt(1, 1).SetValue("Region {{Region}}");
                sheet.CellAt(3, 1).SetValue("Metric");
                sheet.CellAt(3, 2).SetValue("Value");
                sheet.CellAt(4, 1).SetValue("Sales");
                sheet.CellAt(4, 2).SetValue("{{Amount}}");
                sheet.AddChartFromRange("A3:B4", row: 6, column: 3, widthPixels: 320, heightPixels: 180, type: ExcelChartType.ColumnClustered, hasHeaders: true, title: "Regional total");
                document.Save(false);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                ChartPart sourceChartPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
                EmbeddedPackagePart sourceEmbeddedPackage = sourceChartPart.AddEmbeddedPackagePart(EmbeddedPackagePartType.Xlsx);
                using var stream = new MemoryStream(embeddedBytes);
                sourceEmbeddedPackage.FeedData(stream);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                int replacements = document.ApplyTemplateSheets(
                    "Chart Template",
                    new[] {
                        new RegionSheetTemplateModel {
                            Region = "North",
                            Amount = 125m
                        },
                        new RegionSheetTemplateModel {
                            Region = "South",
                            Amount = 250m
                        }
                    },
                    (model, index) => model.Region,
                    new ExcelTemplateOptions {
                        MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
                    });

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                foreach (string sheetName in new[] { "North", "South" }) {
                    WorksheetPart worksheetPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, sheetName);
                    ChartPart chartPart = Assert.Single(worksheetPart.DrawingsPart!.ChartParts);
                    EmbeddedPackagePart embeddedPackagePart = Assert.Single(chartPart.GetPartsOfType<EmbeddedPackagePart>());
                    using Stream embeddedStream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read);
                    using var buffer = new MemoryStream();
                    embeddedStream.CopyTo(buffer);

                    Assert.Equal(embeddedBytes, buffer.ToArray());
                }
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatsTemplateSheetPreservesChartDrawingParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingSheets.ChartDrawingPart.xlsx");
            byte[] imageBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Chart Template");
                sheet.CellAt(1, 1).SetValue("Region {{Region}}");
                sheet.CellAt(3, 1).SetValue("Metric");
                sheet.CellAt(3, 2).SetValue("Value");
                sheet.CellAt(4, 1).SetValue("Sales");
                sheet.CellAt(4, 2).SetValue("{{Amount}}");
                sheet.AddChartFromRange("A3:B4", row: 6, column: 3, widthPixels: 320, heightPixels: 180, type: ExcelChartType.ColumnClustered, hasHeaders: true, title: "Regional total");
                document.Save(false);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                ChartPart sourceChartPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
                string chartDrawingRelationshipId = "rIdUserShapes";
                ChartDrawingPart sourceChartDrawingPart = sourceChartPart.AddNewPart<ChartDrawingPart>(chartDrawingRelationshipId);
                new DocumentFormat.OpenXml.Drawing.Charts.UserShapes().Save(sourceChartDrawingPart);
                ImagePart sourceImagePart = sourceChartDrawingPart.AddImagePart(ImagePartType.Png, "rIdImage1");
                using (var stream = new MemoryStream(imageBytes)) {
                    sourceImagePart.FeedData(stream);
                }

                sourceChartPart.ChartSpace!.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.UserShapesReference {
                    Id = chartDrawingRelationshipId
                });
                sourceChartPart.ChartSpace.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                int replacements = document.ApplyTemplateSheets(
                    "Chart Template",
                    new[] {
                        new RegionSheetTemplateModel {
                            Region = "North",
                            Amount = 125m
                        },
                        new RegionSheetTemplateModel {
                            Region = "South",
                            Amount = 250m
                        }
                    },
                    (model, index) => model.Region,
                    new ExcelTemplateOptions {
                        MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
                    });

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                foreach (string sheetName in new[] { "North", "South" }) {
                    WorksheetPart worksheetPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, sheetName);
                    ChartPart chartPart = Assert.Single(worksheetPart.DrawingsPart!.ChartParts);
                    DocumentFormat.OpenXml.Drawing.Charts.UserShapesReference userShapesReference =
                        Assert.Single(chartPart.ChartSpace!.Elements<DocumentFormat.OpenXml.Drawing.Charts.UserShapesReference>());
                    ChartDrawingPart chartDrawingPart = Assert.Single(chartPart.GetPartsOfType<ChartDrawingPart>());
                    ImagePart imagePart = Assert.Single(chartDrawingPart.ImageParts);

                    using Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                    using var buffer = new MemoryStream();
                    imageStream.CopyTo(buffer);

                    Assert.Equal(chartPart.GetIdOfPart(chartDrawingPart), userShapesReference.Id!.Value);
                    Assert.NotNull(chartDrawingPart.UserShapes);
                    Assert.Equal(imageBytes, buffer.ToArray());
                }
            }
        }

        [Fact]
        public void Test_ExcelTemplate_RepeatsTemplateSheetPreservesDiagramDrawingParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.RepeatingSheets.DiagramDrawingParts.xlsx");
            byte[] imageBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
            string dataXml = "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><dgm:ptLst /><dgm:cxnLst /><dgm:bg /><dgm:whole /></dgm:dataModel>";
            string layoutXml = "<dgm:layoutDef uniqueId=\"urn:officeimo:test:layout\" xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><dgm:title val=\"\" /><dgm:desc val=\"\" /><dgm:catLst><dgm:cat type=\"list\" pri=\"400\" /></dgm:catLst><dgm:layoutNode name=\"diagram\" /></dgm:layoutDef>";
            string colorsXml = "<dgm:colorsDef uniqueId=\"urn:officeimo:test:colors\" xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><dgm:title val=\"\" /><dgm:desc val=\"\" /><dgm:catLst><dgm:cat type=\"accent1\" pri=\"11200\" /></dgm:catLst></dgm:colorsDef>";
            string styleXml = "<dgm:styleDef uniqueId=\"urn:officeimo:test:style\" xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><dgm:title val=\"\" /><dgm:desc val=\"\" /><dgm:catLst><dgm:cat type=\"simple\" pri=\"10100\" /></dgm:catLst></dgm:styleDef>";

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Diagram Template");
                sheet.CellAt(1, 1).SetValue("Region {{Region}}");
                sheet.CellAt(3, 1).SetValue("Metric");
                sheet.CellAt(3, 2).SetValue("Value");
                sheet.CellAt(4, 1).SetValue("Sales");
                sheet.CellAt(4, 2).SetValue("{{Amount}}");
                sheet.AddChartFromRange("A3:B4", row: 6, column: 3, widthPixels: 320, heightPixels: 180, type: ExcelChartType.ColumnClustered, hasHeaders: true, title: "Regional total");
                document.Save(false);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                DrawingsPart drawingsPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
                DiagramLayoutDefinitionPart layoutPart = drawingsPart.AddNewPart<DiagramLayoutDefinitionPart>("rIdDiagramLayout");
                DiagramColorsPart colorsPart = drawingsPart.AddNewPart<DiagramColorsPart>("rIdDiagramColors");
                DiagramStylePart stylePart = drawingsPart.AddNewPart<DiagramStylePart>("rIdDiagramStyle");
                DiagramDataPart dataPart = drawingsPart.AddNewPart<DiagramDataPart>("rIdDiagramData");
                FeedPart(layoutPart, layoutXml);
                FeedPart(colorsPart, colorsXml);
                FeedPart(stylePart, styleXml);
                FeedPart(dataPart, dataXml);
                ImagePart dataImagePart = dataPart.AddImagePart(ImagePartType.Png, "rIdDataImage");
                using (var stream = new MemoryStream(imageBytes)) {
                    dataImagePart.FeedData(stream);
                }

                drawingsPart.WorksheetDrawing!.Append(CreateDiagramAnchor(
                    drawingsPart.GetIdOfPart(layoutPart)!,
                    drawingsPart.GetIdOfPart(colorsPart)!,
                    drawingsPart.GetIdOfPart(stylePart)!,
                    drawingsPart.GetIdOfPart(dataPart)!));
                drawingsPart.WorksheetDrawing.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                int replacements = document.ApplyTemplateSheets(
                    "Diagram Template",
                    new[] {
                        new RegionSheetTemplateModel {
                            Region = "North",
                            Amount = 125m
                        },
                        new RegionSheetTemplateModel {
                            Region = "South",
                            Amount = 250m
                        }
                    },
                    (model, index) => model.Region,
                    new ExcelTemplateOptions {
                        MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
                    });

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                foreach (string sheetName in new[] { "North", "South" }) {
                    WorksheetPart worksheetPart = GetWorksheetPartByNameForTemplateTests(spreadsheet, sheetName);
                    DrawingsPart drawingsPart = worksheetPart.DrawingsPart!;
                    Dgm.RelationshipIds relationships =
                        Assert.Single(drawingsPart.WorksheetDrawing!.Descendants<Dgm.RelationshipIds>());
                    DiagramLayoutDefinitionPart layoutPart = Assert.Single(drawingsPart.DiagramLayoutDefinitionParts);
                    DiagramColorsPart colorsPart = Assert.Single(drawingsPart.DiagramColorsParts);
                    DiagramStylePart stylePart = Assert.Single(drawingsPart.DiagramStyleParts);
                    DiagramDataPart dataPart = Assert.Single(drawingsPart.DiagramDataParts);
                    ImagePart dataImagePart = Assert.Single(dataPart.ImageParts);

                    using Stream imageStream = dataImagePart.GetStream(FileMode.Open, FileAccess.Read);
                    using var buffer = new MemoryStream();
                    imageStream.CopyTo(buffer);

                    Assert.Equal(drawingsPart.GetIdOfPart(layoutPart), relationships.LayoutPart);
                    Assert.Equal(drawingsPart.GetIdOfPart(colorsPart), relationships.ColorPart);
                    Assert.Equal(drawingsPart.GetIdOfPart(stylePart), relationships.StylePart);
                    Assert.Equal(drawingsPart.GetIdOfPart(dataPart), relationships.DataPart);
                    Assert.Equal(layoutXml, ReadPartText(layoutPart));
                    Assert.Equal(colorsXml, ReadPartText(colorsPart));
                    Assert.Equal(styleXml, ReadPartText(stylePart));
                    Assert.Equal(dataXml, ReadPartText(dataPart));
                    Assert.Equal(imageBytes, buffer.ToArray());
                }
            }
        }

        [Fact]
        public void Test_ExcelTemplate_ImageMarkersBindBytesAndStreams() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.ImageMarkers.xlsx");
            byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("{{Logo}}");
                sheet.CellAt(3, 1).SetValue("{{Badge}}");

                using var badgeStream = new MemoryStream(png);
                int replacements = sheet.ApplyTemplate(new Dictionary<string, object?> {
                    ["Logo"] = ExcelTemplateImage.FromBytes(png, widthPixels: 24, heightPixels: 18, name: "Logo", altText: "Company logo"),
                    ["Badge"] = ExcelTemplateImage.FromStream(badgeStream, widthPixels: 12, heightPixels: 10, name: "Badge", altText: "Status badge")
                });

                Assert.Equal(2, replacements);
                Assert.Equal(2, sheet.Images.Count());
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var pictures = wsPart.DrawingsPart!.WorksheetDrawing!.Descendants<Xdr.Picture>().ToList();
                var extents = wsPart.DrawingsPart.WorksheetDrawing.Descendants<Xdr.Extent>().ToList();

                Assert.Equal(2, pictures.Count);
                Assert.Contains(pictures, picture => picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name == "Logo");
                Assert.Contains(pictures, picture => picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name == "Badge");
                Assert.Contains(extents, extent => extent.Cx!.Value == 24L * 9525L && extent.Cy!.Value == 18L * 9525L);
                Assert.Contains(extents, extent => extent.Cx!.Value == 12L * 9525L && extent.Cy!.Value == 10L * 9525L);
                Assert.DoesNotContain("{{", wsPart.Worksheet!.OuterXml, StringComparison.Ordinal);
                Assert.Equal(2, wsPart.DrawingsPart.ImageParts.Count());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_ImageFromUrlDefaultsToSharedPngContentType() {
            var image = ExcelTemplateImage.FromUrl("https://example.test/logo.png");

            Assert.Equal(OfficeImageInfo.GetMimeType(OfficeImageFormat.Png), image.ContentType);
        }

        [Fact]
        public void Test_ExcelTemplate_BindsObjectModelAndFormatAliases() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.ObjectModel.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("Customer: {{Customer.Name}}");
                sheet.CellAt(2, 1).SetValue("Total: {{Total:currency}}");
                sheet.CellAt(3, 1).SetValue("Issued: {{Issued:yyyy-MM-dd}}");
                sheet.CellAt(4, 1).SetValue("Completion: {{Completion:percent}}");

                int replacements = document.ApplyTemplate(new InvoiceTemplateModel {
                    Customer = new CustomerTemplateModel { Name = "Adatum" },
                    Total = 123.45m,
                    Issued = new DateTime(2026, 5, 21),
                    Completion = 0.5
                }, System.Globalization.CultureInfo.GetCultureInfo("en-US"));

                Assert.Equal(4, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Invoice"];
                Assert.Equal("Customer: Adatum", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Total: $123.45", sheet.CellAt(2, 1).GetValue<string>());
                Assert.Equal("Issued: 2026-05-21", sheet.CellAt(3, 1).GetValue<string>());
                Assert.Equal("Completion: 50.00%", sheet.CellAt(4, 1).GetValue<string>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_OptionsSupportCustomFormatters() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.CustomFormatters.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("Customer {{Name:upper}}");
                sheet.CellAt(2, 1).SetValue("Risk {{Score:risk}}");
                sheet.CellAt(3, 1).SetValue("{{Nullable:fallback}}");

                var options = new ExcelTemplateOptions {
                    FormatProvider = System.Globalization.CultureInfo.InvariantCulture,
                    ThrowOnMissing = true
                }
                    .AddFormatter("upper", (value, provider) =>
                        Convert.ToString(value, provider as System.Globalization.CultureInfo ?? System.Globalization.CultureInfo.CurrentCulture)?.ToUpperInvariant() ?? string.Empty)
                    .AddFormatter("risk", (value, provider) =>
                        Convert.ToDouble(value, provider as System.Globalization.CultureInfo ?? System.Globalization.CultureInfo.CurrentCulture) >= 80 ? "High" : "Low")
                    .AddFormatter("fallback", (value, provider) =>
                        value == null ? "n/a" : Convert.ToString(value, provider as System.Globalization.CultureInfo ?? System.Globalization.CultureInfo.CurrentCulture) ?? string.Empty);

                int replacements = sheet.ApplyTemplate(new Dictionary<string, object?> {
                    ["Name"] = "Adatum",
                    ["Score"] = 87,
                    ["Nullable"] = null
                }, options);

                Assert.Equal(3, replacements);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document["Template"];
                Assert.Equal("Customer ADATUM", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Risk High", sheet.CellAt(2, 1).GetValue<string>());
                Assert.Equal("n/a", sheet.CellAt(3, 1).GetValue<string>());
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelTemplate_WholeCellMarkersWriteTypedValuesAndNumberFormats() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.TypedMarkers.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Invoice");
                sheet.CellAt(1, 1).SetValue("{{Total:currency}}");
                sheet.CellAt(2, 1).SetValue("{{Completion:percent}}");
                sheet.CellAt(3, 1).SetValue("{{Issued:date}}");

                int replacements = sheet.ApplyTemplate(new {
                    Total = 123.45m,
                    Completion = 0.5,
                    Issued = new DateTime(2026, 5, 21)
                }, System.Globalization.CultureInfo.GetCultureInfo("en-US"));

                Assert.Equal(3, replacements);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                string stylesXml = spreadsheet.WorkbookPart.WorkbookStylesPart!.Stylesheet.OuterXml;

                Assert.Equal(CellValues.Number, cells["A1"].DataType!.Value);
                Assert.Equal("123.45", cells["A1"].CellValue!.Text);
                Assert.Equal(CellValues.Number, cells["A2"].DataType!.Value);
                Assert.Equal("0.5", cells["A2"].CellValue!.Text);
                Assert.Equal(CellValues.Number, cells["A3"].DataType!.Value);
                Assert.DoesNotContain("{{", wsPart.Worksheet.OuterXml, StringComparison.Ordinal);
                Assert.Contains("$", stylesXml, StringComparison.Ordinal);
                Assert.Contains("0.00%", stylesXml, StringComparison.Ordinal);
                Assert.Contains("yyyy-mm-dd", stylesXml, StringComparison.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_DurationAliasFormatsTextAndTypedCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.DurationAlias.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Schedule");
                sheet.CellAt(1, 1).SetValue("Elapsed {{Duration:duration}}");
                sheet.CellAt(2, 1).SetValue("{{Duration:duration}}");

                int replacements = sheet.ApplyTemplate(new {
                    Duration = TimeSpan.FromMinutes(1650)
                }, System.Globalization.CultureInfo.InvariantCulture);

                Assert.Equal(2, replacements);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet!.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
                string stylesXml = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.OuterXml;

                Assert.Equal("Elapsed 27:30:00", cells["A1"].InlineString!.Text!.Text);
                Assert.Equal(CellValues.Number, cells["A2"].DataType!.Value);
                Assert.Equal(TimeSpan.FromMinutes(1650).TotalDays.ToString(System.Globalization.CultureInfo.InvariantCulture), cells["A2"].CellValue!.Text);
                Assert.Contains("[h]:mm:ss", stylesXml, StringComparison.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Test_ExcelTemplate_InspectionReportsMarkersAndMissingBindings() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.Inspect.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("Invoice {{Invoice.Number}}");
                sheet.CellAt(2, 1).SetValue("{{Total:currency}}");
                sheet.CellAt(3, 1).SetValue("Customer {{Customer.Name}}");
                sheet.CellAt(4, 1).SetValue("{{Logo}}");
                sheet.CellAt(5, 1).SetValue("Issued {{Issued:date}}");

                ExcelTemplateInspection template = document.InspectTemplate();
                Assert.Equal(5, template.TotalMarkers);
                Assert.False(template.HasBindingInfo);
                Assert.Contains("Invoice.Number", template.UniqueMarkers);
                Assert.Contains(template.Markers, marker => marker.Name == "Total"
                    && marker.Format == "currency"
                    && marker.IsWholeCell
                    && marker.CellReference == "A2");

                ExcelTemplateInspection missing = document.InspectTemplate(new Dictionary<string, object?> {
                    ["Invoice.Number"] = "INV-001",
                    ["Total"] = 10,
                    ["Logo"] = ExcelTemplateImage.FromBytes(new byte[] { 1, 2, 3 }, widthPixels: 12, heightPixels: 12),
                    ["Issued"] = new DateTime(2026, 5, 28)
                });

                Assert.True(missing.HasBindingInfo);
                Assert.False(missing.AllMarkersBound);
                Assert.Single(missing.MissingMarkerNames);
                Assert.Equal("Customer.Name", missing.MissingMarkerNames[0]);
                Assert.Contains(missing.Markers, marker => marker.Name == "Total"
                    && marker.BoundValueKind == "number"
                    && marker.BoundValueTypeName == "Int32");
                Assert.Contains(missing.Markers, marker => marker.Name == "Logo"
                    && marker.BoundValueKind == "image"
                    && marker.BoundValueTypeName == "ExcelTemplateImage");
                Assert.Contains(missing.Markers, marker => marker.Name == "Issued"
                    && marker.BoundValueKind == "date/time"
                    && marker.BoundValueTypeName == "DateTime");
                InvalidOperationException missingException = Assert.Throws<InvalidOperationException>(() => missing.EnsureAllMarkersBound());
                Assert.Contains("Customer.Name", missingException.Message);

                string markdown = missing.ToMarkdown();
                Assert.Contains("# Excel Template Markers", markdown);
                Assert.Contains("| Template | A2 | Total | currency | yes | yes | number |", markdown);
                Assert.Contains("| Template | A3 | Customer.Name |  | no | no |  |", markdown);
                Assert.Contains("| Template | A4 | Logo |  | yes | yes | image |", markdown);
                Assert.Contains("| Template | A5 | Issued | date | no | yes | date/time |", markdown);

                ExcelTemplateInspection complete = sheet.InspectTemplate(new InvoiceTemplateModel {
                    Customer = new CustomerTemplateModel { Name = "Adatum" },
                    Total = 10,
                    Invoice = new InvoiceNumberTemplateModel { Number = "INV-001" },
                    Logo = ExcelTemplateImage.FromBytes(new byte[] { 4, 5, 6 }, widthPixels: 12, heightPixels: 12),
                    Issued = new DateTime(2026, 5, 28)
                });

                Assert.Same(complete, complete.EnsureAllMarkersBound());
                Assert.True(complete.AllMarkersBound);
                Assert.Empty(complete.MissingMarkers);
                Assert.Contains(complete.Markers, marker => marker.Name == "Logo" && marker.BoundValueKind == "image");
                Assert.Throws<InvalidOperationException>(() => template.EnsureAllMarkersBound());
            }
        }

        private static WorksheetPart GetWorksheetPartByNameForTemplateTests(SpreadsheetDocument document, string sheetName) {
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == sheetName);
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        }

        private static Xdr.TwoCellAnchor CreateDiagramAnchor(string layoutRelId, string colorsRelId, string styleRelId, string dataRelId) {
            return new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("5"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("8"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("13"),
                    new Xdr.RowOffset("0")),
                new Xdr.GraphicFrame(
                    new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties {
                            Id = 500U,
                            Name = "Template diagram"
                        },
                        new Xdr.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks {
                            NoChangeAspect = true
                        })),
                    new Xdr.Transform(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 3200000L, Cy = 1800000L }),
                    new A.Graphic(
                        new A.GraphicData(
                            new Dgm.RelationshipIds {
                                LayoutPart = layoutRelId,
                                ColorPart = colorsRelId,
                                StylePart = styleRelId,
                                DataPart = dataRelId
                            }) {
                                Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
                            })),
                new Xdr.ClientData());
        }

        private static void FeedPart(OpenXmlPart part, string xml) {
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(stream);
        }

        private static string ReadPartText(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private sealed class InvoiceTemplateModel {
            public CustomerTemplateModel Customer { get; set; } = new CustomerTemplateModel();
            public InvoiceNumberTemplateModel Invoice { get; set; } = new InvoiceNumberTemplateModel();
            public ExcelTemplateImage? Logo { get; set; }
            public decimal Total { get; set; }
            public DateTime Issued { get; set; }
            public double Completion { get; set; }
        }

        private sealed class InvoiceNumberTemplateModel {
            public string Number { get; set; } = string.Empty;
        }

        private sealed class CustomerTemplateModel {
            public string Name { get; set; } = string.Empty;
        }

        private sealed class RegionSheetTemplateModel {
            public string Region { get; set; } = string.Empty;
            public string Owner { get; set; } = string.Empty;
            public decimal Amount { get; set; }
        }
    }
}
