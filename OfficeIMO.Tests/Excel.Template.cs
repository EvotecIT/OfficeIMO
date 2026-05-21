using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
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
        public void Test_ExcelTemplate_CanThrowOnMissingMarker() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.MissingMarker.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("Hello {{Name}}");

                Assert.Throws<InvalidOperationException>(() =>
                    sheet.ApplyTemplate(new Dictionary<string, object?>(), throwOnMissing: true));
            }
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
        public void Test_ExcelTemplate_InspectionReportsMarkersAndMissingBindings() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelTemplate.Inspect.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Template");
                sheet.CellAt(1, 1).SetValue("Invoice {{Invoice.Number}}");
                sheet.CellAt(2, 1).SetValue("{{Total:currency}}");
                sheet.CellAt(3, 1).SetValue("Customer {{Customer.Name}}");

                ExcelTemplateInspection template = document.InspectTemplate();
                Assert.Equal(3, template.TotalMarkers);
                Assert.False(template.HasBindingInfo);
                Assert.Contains("Invoice.Number", template.UniqueMarkers);
                Assert.Contains(template.Markers, marker => marker.Name == "Total"
                    && marker.Format == "currency"
                    && marker.IsWholeCell
                    && marker.CellReference == "A2");

                ExcelTemplateInspection missing = document.InspectTemplate(new Dictionary<string, object?> {
                    ["Invoice.Number"] = "INV-001",
                    ["Total"] = 10
                });

                Assert.True(missing.HasBindingInfo);
                Assert.False(missing.AllMarkersBound);
                Assert.Single(missing.MissingMarkerNames);
                Assert.Equal("Customer.Name", missing.MissingMarkerNames[0]);
                InvalidOperationException missingException = Assert.Throws<InvalidOperationException>(() => missing.EnsureAllMarkersBound());
                Assert.Contains("Customer.Name", missingException.Message);

                string markdown = missing.ToMarkdown();
                Assert.Contains("# Excel Template Markers", markdown);
                Assert.Contains("| Template | A2 | Total | currency | yes | yes |", markdown);
                Assert.Contains("| Template | A3 | Customer.Name |  | no | no |", markdown);

                ExcelTemplateInspection complete = sheet.InspectTemplate(new InvoiceTemplateModel {
                    Customer = new CustomerTemplateModel { Name = "Adatum" },
                    Total = 10,
                    Invoice = new InvoiceNumberTemplateModel { Number = "INV-001" }
                });

                Assert.Same(complete, complete.EnsureAllMarkersBound());
                Assert.True(complete.AllMarkersBound);
                Assert.Empty(complete.MissingMarkers);
                Assert.Throws<InvalidOperationException>(() => template.EnsureAllMarkersBound());
            }
        }

        private sealed class InvoiceTemplateModel {
            public CustomerTemplateModel Customer { get; set; } = new CustomerTemplateModel();
            public InvoiceNumberTemplateModel Invoice { get; set; } = new InvoiceNumberTemplateModel();
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
    }
}
