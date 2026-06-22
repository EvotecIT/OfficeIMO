using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelCustomDocumentProperties_RoundTripTypedValues() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCustomDocumentProperties.xlsx");
            DateTime reviewedAt = new DateTime(2026, 6, 22, 8, 30, 0, DateTimeKind.Utc);

            using (var document = ExcelDocument.Create(filePath)) {
                document.SetCustomDocumentProperty("ReleaseStatus", "Approved");
                document.SetCustomDocumentProperty("Ticket", 42);
                document.SetCustomDocumentProperty("Score", 98.5D);
                document.SetCustomDocumentProperty("Reviewed", true);
                document.SetCustomDocumentProperty("ReviewedAt", reviewedAt);
                document.AddWorkSheet("Data").CellValue(1, 1, "Ready");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.Equal("Approved", document.CustomDocumentProperties["ReleaseStatus"].Text);
                Assert.Equal(42, document.CustomDocumentProperties["Ticket"].NumberInteger);
                Assert.Equal(98.5D, document.CustomDocumentProperties["Score"].NumberDouble);
                Assert.True(document.CustomDocumentProperties["Reviewed"].Bool);
                Assert.Equal(reviewedAt, document.CustomDocumentProperties["ReviewedAt"].Date);

                document.SetCustomDocumentProperty("ReleaseStatus", "Published");
                Assert.True(document.RemoveCustomDocumentProperty("Ticket"));
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.Equal("Published", document.CustomDocumentProperties["ReleaseStatus"].Text);
                Assert.False(document.CustomDocumentProperties.ContainsKey("Ticket"));
            }
        }

        [Fact]
        public void Test_ExcelCustomDocumentProperties_PreserveNumericCompatibilityTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCustomDocumentProperties.NumericCompatibility.xlsx");
            long largeTicket = (long)int.MaxValue + 42L;

            using (var document = ExcelDocument.Create(filePath)) {
                document.SetCustomDocumentProperty("LargeTicket", largeTicket);
                document.SetCustomDocumentProperty("Score", 12345.6789012345D);
                document.SetCustomDocumentProperty("Reviewed", true);
                document.AddWorkSheet("Data").CellValue(1, 1, "Ready");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                CustomFilePropertiesPart customPart = package.CustomFilePropertiesPart!;
                CustomDocumentProperty reviewed = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "Reviewed");
                reviewed.VTBool = new VTBool("1");
                customPart.Properties.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: false)) {
                Assert.Equal(largeTicket, document.CustomDocumentProperties["LargeTicket"].Value);
                Assert.True(document.CustomDocumentProperties["Reviewed"].Bool);
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                CustomFilePropertiesPart customPart = package.CustomFilePropertiesPart!;
                CustomDocumentProperty large = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "LargeTicket");
                CustomDocumentProperty score = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "Score");
                Assert.NotNull(large.VTInt64);
                Assert.Equal(largeTicket.ToString(System.Globalization.CultureInfo.InvariantCulture), large.VTInt64!.Text);
                Assert.NotNull(score.VTDouble);
            }
        }

        [Fact]
        public void Test_ExcelCustomDocumentProperties_DisqualifyDirectDataSetFastSave() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCustomDocumentProperties.DirectDataSet.xlsx");
            var data = new DataTable("Data");
            data.Columns.Add("Name", typeof(string));
            data.Rows.Add("Alpha");

            using (var document = ExcelDocument.Create(filePath)) {
                var dataSet = new DataSet();
                dataSet.Tables.Add(data);
                document.InsertDataSet(dataSet);
                document.SetCustomDocumentProperty("Workflow", "Reviewed");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Equal("Reviewed", document.CustomDocumentProperties["Workflow"].Text);
            }
        }
    }
}
