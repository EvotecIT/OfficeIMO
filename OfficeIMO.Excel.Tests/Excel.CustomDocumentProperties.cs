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
                document.AddSheet("Data").Cell(1, 1).SetValue("Ready");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.Equal("Approved", document.CustomProperties["ReleaseStatus"].Text);
                Assert.Equal(42, document.CustomProperties["Ticket"].NumberInteger);
                Assert.Equal(98.5D, document.CustomProperties["Score"].NumberDouble);
                Assert.True(document.CustomProperties["Reviewed"].Bool);
                Assert.Equal(reviewedAt, document.CustomProperties["ReviewedAt"].Date);

                document.SetCustomDocumentProperty("ReleaseStatus", "Published");
                Assert.True(document.RemoveCustomDocumentProperty("Ticket"));
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.Equal("Published", document.CustomProperties["ReleaseStatus"].Text);
                Assert.False(document.CustomProperties.ContainsKey("Ticket"));
            }
        }

        [Fact]
        public void Test_ExcelCustomDocumentProperties_PreserveNumericCompatibilityTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCustomDocumentProperties.NumericCompatibility.xlsx");
            long largeTicket = (long)int.MaxValue + 42L;
            ulong unsignedTicket = ulong.MaxValue;

            using (var document = ExcelDocument.Create(filePath)) {
                document.SetCustomDocumentProperty("LargeTicket", largeTicket);
                document.SetCustomDocumentProperty("UnsignedTicket", unsignedTicket);
                document.SetCustomDocumentProperty("Score", 12345.6789012345D);
                document.SetCustomDocumentProperty("Reviewed", true);
                document.AddSheet("Data").Cell(1, 1).SetValue("Ready");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                CustomFilePropertiesPart customPart = package.CustomFilePropertiesPart!;
                CustomDocumentProperty reviewed = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "Reviewed");
                reviewed.VTBool = new VTBool("1");
                customPart.Properties.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.Equal(largeTicket, document.CustomProperties["LargeTicket"].Value);
                Assert.Equal(unsignedTicket, document.CustomProperties["UnsignedTicket"].Value);
                Assert.True(document.CustomProperties["Reviewed"].Bool);
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                CustomFilePropertiesPart customPart = package.CustomFilePropertiesPart!;
                CustomDocumentProperty large = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "LargeTicket");
                CustomDocumentProperty unsigned = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "UnsignedTicket");
                CustomDocumentProperty score = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "Score");
                Assert.NotNull(large.VTInt64);
                Assert.Equal(largeTicket.ToString(System.Globalization.CultureInfo.InvariantCulture), large.VTInt64!.Text);
                Assert.NotNull(unsigned.VTUnsignedInt64);
                Assert.Equal(unsignedTicket.ToString(System.Globalization.CultureInfo.InvariantCulture), unsigned.VTUnsignedInt64!.Text);
                Assert.NotNull(score.VTDouble);
            }
        }

        [Fact]
        public void Test_ExcelCustomDocumentProperties_RoundTripBinaryValue() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCustomDocumentProperties.Binary.xlsx");
            byte[] payload = { 0x00, 0x01, 0x42, 0x80, 0xff };

            using (var document = ExcelDocument.Create(filePath)) {
                document.SetCustomDocumentProperty("BinaryPayload", payload);
                document.AddSheet("Data").Cell(1, 1).SetValue("Binary");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                CustomFilePropertiesPart customPart = package.CustomFilePropertiesPart!;
                CustomDocumentProperty binary = customPart.Properties!.Elements<CustomDocumentProperty>().First(property => property.Name == "BinaryPayload");
                Assert.NotNull(binary.VTBlob);
                Assert.Equal(Convert.ToBase64String(payload), binary.VTBlob!.Text);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelCustomProperty property = document.CustomProperties["BinaryPayload"];
                Assert.Equal(ExcelCustomPropertyType.Binary, property.PropertyType);
                Assert.Equal(payload, property.Binary);
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Equal("Reviewed", document.CustomProperties["Workflow"].Text);
            }
        }

        [Fact]
        public void Test_ExcelCustomDocumentProperties_DirectValueEditsAreTracked() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCustomDocumentProperties.DirectValueEdit.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddSheet("Data").Cell(1, 1).SetValue("Ready");
                document.SetCustomDocumentProperty("Workflow", "Draft");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                document.CustomProperties["Workflow"].Value = "Reviewed";
                document.CustomProperties["Workflow"].PropertyType = ExcelCustomPropertyType.Text;
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Equal("Reviewed", document.CustomProperties["Workflow"].Text);
            }
        }

        [Fact]
        public void Test_ExcelCustomDocumentProperties_DirectDictionaryAccessValidatesKeys() {
            using var document = ExcelDocument.Create(new MemoryStream());

            Assert.Throws<ArgumentException>(() => document.CustomProperties[" "] = new ExcelCustomProperty("Invalid"));
            Assert.Throws<ArgumentException>(() => document.CustomProperties.Add("\t", new ExcelCustomProperty("Invalid")));

            document.CustomProperties[" Workflow "] = new ExcelCustomProperty("Ready");

            Assert.True(document.CustomProperties.ContainsKey("Workflow"));
            Assert.True(document.CustomProperties.ContainsKey(" Workflow "));
            Assert.Equal("Ready", document.CustomProperties["Workflow"].Text);
        }
    }
}
