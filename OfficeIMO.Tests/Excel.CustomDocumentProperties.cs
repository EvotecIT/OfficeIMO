using System;
using System.IO;
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
    }
}
