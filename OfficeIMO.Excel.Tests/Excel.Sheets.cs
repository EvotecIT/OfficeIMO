using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests related to worksheet collection behavior.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_SheetsGetter_DoesNotIncreaseIdCount() {
            string filePath = Path.Combine(_directoryDocuments, "BasicExcel.xlsx");
            using var document = ExcelDocument.Load(filePath);

            // First access to Sheets initializes the ID list
            var sheets1 = document.Sheets;
            int idCountAfterFirst = document.id.Count;

            // Subsequent accesses should not change the ID list count
            var sheets2 = document.Sheets;
            int idCountAfterSecond = document.id.Count;
            var sheets3 = document.Sheets;
            int idCountAfterThird = document.id.Count;

            Assert.Equal(idCountAfterFirst, idCountAfterSecond);
            Assert.Equal(idCountAfterSecond, idCountAfterThird);

            // Ensure the ID list contains a unique entry for each sheet plus the initial 0
            Assert.Equal(sheets1.Count + 1, idCountAfterFirst);
            Assert.Equal(document.id.Count, document.id.Distinct().Count());
        }
    }
}
