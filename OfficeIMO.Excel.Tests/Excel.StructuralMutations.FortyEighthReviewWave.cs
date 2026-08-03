using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralReferences_PreserveQuotedExternalThreeDimensionalQualifiers() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.CellFormula(
                1,
                1,
                "'[Other.xlsx]First:Data'!A1+'[Other.xlsx]First':'[Other.xlsx]Data'!A1+'Data'!A1");

            data.InsertRows(1);

            Assert.Equal(
                "'[Other.xlsx]First:Data'!A1+'[Other.xlsx]First':'[Other.xlsx]Data'!A1+'Data'!A2",
                Assert.Single(summary.GetFormulaCells()).Formula);
        }
    }
}
