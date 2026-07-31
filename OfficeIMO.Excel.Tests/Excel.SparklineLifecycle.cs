using System.Linq;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_Sparklines_SupportReadUpdateDeleteLifecycle() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AddSparklines("A1:C2", "D1:D2");

            Assert.Equal(new[] { "D1", "D2" }, sheet.GetSparklines().Select(item => item.LocationRange));
            Assert.Equal(2, sheet.SetSparklineType("D1:D2", SparklineTypeValues.Column));
            Assert.All(sheet.GetSparklines(), item => Assert.Equal(SparklineTypeValues.Column, item.Type));
            Assert.Equal(1, sheet.RemoveSparklines("D1"));
            Assert.Equal("D2", Assert.Single(sheet.GetSparklines()).LocationRange);
            Assert.Equal(1, sheet.ClearSparklines());
            Assert.Empty(sheet.GetSparklines());
        }
    }
}
