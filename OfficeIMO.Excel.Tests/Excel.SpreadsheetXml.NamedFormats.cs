using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelSpreadsheetXmlNamedFormatsTests {
        [Theory]
        [InlineData("Medium Date", "dd-mmm-yy")]
        [InlineData("Short Time", "h:mm")]
        [InlineData("Medium Time", "h:mm AM/PM")]
        [InlineData("Long Time", "h:mm:ss")]
        [InlineData("Fixed", "0.00")]
        [InlineData("Standard", "#,##0.00")]
        [InlineData("Percent", "0.00%")]
        [InlineData("Scientific", "0.00E+00")]
        [InlineData("Yes/No", "\"Yes\";\"Yes\";\"No\"")]
        [InlineData("True/False", "\"True\";\"True\";\"False\"")]
        [InlineData("On/Off", "\"On\";\"On\";\"Off\"")]
        [InlineData("General Number", "General")]
        public void Named_formats_translate_to_usable_OpenXml_codes(string name, string expected) {
            using ExcelDocument document = ExcelDocument.ImportSpreadsheetXml(Input(name)).RequireNoLoss();
            Assert.Equal(expected, document.Sheets[0].GetCellStyle(1, 1).NumberFormatCode);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Theory]
        [InlineData("Currency")]
        [InlineData("Euro Currency")]
        public void Locale_dependent_currency_formats_report_approximation(string name) {
            var result = ExcelDocument.ImportSpreadsheetXml(Input(name));
            using ExcelDocument document = result.Value;
            Assert.Contains(result.Report.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Approximation);
            Assert.Throws<OfficeConversionException>(() => result.RequireNoLoss());
            Assert.DoesNotContain("Currency", document.Sheets[0].GetCellStyle(1, 1).NumberFormatCode);
        }

        private static byte[] Input(string name) => Encoding.UTF8.GetBytes(
            "<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'>"
            + "<Styles><Style ss:ID='fmt'><NumberFormat ss:Format='" + name + "'/></Style></Styles>"
            + "<Worksheet ss:Name='Data'><Table><Row><Cell ss:StyleID='fmt'><Data ss:Type='Number'>12.5</Data></Cell></Row></Table></Worksheet></Workbook>");
    }
}
