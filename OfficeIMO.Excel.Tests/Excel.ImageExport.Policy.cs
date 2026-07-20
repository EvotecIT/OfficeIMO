using System.Reflection;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Fact]
        public void ExcelImageExportDiagnosticClassifier_CoversEveryPublishedCode() {
            string[] codes = typeof(ExcelImageExportDiagnosticCodes)
                .GetFields(BindingFlags.Public | BindingFlags.Static)
                .Where(field => field.IsLiteral && field.FieldType == typeof(string))
                .Select(field => Assert.IsType<string>(field.GetRawConstantValue()))
                .ToArray();

            Assert.NotEmpty(codes);
            Assert.All(codes, code =>
                Assert.True(Enum.IsDefined(typeof(OfficeImageExportLossKind), ExcelImageExportDiagnosticClassifier.Classify(code))));
        }

        [Fact]
        public void ExcelImageExportDiagnosticClassifier_TreatsCategoryAxisFormatFallbackAsApproximation() {
            Assert.Equal(
                OfficeImageExportLossKind.Approximation,
                ExcelImageExportDiagnosticClassifier.Classify(
                    ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported));
        }

        [Fact]
        public void ExcelRange_StrictOmissionPolicyRejectsUnrenderedCommentBody() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Strict");
            sheet.CellValue(1, 1, "Reviewed");
            sheet.SetComment("A1", "Needs design review", "Reviewer");

            OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
                sheet.Range("A1:B2").ExportImage(
                    OfficeImageExportFormat.Svg,
                    new ExcelImageExportOptions {
                        ShowGridlines = false,
                        Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                    }));

            Assert.Contains(
                exception.Diagnostics,
                diagnostic =>
                    diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported &&
                    diagnostic.LossKind == OfficeImageExportLossKind.Omission);
        }

        [Fact]
        public void ExcelRange_StrictOmissionPolicyAllowsRenderedCommentApproximation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Approximation");
            sheet.CellValue(1, 1, "Reviewed");
            sheet.SetComment("A1", "Needs design review", "Reviewer");

            OfficeImageExportResult result = sheet.Range("A1:F8").ExportImage(
                OfficeImageExportFormat.Svg,
                new ExcelImageExportOptions {
                    ShowGridlines = false,
                    ShowCommentBodies = true,
                    DefaultColumnWidthPixels = 92D,
                    DefaultRowHeightPixels = 28D,
                    Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
                });

            Assert.Contains(
                result.Diagnostics,
                diagnostic =>
                    diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation &&
                    diagnostic.LossKind == OfficeImageExportLossKind.Approximation);
        }
    }
}
