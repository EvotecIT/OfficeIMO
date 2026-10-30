using System.Data;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("simple", false)]
        [InlineData("extended", false)]
        [InlineData("dataset", false)]
        [InlineData("extended", true)]
        public async Task Save_file_commit_failure_is_reported_and_retained_edits_remain_usable(string writer, bool asynchronous) {
            string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-SaveFailure-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try {
                // Confirm that the same public input selects each file-save owner before testing its failure path.
                using (var control = CreateFileFailureWorkbook(writer, out _, out _)) {
                    control.Save(Path.Combine(directory, "control.xlsx"));
                    ExcelSavePackageWriter expected = writer == "dataset"
                        ? ExcelSavePackageWriter.DirectDataSetPackage
                        : writer == "extended" ? ExcelSavePackageWriter.ExtendedPackage : ExcelSavePackageWriter.SimplePackage;
                    Assert.Equal(expected, control.LastSaveDiagnostics.Writer);
                }

                using var document = CreateFileFailureWorkbook(writer, out ExcelSheet? retainedSheet, out ExcelCustomProperty? retainedProperty);
                ExcelBuiltinDocumentProperties retainedBuiltinProperties = document.BuiltinDocumentProperties;
                string blockedPath = Path.Combine(directory, "blocked.xlsx");
                Directory.CreateDirectory(blockedPath);

                // Coordinate a real filesystem obstacle disappearing only if saving retries another writer.
                // The assertion concerns the reported failure and resulting files, not timing callback counts.
                bool initialPreparationFinished = writer == "dataset";
                document.Execution.OnTiming = (operation, _) => {
                    if (initialPreparationFinished && operation == "Save.PrepareWorkbook.MaterializeDeferredDataSet"
                        && Directory.Exists(blockedPath)) {
                        Directory.Delete(blockedPath);
                    }
                    if (operation == "Save.PrepareWorkbook") initialPreparationFinished = true;
                };

                Exception? failure = asynchronous
                    ? await Record.ExceptionAsync(() => document.SaveAsync(blockedPath))
                    : Record.Exception(() => document.Save(blockedPath));
                document.Execution.OnTiming = null;

                Assert.NotNull(failure);
                Assert.True(failure is IOException || failure is UnauthorizedAccessException, failure.ToString());
                Assert.True(Directory.Exists(blockedPath));
                Assert.False(File.Exists(blockedPath));
                Assert.Null(document.FilePath);

                ExcelSheet sheet = retainedSheet ?? document.Sheets.Single();
                sheet.CellValue(1, 1, "After file failure");
                retainedBuiltinProperties.Title = "Still editable";
                if (retainedProperty != null) {
                    retainedProperty.Value = "After file failure";
                    Assert.Same(retainedProperty, document.CustomDocumentProperties["Review"]);
                } else {
                    document.SetCustomDocumentProperty("Review", "After file failure");
                }

                string recoveredPath = Path.Combine(directory, "recovered.xlsx");
                document.Save(recoveredPath);
                using var reopened = ExcelDocument.Load(recoveredPath);
                Assert.Equal("After file failure", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
                Assert.Equal("Still editable", reopened.BuiltinDocumentProperties.Title);
                Assert.Equal("After file failure", reopened.CustomDocumentProperties["Review"].Value);
            } finally {
                Directory.Delete(directory, recursive: true);
            }
        }

        private static ExcelDocument CreateFileFailureWorkbook(
            string writer, out ExcelSheet? retainedSheet, out ExcelCustomProperty? retainedProperty) {
            var document = ExcelDocument.Create();
            retainedSheet = null;
            retainedProperty = null;
            if (writer == "dataset") {
                var data = new DataSet();
                DataTable table = data.Tables.Add("Report");
                table.Columns.Add("Name", typeof(string));
                table.Rows.Add("Before file failure");
                document.InsertDataSet(data, createTables: false, includeHeaders: false,
                    includeAutoFilter: false, autoFit: false);
                return document;
            }

            retainedSheet = document.AddSheet("Report");
            retainedSheet.CellValue(1, 1, "Before file failure");
            retainedProperty = new ExcelCustomProperty("Before file failure");
            document.SetCustomDocumentProperty("Review", retainedProperty);
            if (writer == "extended") {
                var chart = new ExcelChartData(new[] { "Before file failure" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) });
                retainedSheet.AddChart(OfficeChartType.ColumnClustered, chart, new OfficeIMO.Excel.ExcelChartOptions { Row = 4, Column = 1, Width = OfficeLength.Pixels(300), Height = OfficeLength.Pixels(180), Title = "Retained chart" });
            }
            return document;
        }
    }
}
