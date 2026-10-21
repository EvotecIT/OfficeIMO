using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("failure", false)]
        [InlineData("file", false)]
        [InlineData("file", true)]
        [InlineData("stream", false)]
        [InlineData("template", false)]
        public void Save_retained_drawings_and_aliases_preserve_later_edits(string route, bool standardWriter) {
            string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-RetainedDrawings-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try {
                using var document = ExcelDocument.Create();
                document.SheetCachingEnabled = false;
                var (chart, image) = CreateRetainedDrawingWorkbook(document);
                ExcelChart alias = document["Report"].Charts.Single();
                ExcelImage imageAlias = document["Report"].Images.Single();
                document["Report"].Name = "Renamed";
                GC.Collect();
                GC.WaitForPendingFinalizers();
                var options = new ExcelSaveOptions { DisableFastPackageWriter = standardWriter };
                if (route == "failure") {
                    string blocked = Path.Combine(directory, "blocked.xlsx");
                    Directory.CreateDirectory(blocked);
                    Exception? error = Record.Exception(() => document.Save(blocked));
                    Assert.True(error is IOException || error is UnauthorizedAccessException, error?.ToString());
                } else if (route == "stream") {
                    using var first = new MemoryStream();
                    document.Save(first, options);
                } else {
                    document.Save(Path.Combine(directory, route == "template" ? "first.xltx" : "first.xlsx"), options);
                }

                chart.Name = "Retained chart";
                alias.SetTitle("After save");
                image.Name = "Retained image";
                imageAlias.Title = "After save";
                Assert.Equal("After save", chart.Title);
                Assert.Equal("After save", image.Title);
                Assert.True(chart.TryGetData(out _));
                Assert.Equal("Renamed", chart.DataRange?.SheetName);
                string final = Path.Combine(directory, "final.xlsx");
                document.Save(final, options);
                using var reopened = ExcelDocument.Load(final);
                Assert.Equal("Retained chart", reopened["Renamed"].Charts.Single().Name);
                Assert.Equal("After save", reopened["Renamed"].Charts.Single().Title);
                Assert.Equal("Retained image", reopened["Renamed"].Images.Single().Name);
                Assert.Equal("After save", reopened["Renamed"].Images.Single().Title);
            } finally {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void Save_clean_workbook_preserves_a_chart_name_only_edit() {
            using var document = ExcelDocument.Create();
            CreateRetainedDrawingWorkbook(document);
            using var seed = new MemoryStream();
            document.Save(seed);
            using var loaded = ExcelDocument.Load(seed);
            loaded["Report"].Charts.Single().Name = "Name only";
            using var result = new MemoryStream();
            loaded.Save(result);
            using var reopened = ExcelDocument.Load(result);
            Assert.Equal("Name only", reopened["Report"].Charts.Single().Name);
        }

        [Fact]
        public void Mutation_rollback_preserves_retained_drawing_handles() {
            using var document = ExcelDocument.Create();
            var (chart, image) = CreateRetainedDrawingWorkbook(document);
            ExcelSheet sheet = document["Report"];
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                chart.SetTitle("Interrupted title");
                image.Title = "Interrupted image";
                throw new InvalidOperationException("Abort transaction");
            }, 0, new ExcelMutationPlanOptions(), default));
            Assert.Equal("Before title", chart.Title);
            Assert.Equal("Before image", image.Title);
            chart.SetTitle("After rollback");
            image.Title = "After rollback";
            using var result = new MemoryStream();
            document.Save(result);
            using var reopened = ExcelDocument.Load(result);
            Assert.Equal("After rollback", reopened["Report"].Charts.Single().Title);
            Assert.Equal("After rollback", reopened["Report"].Images.Single().Title);
        }

        [Fact]
        public void Collected_picture_lookups_preserve_live_drawing_aliases_after_save() {
            using var document = ExcelDocument.Create();
            var (chart, image) = CreateRetainedDrawingWorkbook(document);
            ExcelSheet sheet = document["Report"];
            string imageName = image.Name;
            WeakReference<ExcelImage> discarded = CreateDiscardedPictureLookups(sheet, imageName);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Assert.False(discarded.TryGetTarget(out _));
            CreateDiscardedPictureLookups(sheet, imageName);

            string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-DrawingAliases-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try {
                string path = Path.Combine(directory, "report.xlsx");
                document.Save(path);
                chart.SetTitle("Retained after lookups");
                image.Title = "Retained after lookups";
                document.Save(path);
                using var reopened = ExcelDocument.Load(path);
                Assert.Equal("Retained after lookups", reopened["Report"].Charts.Single().Title);
                Assert.Equal("Retained after lookups", reopened["Report"].Images.Single().Title);
            } finally {
                Directory.Delete(directory, recursive: true);
            }
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static WeakReference<ExcelImage> CreateDiscardedPictureLookups(ExcelSheet sheet, string name) {
            ExcelImage? last = null;
            for (int index = 0; index < 1024; index++) last = sheet.GetImage(name);
            return new WeakReference<ExcelImage>(last!);
        }

        [Fact]
        public void Removed_worksheet_drawings_do_not_attach_to_a_replacement() {
            using var document = ExcelDocument.Create();
            var (chart, image) = CreateRetainedDrawingWorkbook(document);
            document.AddWorksheet("Keep");
            document.RemoveWorksheet("Report");
            CreateRetainedDrawingWorkbook(document);
            using var result = new MemoryStream();
            document.Save(result);
            Assert.Throws<InvalidOperationException>(() => chart.Name = "Stale chart");
            Assert.Throws<InvalidOperationException>(() => image.Title = "Stale image");
            using var reopened = ExcelDocument.Load(result);
            Assert.Equal("Before title", reopened["Report"].Charts.Single().Title);
            Assert.Equal("Before image", reopened["Report"].Images.Single().Title);
        }

        private static (ExcelChart Chart, ExcelImage Image) CreateRetainedDrawingWorkbook(ExcelDocument document) {
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, 42d);
            var range = new ExcelChartDataRange("Report", 1, 1, 1, 1, hasHeaderRow: false);
            ExcelChart chart = sheet.AddChart(range, 4, 1, 300, 180, title: "Before title");
            byte[] imageBytes = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "Assets", "OfficeIMO.png"));
            ExcelImage image = sheet.AddImage(12, 1, imageBytes,
                "image/png", widthPixels: 20, heightPixels: 20);
            image.Title = "Before image";
            return (chart, image);
        }
    }
}
