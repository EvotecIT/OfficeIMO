using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class WordImageExportTests {
        [Fact]
        public void WordDocument_ExportsEveryEstimatedPageAsNamedImagesAndSnapshots() {
            using WordDocument document = CreateThreePageDocument();

            int estimatedPages = document.GetEstimatedPageCount();
            IReadOnlyList<WordDocumentVisualSnapshot> snapshots = document.CreateVisualSnapshots();
            IReadOnlyList<OfficeImageExportResult> images = document.ExportImages(OfficeImageExportFormat.Svg);

            Assert.Equal(3, estimatedPages);
            Assert.Equal(new[] { 0, 1, 2 }, snapshots.Select(snapshot => snapshot.PageIndex));
            Assert.Equal(new[] { "Page 1", "Page 2", "Page 3" }, images.Select(image => image.Name));
            Assert.Contains("First page", Encoding.UTF8.GetString(images[0].Bytes), StringComparison.Ordinal);
            Assert.Contains("Second page", Encoding.UTF8.GetString(images[1].Bytes), StringComparison.Ordinal);
            Assert.Contains("Third page", Encoding.UTF8.GetString(images[2].Bytes), StringComparison.Ordinal);
            Assert.Contains(snapshots[0].Fragments, fragment =>
                fragment.Text.Contains("First page", StringComparison.Ordinal) &&
                fragment.Region != null);
            Assert.Contains(snapshots[1].Fragments, fragment =>
                fragment.Text.Contains("Second page", StringComparison.Ordinal) &&
                fragment.Region != null);
            Assert.Contains(snapshots[2].Fragments, fragment =>
                fragment.Text.Contains("Third page", StringComparison.Ordinal) &&
                fragment.Region != null);
        }

        [Fact]
        public void WordDocument_BatchExportHonorsPageRange() {
            using WordDocument document = CreateThreePageDocument();
            var options = new WordImageExportOptions { PageIndex = 1, PageCount = 1 };

            IReadOnlyList<WordDocumentVisualSnapshot> snapshots = document.CreateVisualSnapshots(options);
            IReadOnlyList<OfficeImageExportResult> images = document.ExportImages(OfficeImageExportFormat.Png, options);

            WordDocumentVisualSnapshot snapshot = Assert.Single(snapshots);
            OfficeImageExportResult image = Assert.Single(images);
            Assert.Equal(1, snapshot.PageIndex);
            Assert.Equal("Page 2", image.Name);
            Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
            Assert.Equal(image.Width, raster!.Width);
        }

        [Fact]
        public void WordDocument_ImageExportProjectsNativeChartThroughSharedDrawingRenderer() {
            using WordDocument document = WordDocument.Create();
            WordChart chart = document.AddChart("Delivery status", false, 420, 240);
            chart.AddPie("Complete", 72);
            chart.AddPie("Remaining", 28);

            WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
            OfficeImageExportResult image = document.ExportImage(
                OfficeImageExportFormat.Png,
                new WordImageExportOptions {
                    Policy = new OfficeImageExportPolicy {
                        RequireNoOmissions = true,
                        RequireNoFailures = true
                    }
                });

            Assert.DoesNotContain(snapshot.Diagnostics,
                diagnostic => diagnostic.Code == WordImageExportDiagnosticCodes.UnsupportedChart);
            Assert.DoesNotContain(image.Diagnostics,
                diagnostic => diagnostic.Code == WordImageExportDiagnosticCodes.UnsupportedChart);
            Assert.True(snapshot.Drawing.Elements.Count > 0);
            Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
            Assert.NotNull(raster);
        }

        [Fact]
        public async Task WordDocument_BatchBuilderSavesSelectedPagesSynchronouslyAndAsynchronously() {
            string syncFolder = Path.Combine(Path.GetTempPath(), "OfficeIMO-Word-Pages-" + Guid.NewGuid().ToString("N"));
            string asyncFolder = Path.Combine(Path.GetTempPath(), "OfficeIMO-Word-Pages-Async-" + Guid.NewGuid().ToString("N"));
            try {
                using WordDocument document = CreateThreePageDocument();

                IReadOnlyList<OfficeImageExportResult> syncResults = document.ToImages()
                    .FromPage(1)
                    .TakePages(2)
                    .AsSvg()
                    .Save(syncFolder);
                IReadOnlyList<OfficeImageExportResult> asyncResults = await document.SaveAsImagesAsync(
                    asyncFolder,
                    OfficeImageExportFormat.Png,
                    new WordImageExportOptions { PageCount = 2 });

                Assert.Equal(2, syncResults.Count);
                Assert.Equal(2, asyncResults.Count);
                Assert.True(File.Exists(Path.Combine(syncFolder, "Page 2.svg")));
                Assert.True(File.Exists(Path.Combine(syncFolder, "Page 3.svg")));
                Assert.True(File.Exists(Path.Combine(asyncFolder, "Page 1.png")));
                Assert.True(File.Exists(Path.Combine(asyncFolder, "Page 2.png")));
            } finally {
                if (Directory.Exists(syncFolder)) Directory.Delete(syncFolder, recursive: true);
                if (Directory.Exists(asyncFolder)) Directory.Delete(asyncFolder, recursive: true);
            }
        }

        [Fact]
        public void WordDocument_BatchExportRejectsInvalidRanges() {
            using WordDocument document = CreateThreePageDocument();

            Assert.Throws<ArgumentOutOfRangeException>(() => document.ExportImages(
                OfficeImageExportFormat.Png,
                new WordImageExportOptions { PageIndex = 3 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => document.ExportImages(
                OfficeImageExportFormat.Png,
                new WordImageExportOptions { PageCount = 0 }));
        }

        [Fact]
        public void WordDocument_VisualSnapshotBatchHonorsCancellation() {
            using WordDocument document = CreateThreePageDocument();
            using var cancellation = new System.Threading.CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() =>
                document.CreateVisualSnapshots(
                    new WordImageExportOptions(),
                    cancellation.Token));
        }

        [Fact]
        public void WordDocument_VisualSnapshotBatchEnforcesMaximumOutputCount() {
            using WordDocument document = CreateThreePageDocument();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                document.CreateVisualSnapshots(new WordImageExportOptions {
                    MaximumOutputCount = 2
                }));

            Assert.Contains(nameof(OfficeImageExportOptions.MaximumOutputCount), exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task WordDocument_VisualSnapshotBatchCancelsDuringASingleLargeRun() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph(new string('x', 2000000));

            await AssertVisualSnapshotCancelsDuringWork(
                document,
                WordImageCancellationCheckpoint.PlainTextWrapping);
        }

        [Fact]
        public async Task WordDocument_VisualSnapshotBatchCancelsDuringNestedTableMeasurement() {
            using WordDocument document = WordDocument.Create();
            WordTable outer = document.AddTable(1, 1);
            WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1);
            nested.Rows[0].Cells[0].Paragraphs[0].Text = new string('y', 2000000);

            await AssertVisualSnapshotCancelsDuringWork(
                document,
                WordImageCancellationCheckpoint.NestedTableMeasurement);
        }

        private static async Task AssertVisualSnapshotCancelsDuringWork(
            WordDocument document,
            WordImageCancellationCheckpoint targetCheckpoint) {
            using var cancellation = new System.Threading.CancellationTokenSource();
            using var checkpointReached = new System.Threading.ManualResetEventSlim();
            using var releaseCheckpoint = new System.Threading.ManualResetEventSlim();
            var options = new WordImageExportOptions {
                CancellationCheckpoint = checkpoint => {
                    if (checkpoint != targetCheckpoint) {
                        return;
                    }
                    checkpointReached.Set();
                    if (!releaseCheckpoint.Wait(TimeSpan.FromSeconds(5))) {
                        throw new TimeoutException("Cancellation checkpoint was not released.");
                    }
                }
            };
            Task render = Task.Run(() => {
                document.CreateVisualSnapshots(
                    options,
                    cancellation.Token);
            });

            bool reached = checkpointReached.Wait(TimeSpan.FromSeconds(5));
            cancellation.Cancel();
            releaseCheckpoint.Set();

            Task completed = await Task.WhenAny(render, Task.Delay(TimeSpan.FromSeconds(5)));
            Assert.True(reached);
            Assert.Same(render, completed);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await render);
        }

        private static WordDocument CreateThreePageDocument() {
            WordDocument document = WordDocument.Create();
            document.AddParagraph("First page");
            document.AddPageBreak();
            document.AddParagraph("Second page");
            document.AddPageBreak();
            document.AddParagraph("Third page");
            return document;
        }
    }
}
