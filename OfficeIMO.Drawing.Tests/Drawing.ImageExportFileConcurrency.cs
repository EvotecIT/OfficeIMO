using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData("sync")]
    [InlineData("async")]
    [InlineData("mixed")]
    public async Task CreateUniquePreservesEveryPayloadAcrossConcurrentWriters(string saveMode) {
        const int writerCount = 12;
        const int rounds = 4;
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        OfficeImageExportResult[] images = Enumerable.Range(0, writerCount)
            .Select(index => new OfficeImageExportResult(
                OfficeImageExportFormat.Png,
                1,
                1,
                OfficePngWriter.Encode(new OfficeRasterImage(
                    1,
                    1,
                    OfficeColor.FromRgb((byte)(index * 20), 40, 80)))))
            .ToArray();

        try {
            // Repeated, simultaneous starts exercise the claim's release/reacquisition window.
            // Different bytes identify an overwritten result even if every reported file exists.
            for (int round = 0; round < rounds; round++) {
                string folder = Path.Combine(root, round.ToString(System.Globalization.CultureInfo.InvariantCulture));
                Directory.CreateDirectory(folder);
                string path = Path.Combine(folder, "preview.png");
                using var ready = new CountdownEvent(writerCount);
                using var start = new ManualResetEventSlim(false);

                Task<OfficeImageExportResult>[] saves = images.Select((image, index) =>
                    Task.Factory.StartNew(async () => {
                        ready.Signal();
                        start.Wait();
                        bool asynchronous = saveMode == "async" || (saveMode == "mixed" && index % 2 != 0);
                        return asynchronous
                            ? await image.SaveAsync(path, OfficeImageExportFileConflictPolicy.CreateUnique).ConfigureAwait(false)
                            : image.Save(path, OfficeImageExportFileConflictPolicy.CreateUnique);
                    }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default).Unwrap())
                    .ToArray();

                // The barrier uses dedicated workers, not the pool needed by async file I/O.
                bool allWritersReady = ready.Wait(TimeSpan.FromSeconds(15));
                start.Set();
                OfficeImageExportResult[] saved = await Task.WhenAll(saves);
                Assert.True(allWritersReady, "All concurrent image writers must reach the start barrier.");

                string[] savedPaths = saved.Select(item => item.SavedPath!).ToArray();
                Assert.Equal(writerCount, savedPaths.Distinct(StringComparer.OrdinalIgnoreCase).Count());
                Assert.Contains(Path.GetFullPath(path), savedPaths);
                for (int index = 0; index < writerCount; index++) {
                    Assert.Equal(images[index].Bytes, File.ReadAllBytes(savedPaths[index]));
                }

                Assert.Equal(
                    savedPaths.OrderBy(item => item, StringComparer.Ordinal),
                    Directory.GetFileSystemEntries(folder).OrderBy(item => item, StringComparer.Ordinal));
                Directory.Delete(folder, recursive: true);
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }
}
