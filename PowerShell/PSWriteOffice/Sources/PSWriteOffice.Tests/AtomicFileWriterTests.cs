using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.IO;
using System.Text;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Tests;

public sealed class AtomicFileWriterTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WriteUnique_preserves_each_competing_writers_payload(bool includeCanonicalWriters)
    {
        var root = Path.Combine(Path.GetTempPath(), $"OfficeIMO-output-{Guid.NewGuid():N}");
        try
        {
            for (var round = 0; round < 2; round++)
            {
                var saved = await WriteCompetingOutputs(Path.Combine(root, round.ToString()),
                    unique: true, includeCanonicalWriters);
                Assert.Equal(includeCanonicalWriters ? 6 : 12,
                    saved.Count(result => !includeCanonicalWriters || result.Index % 2 == 0));
                if (includeCanonicalWriters)
                {
                    Assert.InRange(saved.Count(result => result.Index % 2 == 1), 0, 1);
                }
            }
        }
        finally
        {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Write_without_overwrite_preserves_the_only_successful_writers_payload(bool includeCanonicalWriters)
    {
        var root = Path.Combine(Path.GetTempPath(), $"OfficeIMO-output-{Guid.NewGuid():N}");
        try
        {
            for (var round = 0; round < 2; round++)
            {
                var saved = await WriteCompetingOutputs(Path.Combine(root, round.ToString()),
                    unique: false, includeCanonicalWriters);
                Assert.Single(saved);
            }
        }
        finally
        {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Output_respects_a_destination_claim_owned_by_a_canonical_writer(bool unique)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"OfficeIMO-output-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var destination = Path.Combine(directory, "report.txt");
            var uniquePath = Path.Combine(directory, "report-2.txt");
            var writeCount = 0;
            Action<string> serialize = temporaryPath =>
            {
                writeCount++;
                File.WriteAllText(temporaryPath, "new-payload");
            };

            using (OfficeFileCommitClaim.Acquire(destination))
            {
                if (unique)
                {
                    Assert.Equal(uniquePath, OfficeFileOutput.WriteUnique(directory, "report.txt", serialize));
                    Assert.Equal("new-payload", File.ReadAllText(uniquePath));
                }
                else
                {
                    Assert.Throws<IOException>(() => OfficeFileOutput.Write(destination, overwrite: false, serialize));
                }
                Assert.False(File.Exists(destination));
                Assert.Equal(1, writeCount);
            }

            Assert.Equal(unique ? new[] { uniquePath } : Array.Empty<string>(), Directory.GetFileSystemEntries(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteUnique_preserves_an_unrelated_file_with_a_claim_shaped_name(bool markerShapedContent)
    {
        if (!System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows)) return;
        var directory = Path.Combine(Path.GetTempPath(), $"OfficeIMO-foreign-claim-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var destination = Path.Combine(directory, "report.txt");
            var foreignPath = OfficeFileCommitClaim.CreatePath(destination);
            byte[] foreignBytes = markerShapedContent
                ? Encoding.ASCII.GetBytes("OfficeIMO-commit-claim-v1:").Concat(Guid.NewGuid().ToByteArray()).ToArray()
                : Encoding.UTF8.GetBytes("keep this unrelated file");
            File.WriteAllBytes(foreignPath, foreignBytes);

            var output = OfficeFileOutput.WriteUnique(directory, "report.txt", Encoding.UTF8.GetBytes("payload"));

            Assert.Equal(Path.Combine(directory, "report-2.txt"), output);
            Assert.Equal("payload", File.ReadAllText(output));
            Assert.Equal(foreignBytes, File.ReadAllBytes(foreignPath));
            Assert.False(File.Exists(destination));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void FailedStagingPreservesExistingOutputAndRemovesPartialFile() {
        var directory = Path.Combine(Path.GetTempPath(), $"OfficeIMO-output-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try {
            var output = Path.Combine(directory, "report.txt");
            File.WriteAllText(output, "existing report");
            Assert.Throws<InvalidOperationException>(() =>
                OfficeFileOutput.Write(output, overwrite: true, temporaryPath => {
                    File.WriteAllText(temporaryPath, "incomplete report");
                    throw new InvalidOperationException("serialization failed");
                }));
            Assert.Equal("existing report", File.ReadAllText(output));
            Assert.Equal(new[] { output }, Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void WriteUnique_writes_payload_once_when_names_collide()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            File.WriteAllText(Path.Combine(directory, "report.txt"), "existing-1");
            File.WriteAllText(Path.Combine(directory, "report-2.txt"), "existing-2");
            var writeCount = 0;

            var outputPath = OfficeFileOutput.WriteUnique(directory, "report.txt", temporaryPath =>
            {
                writeCount++;
                File.WriteAllText(temporaryPath, "new-payload", new UTF8Encoding(false));
            });

            Assert.Equal(1, writeCount);
            Assert.Equal("report-3.txt", Path.GetFileName(outputPath));
            Assert.Equal("new-payload", File.ReadAllText(outputPath));
            Assert.Equal("existing-1", File.ReadAllText(Path.Combine(directory, "report.txt")));
            Assert.Equal("existing-2", File.ReadAllText(Path.Combine(directory, "report-2.txt")));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_cleans_temporary_file_when_writer_fails()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                OfficeFileOutput.WriteUnique(directory, "report.txt", temporaryPath =>
                {
                    File.WriteAllText(temporaryPath, "partial");
                    throw new InvalidOperationException("writer failed");
                }));

            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_keeps_temporary_name_within_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 220) + ".bin";

            var outputPath = OfficeFileOutput.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));

            Assert.Equal(fileName, Path.GetFileName(outputPath));
            Assert.Equal("payload", File.ReadAllText(outputPath));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_truncates_initial_name_to_portable_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 260) + ".bin";

            var outputPath = OfficeFileOutput.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.Equal(255, outputName.Length);
            Assert.Equal(new string('a', 251) + ".bin", outputName);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_truncates_initial_unicode_name_to_portable_byte_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('\u00e9', 130) + ".bin";

            var outputPath = OfficeFileOutput.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.True(Encoding.UTF8.GetByteCount(outputName) <= 255);
            Assert.EndsWith(".bin", outputName, StringComparison.Ordinal);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void GetUniquePath_bounds_initial_and_collision_names()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 260) + ".pdf";
            var initialPath = PdfCommandUtilities.GetUniquePath(directory, fileName);
            var initialName = Path.GetFileName(initialPath);
            Assert.Equal(255, initialName.Length);
            Assert.EndsWith(".pdf", initialName, StringComparison.Ordinal);

            File.WriteAllText(initialPath, "existing");
            var collisionPath = PdfCommandUtilities.GetUniquePath(directory, fileName);
            var collisionName = Path.GetFileName(collisionPath);
            Assert.Equal(255, collisionName.Length);
            Assert.EndsWith("-2.pdf", collisionName, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_keeps_collision_name_within_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 251) + ".bin";
            File.WriteAllText(Path.Combine(directory, fileName), "existing");

            var outputPath = OfficeFileOutput.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.Equal(255, outputName.Length);
            Assert.Equal(new string('a', 249) + "-2.bin", outputName);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Equal("existing", File.ReadAllText(Path.Combine(directory, fileName)));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_keeps_unicode_collision_name_within_portable_byte_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('\u00e9', 125) + ".bin";
            File.WriteAllText(Path.Combine(directory, fileName), "existing");

            var outputPath = OfficeFileOutput.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.True(Encoding.UTF8.GetByteCount(outputName) <= 255);
            Assert.EndsWith("-2.bin", outputName, StringComparison.Ordinal);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Equal("existing", File.ReadAllText(Path.Combine(directory, fileName)));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Write_keeps_temporary_name_within_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 220) + ".pdf";
            var outputPath = Path.Combine(directory, fileName);

            OfficeFileOutput.Write(outputPath, Encoding.UTF8.GetBytes("payload"), overwrite: false);

            Assert.Equal("payload", File.ReadAllText(outputPath));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static async Task<(int Index, string Path)[]> WriteCompetingOutputs(
        string directory, bool unique, bool includeCanonicalWriters)
    {
        const int writers = 12;
        Directory.CreateDirectory(directory);
        var destination = Path.Combine(directory, "report.bin");
        var payloads = Enumerable.Range(0, writers).Select(index => new byte[] { (byte)index, 1, 2, 3 }).ToArray();
        var writeCounts = new int[writers];
        using var ready = new CountdownEvent(writers);
        using var start = new ManualResetEventSlim(false);
        Task<(int Index, string? Path)>[] attempts = Enumerable.Range(0, writers).Select(index =>
            Task.Factory.StartNew(() =>
            {
                var canonical = includeCanonicalWriters && index % 2 == 1;
                void WaitForPublication()
                {
                    Interlocked.Increment(ref writeCounts[index]);
                    ready.Signal();
                    if (!start.Wait(TimeSpan.FromSeconds(20)))
                        throw new TimeoutException("The output publication barrier was not released.");
                }

                try
                {
                    if (canonical)
                    {
                        OfficeFileCommit.Write(destination, stream =>
                        {
                            stream.Write(payloads[index], 0, payloads[index].Length);
                            WaitForPublication();
                        }, OfficeFileCommit.ConflictPolicy.FailIfExists);
                        return (Index: index, Path: (string?)destination);
                    }

                    Action<string> serialize = temporaryPath =>
                    {
                        File.WriteAllBytes(temporaryPath, payloads[index]);
                        WaitForPublication();
                    };
                    if (unique)
                        return (Index: index, Path: (string?)OfficeFileOutput.WriteUnique(directory, "report.bin", serialize));
                    OfficeFileOutput.Write(destination, overwrite: false, serialize);
                    return (Index: index, Path: (string?)destination);
                }
                catch (IOException) when (canonical || !unique)
                {
                    return (Index: index, Path: (string?)null);
                }
            }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)).ToArray();

        bool allWritersReady;
        try
        {
            allWritersReady = ready.Wait(TimeSpan.FromSeconds(15));
        }
        finally
        {
            start.Set();
        }
        var results = await Task.WhenAll(attempts);
        Assert.True(allWritersReady, "All serializers must reach the publication barrier.");
        Assert.All(writeCounts, count => Assert.Equal(1, count));
        var saved = results.Where(result => result.Path != null)
            .Select(result => (result.Index, Path: result.Path!)).ToArray();
        Assert.Equal(saved.Length, saved.Select(result => result.Path).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        foreach (var result in saved)
            Assert.Equal(payloads[result.Index], File.ReadAllBytes(result.Path));
        Assert.Equal(saved.Select(result => result.Path).OrderBy(path => path, StringComparer.OrdinalIgnoreCase),
            Directory.GetFileSystemEntries(directory).OrderBy(path => path, StringComparer.OrdinalIgnoreCase));
        return saved;
    }
}
