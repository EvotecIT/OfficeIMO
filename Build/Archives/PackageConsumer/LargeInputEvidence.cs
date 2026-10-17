using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;

namespace OfficeIMO.Archives.Qualification {

    /// <summary>Opt-in large-input allocation evidence; excluded from normal correctness CI.</summary>
    internal static class LargeInputEvidence {
        internal static int Run(string[] args) {
            if (args.Length != 2) throw new ArgumentException("Expected: <large ZIP path> <JSON report>");
            string input = args[0];
            if (!File.Exists(input)) CreateInput(input);
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            var results = new List<object>();
            Measure("Inspection", () => {
                using var inspection = ArchiveInspector.Open(input);
                using var member = inspection.OpenRead(inspection.Entries.Single(e => e.RawName == "member.txt").Ordinal);
                using var text = new StreamReader(member);
                if (text.ReadToEnd() != "complete") throw new InvalidDataException("Inspection payload mismatch.");
                return inspection.Entries.Count;
            });
            Measure("Traversal", () => {
                var traversal = ArchiveReader.TraverseZip(input);
                if (traversal.Entries.Count != 2) throw new InvalidDataException("Traversal entry count mismatch.");
                return traversal.Entries.Count;
            });
            Measure("Reader", () => {
                var chunks = reader.Read(input, new() { ComputeHashes = false }).ToArray();
                if (!chunks.Any(c => c.Text.Contains("complete"))) throw new InvalidDataException("Reader payload mismatch.");
                return chunks.Length;
            });
            var cleanup = new List<string>();
            foreach (bool cancel in new[] { false, true }) {
                VerifyFailureCleanup(input, cancel);
                cleanup.Add(cancel ? "Cancellation: source restored/open, private snapshot removed"
                    : "I/O failure: source restored/open, private snapshot removed");
            }
            using var hashInput = File.OpenRead(input);
            File.WriteAllText(args[1], JsonSerializer.Serialize(new {
                failureCleanup = cleanup,
                runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
                assembly = typeof(ArchiveInspector).Assembly.GetName().Version?.ToString(),
                inputBytes = new FileInfo(input).Length,
                inputSha256 = Convert.ToHexString(SHA256.HashData(hashInput)).ToLowerInvariant(),
                results,
                boundary = "Opt-in per-thread managed allocation, not peak/RSS or timing qualification"
            }, new JsonSerializerOptions { WriteIndented = true }));
            return 0;

            void Measure(string name, Func<int> action) {
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                long before = GC.GetAllocatedBytesForCurrentThread();
                int count = action();
                long bytes = GC.GetAllocatedBytesForCurrentThread() - before;
                results.Add(new { name, count, allocatedBytes = bytes });
            }
        }

        private static void CreateInput(string path) {
            using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
            using var zip = new ZipArchive(file, ZipArchiveMode.Create);
            using (var padding = zip.CreateEntry("padding.bin", CompressionLevel.NoCompression).Open()) {
                byte[] zeros = new byte[65536];
                for (int i = 0; i < 1024; i++) padding.Write(zeros);
            }
            using var text = new StreamWriter(zip.CreateEntry("member.txt", CompressionLevel.NoCompression).Open());
            text.Write("complete");
        }

        private static void VerifyFailureCleanup(string input, bool cancel) {
            var before = Directory.GetDirectories(Path.GetTempPath(), "officeimo-archive-*").ToHashSet(StringComparer.Ordinal);
            using var cancellation = new CancellationTokenSource();
            using var source = new FaultingFile(input, cancel ? cancellation : null);
            source.Position = 13;
            try {
                using var unexpected = ArchiveInspector.Open(source, cancellationToken: cancellation.Token);
                throw new InvalidDataException("Expected source failure.");
            } catch (OperationCanceledException) when (cancel) { }
            catch (IOException error) when (!cancel && error.Message == "Synthetic snapshot source failure") { }
            if (!source.CanRead || source.Position != 13) throw new InvalidDataException("Source ownership/position changed.");
            if (Directory.GetDirectories(Path.GetTempPath(), "officeimo-archive-*").Except(before).Any())
                throw new InvalidDataException("Snapshot directory leaked.");
        }

        private sealed class FaultingFile(string path, CancellationTokenSource? cancellation)
            : FileStream(path, FileMode.Open, FileAccess.Read) {
            public override int Read(byte[] buffer, int offset, int count) {
                int read = base.Read(buffer, offset, count);
                if (Position >= 65536) {
                    if (cancellation != null) cancellation.Cancel();
                    else throw new IOException("Synthetic snapshot source failure");
                }
                return read;
            }
        }
    }
}
