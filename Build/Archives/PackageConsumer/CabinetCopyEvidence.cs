using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Archives.Qualification {
    /// <summary>Opt-in equivalent-work allocation and private-file cleanup evidence; run with isolated TMPDIR/TEMP.</summary>
    internal static class CabinetCopyEvidence {
        // A harness may set a child-process file-size limit to inject a real staging write failure.
        internal static int RunStagingFailure(string[] args) {
            if (args.Length != 1) throw new ArgumentException("Expected: <existing CAB>");
            var before = Directory.GetDirectories(Path.GetTempPath(), "officeimo-archive-*").ToHashSet(StringComparer.Ordinal);
            using var source = File.OpenRead(args[0]); source.Position = 13;
            using var inspection = ArchiveInspector.Open(source);
            using var cancellation = new CancellationTokenSource();
            using var destination = new DigestSink("CopyTo", cancellation, before);
            try {
                inspection.CopyTo(inspection.Entries.OrderByDescending(e => e.Length).First().Ordinal, destination);
                throw new InvalidDataException("Expected an injected staging write failure.");
            } catch (Exception error) when (error is IOException || error is ArgumentOutOfRangeException range && range.ParamName == "value") {
                if (destination.Bytes != 0 || !source.CanRead || source.Position != 13 || !destination.CanWrite)
                    throw new InvalidDataException("Staging failure changed destination or caller stream ownership.");
                if (Directory.GetDirectories(Path.GetTempPath(), "officeimo-archive-*").Except(before).Any())
                    throw new InvalidDataException("Failed staging file leaked.");
                Console.WriteLine(JsonSerializer.Serialize(new { stage = "Private CAB spool write", failureType = error.GetType().Name, error.HResult,
                    destinationBytes = destination.Bytes, cleanedUp = true,
                    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription }));
            }
            return 0;
        }

        internal static int Run(string[] args) {
            if (args.Length != 2) throw new ArgumentException("Expected: <task-owned CAB> <JSON report>");
            if (!File.Exists(args[0])) {
                CabinetReadEvidence.CreateInput(args[0]);
                using var patch = new FileStream(args[0], FileMode.Open, FileAccess.Write);
                // Select the entire reproducible 32MiB folder as the first member, retaining the second alias range.
                patch.Position = 44;
                using var writer = new BinaryWriter(patch);
                writer.Write(32 * 1024 * 1024); writer.Write(0);
            }
            var results = new List<object>();
            string? expected = null;
            foreach (string operation in new[] { "OpenRead", "CopyTo", "DestinationFailure", "DestinationCancellation" }) {
                var before = Directory.GetDirectories(Path.GetTempPath(), "officeimo-archive-*").ToHashSet(StringComparer.Ordinal);
                using var source = File.OpenRead(args[0]); source.Position = 13;
                using var inspection = ArchiveInspector.Open(source);
                int ordinal = inspection.Entries.OrderByDescending(e => e.Length).First().Ordinal;
                using var cancellation = new CancellationTokenSource();
                using var sink = new DigestSink(operation, cancellation, before);
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                long start = GC.GetAllocatedBytesForCurrentThread();
                try {
                    if (operation == "OpenRead") {
                        using var member = inspection.OpenRead(ordinal);
                        member.CopyTo(sink);
                    } else inspection.CopyTo(ordinal, sink, cancellation.Token);
                    if (operation.StartsWith("Destination", StringComparison.Ordinal)) throw new InvalidDataException("Expected interruption.");
                } catch (IOException error) when (operation == "DestinationFailure" && error.Message == "Synthetic destination failure") { }
                catch (OperationCanceledException) when (operation == "DestinationCancellation") { }
                long allocated = GC.GetAllocatedBytesForCurrentThread() - start;
                string digest = sink.Digest();
                if (operation == "OpenRead") expected = digest;
                if (operation == "CopyTo" && (digest != expected || sink.Bytes != inspection.Entries[ordinal].Length))
                    throw new InvalidDataException("Copy and complete read differ.");
                if (!source.CanRead || source.Position != 13 || !sink.CanWrite) throw new InvalidDataException("Caller stream ownership/position changed.");
                if (Directory.GetDirectories(Path.GetTempPath(), "officeimo-archive-*").Except(before).Any())
                    throw new InvalidDataException("Private copy spool leaked.");
                results.Add(new { operation, allocatedBytes = allocated, sha256 = digest, selectedBytes = sink.Bytes,
                    decodedBytes = inspection.ReadBytes, workBytes = inspection.DecompressionWorkBytes,
                    liveSpoolBytes = sink.LiveSpoolBytes, cleanedUp = true });
            }
            using var input = File.OpenRead(args[0]);
            File.WriteAllText(args[1], JsonSerializer.Serialize(new {
                runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
                assembly = typeof(ArchiveInspector).Assembly.GetName().Version?.ToString(),
                inputBytes = input.Length, inputSha256 = Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant(), results,
                boundary = "Per-thread managed allocation excludes input inventory; identical complete payload hashes for OpenRead/CopyTo. Not peak/RSS or timing."
            }, new JsonSerializerOptions { WriteIndented = true }));
            return 0;
        }

        private sealed class DigestSink(string operation, CancellationTokenSource cancellation, HashSet<string> before) : Stream {
            private readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            private bool _disposed;
            internal long Bytes { get; private set; }
            internal long LiveSpoolBytes { get; private set; }
            internal string Digest() => Convert.ToHexString(_hash.GetHashAndReset()).ToLowerInvariant();
            public override void Write(byte[] buffer, int offset, int count) {
                _hash.AppendData(buffer, offset, count); Bytes += count;
                if (Bytes == count) {
                    LiveSpoolBytes = Directory.GetDirectories(Path.GetTempPath(), "officeimo-archive-*")
                        .Except(before).SelectMany(Directory.GetFiles).Sum(p => new FileInfo(p).Length);
                    if (operation == "DestinationFailure") throw new IOException("Synthetic destination failure");
                    if (operation == "DestinationCancellation") cancellation.Cancel();
                }
            }
            public override bool CanWrite => !_disposed;
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            protected override void Dispose(bool disposing) { if (disposing) { _disposed = true; _hash.Dispose(); } base.Dispose(disposing); }
        }
    }
}
