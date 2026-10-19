using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;

namespace OfficeIMO.Archives.Qualification {
    /// <summary>Opt-in disk-snapshot evidence. Run in an isolated temporary directory without parallel consumers.</summary>
    internal static class ReaderSnapshotEvidence {
        internal static int Run(string[] args) {
            if (args.Length != 3) throw new ArgumentException("Expected: <large ZIP> <JSON report> <expected input copies>");
            string input = args[0];
            long inputLength = new FileInfo(input).Length;
            int expected = int.Parse(args[2]);
            var results = new List<object>();
            foreach (bool streamInput in new[] { false, true }) {
                foreach (string outcome in new[] { "complete", "cancel", "handler-error" }) {
                    var before = Snapshots().Select(f => f.FullName).ToHashSet(StringComparer.Ordinal);
                    using var cancellation = new CancellationTokenSource();
                    using var source = File.OpenRead(input);
                    source.Position = 13;
                    bool observed = false;
                    int copies = 0;
                    long snapshotBytes = 0;
                    var builder = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
                        Id = "qualification.snapshot-observer", Kind = ReaderInputKind.Text,
                        Extensions = new[] { ".txt" }, SupportsIncrementalStream = true,
                        ReadStream = Observe
                    }).AddZipHandler();
                    var reader = builder.Build();
                    try {
                        var chunks = streamInput ? reader.EnumerateChunks(source, "input.zip", new() { ComputeHashes = false }, cancellation.Token)
                            : reader.EnumerateChunks(input, new() { ComputeHashes = false }, cancellation.Token);
                        using var iterator = chunks.GetEnumerator();
                        while (iterator.MoveNext()) { }
                        if (outcome == "cancel") throw new InvalidDataException("Cancellation did not propagate.");
                    } catch (OperationCanceledException) when (outcome == "cancel") { }
                    if (!observed) throw new InvalidDataException("Leaf observer was not reached.");
                    if (!source.CanRead || source.Position != 13) throw new InvalidDataException("Caller stream ownership/position changed.");
                    if (Snapshots().Any(f => !before.Contains(f.FullName))) throw new InvalidDataException("Private snapshot leaked.");
                    if (copies != expected) throw new InvalidDataException($"Expected {expected} input copies; observed {copies}.");
                    results.Add(new { streamInput, outcome, copies, snapshotBytes, cleanedUp = true });

                    IEnumerable<ReaderChunk> Observe(Stream member, string? name, ReaderOptions options, CancellationToken token) {
                        using var text = new StreamReader(member, leaveOpen: true);
                        if (text.ReadToEnd() != "complete") throw new InvalidDataException("Leaf payload mismatch.");
                        var live = Snapshots().Where(f => !before.Contains(f.FullName)).ToArray();
                        copies = live.Count(f => f.Length == inputLength);
                        snapshotBytes = live.Sum(f => f.Length);
                        observed = true;
                        if (outcome == "cancel") { cancellation.Cancel(); token.ThrowIfCancellationRequested(); }
                        if (outcome == "handler-error") throw new IOException("Synthetic leaf failure");
                        yield return new ReaderChunk { Text = "complete", Kind = ReaderInputKind.Text };
                    }
                }
            }
            using var hash = File.OpenRead(input);
            File.WriteAllText(args[1], JsonSerializer.Serialize(new {
                runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
                assembly = typeof(ArchiveInspector).Assembly.GetName().Version?.ToString(),
                inputBytes = inputLength,
                inputSha256 = Convert.ToHexString(SHA256.HashData(hash)).ToLowerInvariant(), results,
                boundary = "Live private snapshot bytes observed inside a custom leaf handler; isolated opt-in process"
            }, new JsonSerializerOptions { WriteIndented = true }));
            return 0;
        }

        internal static int RunRestoreFailure(string[] args) {
            if (args.Length != 3) throw new ArgumentException("Expected: <ZIP> <JSON report> <expected leaked snapshots per route>");
            int expected = int.Parse(args[2]);
            var results = new List<object>();
            foreach (string route in new[] { "ReaderSync", "ReaderAsync", "Traversal" }) {
                var before = Snapshots().Select(f => f.FullName).ToHashSet(StringComparer.Ordinal);
                using var source = new RestoreFailureFile(args[0]);
                try {
                    if (route == "ReaderSync") {
                        using var snapshot = ReaderInputLimits.EnsureSeekableReadStream(source, 512L * 1024 * 1024,
                            CancellationToken.None, out _);
                    } else if (route == "ReaderAsync") {
                        using var snapshot = ReaderInputLimits.EnsureSeekableReadStreamAsync(source, 512L * 1024 * 1024)
                            .GetAwaiter().GetResult();
                    } else ArchiveReader.TraverseZip(source);
                    throw new InvalidDataException("Expected position restoration failure.");
                } catch (IOException error) when (error.Message == "Synthetic position restoration failure") { }
                if (!source.CanRead) throw new InvalidDataException("Caller stream closed.");
                var leaked = Snapshots().Where(f => !before.Contains(f.FullName)).ToArray();
                if (leaked.Length != expected) throw new InvalidDataException($"{route}: expected {expected} leaks, observed {leaked.Length}.");
                results.Add(new { route, leakedSnapshots = leaked.Length, leakedBytes = leaked.Sum(f => f.Length) });
            }
            File.WriteAllText(args[1], JsonSerializer.Serialize(new {
                runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
                assembly = typeof(ArchiveInspector).Assembly.GetName().Version?.ToString(), results,
                boundary = "Opt-in isolated process, successful copy followed by caller Position restoration failure"
            }, new JsonSerializerOptions { WriteIndented = true }));
            return 0;
        }

        private sealed class RestoreFailureFile : FileStream {
            internal RestoreFailureFile(string path) : base(path, FileMode.Open, FileAccess.Read) { base.Position = 13; }
            public override long Position {
                get => base.Position;
                set {
                    if (value == 13) throw new IOException("Synthetic position restoration failure");
                    base.Position = value;
                }
            }
        }

        private static FileInfo[] Snapshots() => new[] { "officeimo-reader-*", "officeimo-archive-*" }
            .SelectMany(pattern => Directory.GetDirectories(Path.GetTempPath(), pattern))
            .SelectMany(directory => Directory.GetFiles(directory))
            .Select(path => new FileInfo(path)).ToArray();
    }
}
