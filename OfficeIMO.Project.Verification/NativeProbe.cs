using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using OfficeIMO.Core.Internal;

// Opt-in, fixture-only experiments. None of these operations claims an MPP codec.
internal static class NativeProbe {
    internal static int Run(string path, string directory) {
        var input = new FileInfo(path);
        if (input.Length > 64 * 1024 * 1024) throw new InvalidDataException("Fixture exceeds 64 MiB.");
        if (Directory.Exists(directory)) throw new IOException("Choose a new output directory.");
        byte[] bytes = File.ReadAllBytes(input.FullName);
        var limits = new OfficeCompoundReadOptions(4096, 2048, 32 * 1024 * 1024, 64 * 1024 * 1024);
        if (!OfficeCompoundFileReader.TryRead(bytes, limits, out var compound, out var error) || compound == null)
            throw new InvalidDataException(error);
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "unchanged.mpp"), bytes);
        byte[] rewritten = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]>());
        File.WriteAllBytes(Path.Combine(directory, "rewritten.mpp"), rewritten);
        if (!OfficeCompoundFileReader.TryRead(rewritten, limits, out var reopened, out error) || reopened == null)
            throw new InvalidDataException(error);
        bool equalStreams = compound.Streams.Count == reopened.Streams.Count && compound.Streams.All(
            s => reopened.Streams.TryGetValue(s.Key, out var value) && s.Value.SequenceEqual(value));
        if (!equalStreams) throw new InvalidDataException("Compound rewrite changed stream bytes.");
        // A same-width replacement deliberately isolates directory/stream retention from
        // record growth. Its external readback does not qualify general task editing.
        string? taskStream = compound.Streams.Keys.SingleOrDefault(p => p.EndsWith("/TBkndTask/Var2Data", StringComparison.Ordinal));
        if (taskStream != null) {
            byte[] changed = (byte[])compound.Streams[taskStream].Clone();
            byte[] oldName = Encoding.Unicode.GetBytes("Build\0");
            var offsets = Enumerable.Range(0, Math.Max(0, changed.Length - oldName.Length + 1))
                .Where(i => changed.AsSpan(i, oldName.Length).SequenceEqual(oldName)).ToArray();
            if (offsets.Length == 1) {
                Encoding.Unicode.GetBytes("Craft\0").CopyTo(changed, offsets[0]);
                File.WriteAllBytes(Path.Combine(directory, "field-edit.mpp"), OfficeCompoundFileWriter.Rewrite(compound,
                    new Dictionary<string, byte[]> { [taskStream] = changed }));
            }
        }
        // A seed-free container experiment: conventional MPP14 directory names and root
        // class identity alone, with no copied Project records or proprietary template.
        var minimal = new[] {
            new OfficeCompoundStream("Props14", new byte[16]),
            new OfficeCompoundStream("   114/Props", Array.Empty<byte>()),
            new OfficeCompoundStream("   114/TBkndTask/FixedMeta", new byte[16]),
            new OfficeCompoundStream("   114/TBkndTask/FixedData", Array.Empty<byte>())
        };
        File.WriteAllBytes(Path.Combine(directory, "minimal-new.mpp"), OfficeCompoundFileWriter.Write(minimal,
            new Guid("74b78f3a-c8c8-11d1-be11-00c04fb6faf1")));
        foreach (var stream in compound.Streams.Where(s => s.Key.Contains("/TBkndTask/") || s.Key.Contains("/TBkndCal/") || s.Key.Contains("/TBkndAssn/")))
            File.WriteAllBytes(Path.Combine(directory, stream.Key.Substring(stream.Key.IndexOf("TBknd", StringComparison.Ordinal)).Replace('/', '-')), stream.Value);
        var matches = new List<object>();
        foreach (var stream in compound.Streams) {
            foreach (string text in new[] { "Delivery", "Design", "Build", "Engineer", "Standard" }) {
                byte[] pattern = Encoding.Unicode.GetBytes(text + "\0");
                for (int offset = 0; offset <= stream.Value.Length - pattern.Length; offset++) {
                    if (stream.Value.AsSpan(offset, pattern.Length).SequenceEqual(pattern))
                        matches.Add(new { stream = stream.Key, offset, text });
                }
            }
        }
        var report = new {
            input = input.Name, sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
            compound.RootEntry.ClassId, equalStreams,
            streams = compound.Streams.Select(s => new { path = s.Key, bytes = s.Value.Length,
                sha256 = Convert.ToHexString(SHA256.HashData(s.Value)).ToLowerInvariant() }),
            stringMatches = matches,
            extractedRecords = File.Exists(Path.ChangeExtension(path, ".xml")) ? NativeFixtureRecords.ExtractAndCompare(compound, Path.ChangeExtension(path, ".xml")) : null,
            limitation = "String locations are feasibility observations, not validated task/calendar/assignment record decoding."
        };
        string json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(Path.Combine(directory, "native-probe.json"), json);
        Console.WriteLine(JsonSerializer.Serialize(new { report.input, report.equalStreams, streamCount = compound.Streams.Count, report.stringMatches }));
        return 0;
    }
}
