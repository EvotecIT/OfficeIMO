using System.Security.Cryptography;
using System.IO.Compression;
using System.Text.Json;
using OfficeIMO.Archives;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;

#if NET8_0_OR_GREATER
if (args.Length > 0 && args[0] == "--cab-staging-failure-evidence")
    return OfficeIMO.Archives.Qualification.CabinetCopyEvidence.RunStagingFailure(args.Skip(1).ToArray());
if (args.Length > 0 && args[0] == "--cab-copy-evidence")
    return OfficeIMO.Archives.Qualification.CabinetCopyEvidence.Run(args.Skip(1).ToArray());
if (args.Length > 0 && args[0] == "--large-input-evidence")
    return OfficeIMO.Archives.Qualification.LargeInputEvidence.Run(args.Skip(1).ToArray());
if (args.Length > 0 && args[0] == "--cab-read-evidence")
    return OfficeIMO.Archives.Qualification.CabinetReadEvidence.Run(args.Skip(1).ToArray());
if (args.Length > 0 && args[0] == "--reader-snapshot-evidence")
    return OfficeIMO.Archives.Qualification.ReaderSnapshotEvidence.Run(args.Skip(1).ToArray());
if (args.Length > 0 && args[0] == "--snapshot-restore-evidence")
    return OfficeIMO.Archives.Qualification.ReaderSnapshotEvidence.RunRestoreFailure(args.Skip(1).ToArray());
#endif

// Opt-in consumer of packed public APIs. It does not download or execute archive contents.
if (args.Length < 3) throw new ArgumentException("Expected: <ZIP corpus> <fresh output directory> <JSON report> [MSI] [independent MSI payload directory] [makecab CAB] [original member]");
string corpus = Path.GetFullPath(args[0]);
string output = Path.GetFullPath(args[1]);
if (Directory.Exists(output)) throw new IOException("Use a fresh task-owned output directory.");
Directory.CreateDirectory(output);
var results = new List<object>();
using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(corpus, "manifest.json")));
foreach (var fixture in manifest.RootElement.GetProperty("fixtures").EnumerateArray()) {
    string name = fixture.GetProperty("file").GetString()!;
    string path = Path.Combine(corpus, name);
    string inputHash = Hash(File.ReadAllBytes(path));
    Require(inputHash == fixture.GetProperty("sha256").GetString(), "Fixture hash mismatch: " + name);
    int? codePage = fixture.GetProperty("legacyCodePage").ValueKind == JsonValueKind.Null ? null : fixture.GetProperty("legacyCodePage").GetInt32();
    if (codePage.HasValue) {
        try { ArchiveReader.TraverseZip(path); throw new InvalidDataException("Unflagged non-ASCII name accepted without a policy"); }
        catch (NotSupportedException) { }
    }
    var traversal = ArchiveReader.TraverseZip(path, new() { ZipLegacyCodePage = codePage });
    using var inspection = ArchiveInspector.Open(path, new() { ZipLegacyCodePage = codePage });
    var entries = fixture.GetProperty("entries").EnumerateArray().ToArray();
    Require(entries.Length == inspection.Entries.Count, "Entry count mismatch");
    Require(traversal.Entries.Select(e => e.RawFullName).SequenceEqual(entries.Select(e => e.GetProperty("name").GetString())), "Traversal name mismatch");
    using var batch = inspection.ReadMembers(inspection.Entries.Select(e => e.Ordinal).ToArray(), 64L * 1024 * 1024);
    for (int i = 0; i < entries.Length; i++) {
        var expected = entries[i];
        Require(batch.Members[i].Entry.RawName == expected.GetProperty("name").GetString(), "Name mismatch");
        Require(HashStream(batch.Members[i].Content) == expected.GetProperty("sha256").GetString(), "Member hash mismatch");
        using var copied = new MemoryStream(); inspection.CopyTo(i, copied);
        Require(Hash(copied.ToArray()) == expected.GetProperty("sha256").GetString(), "Verified copy mismatch");
    }
    Require(inspection.Verify().Members.All(m => m.Status == ArchiveVerificationStatus.Verified), "Verification failed");
    var read = ArchiveReader.ReadZip(File.ReadAllBytes(path), new() { ZipLegacyCodePage = codePage });
    Require(read.Count == entries.Length, "Memory read count mismatch");
    for (int i = 0; i < entries.Length; i++) {
        Require(read[i].RelativePath == entries[i].GetProperty("name").GetString(), "Memory read name mismatch");
        Require(Hash(read[i].Data) == entries[i].GetProperty("sha256").GetString(), "Memory read payload mismatch");
    }
    var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler(new() { ZipLegacyCodePage = codePage }).Build();
    var chunks = reader.Read(path, new() { ComputeHashes = false }).ToArray();
    Require(chunks.Any(c => c.Text.Contains("Zażółć")) && chunks.All(c => !(c.Warnings?.Any() ?? false)), "Reader payload mismatch");
    using var oracleInput = File.OpenRead(path);
    using var timestampOracle = new ZipArchive(oracleInput, ZipArchiveMode.Read);
    var expectedTimes = entries.Select((entry, index) => new {
        Name = entry.GetProperty("name").GetString()!, Time = timestampOracle.Entries[index].LastWriteTime.UtcDateTime
    }).ToDictionary(entry => entry.Name, entry => entry.Time, StringComparer.Ordinal);
    foreach (var chunk in chunks) {
        string memberName = chunk.Location.Path.Substring(chunk.Location.Path.LastIndexOf("::", StringComparison.Ordinal) + 2);
        Require(chunk.SourceLastWriteUtc == expectedTimes[memberName], "Reader timestamp differs from platform ZipArchive: " + memberName);
    }
    var extracted = ArchiveExtractor.Extract(path, Path.Combine(output, name), new() { ZipLegacyCodePage = codePage });
    Require(extracted.Files.Count == entries.Length, "Extracted count mismatch");
    foreach (var expected in entries) {
        string relative = expected.GetProperty("name").GetString()!;
        Require(extracted.Files.Any(f => f.RelativePath == relative), "Extracted name mismatch");
        Require(Hash(File.ReadAllBytes(Path.Combine(output, name, relative))) == expected.GetProperty("sha256").GetString(), "Extracted payload mismatch");
    }
    // Legacy fixtures require the same explicit policy for filesystem publication.
    results.Add(new { file = name, sha256 = inputHash, members = entries.Length, encoding = codePage, verified = true });
}
if (args.Length >= 5) {
    using var msi = ArchiveInspector.Open(args[3]);
    using var batch = msi.ReadMembers(msi.Entries.Select(e => e.Ordinal).ToArray(), 128L * 1024 * 1024);
    var actual = batch.Members.Select(m => HashStream(m.Content)).OrderBy(h => h, StringComparer.Ordinal).ToArray();
    var expected = ExpectedMsiHashes(args[4]);
    Require(actual.SequenceEqual(expected), "MSI payload differs from independent extraction");
    foreach (var entry in msi.Entries) {
        using var copied = new MemoryStream(); msi.CopyTo(entry.Ordinal, copied); copied.Position = 0;
        Require(HashStream(copied) == actualHash(entry.Ordinal), "Verified MSI copy mismatch");
    }
    var extracted = ArchiveExtractor.Extract(args[3], Path.Combine(output, "msi"));
    var published = extracted.Files.Select(f => Hash(File.ReadAllBytes(Path.Combine(output, "msi", f.RelativePath)))).OrderBy(h => h, StringComparer.Ordinal);
    Require(published.SequenceEqual(expected), "Published MSI payload mismatch");
    results.Add(new { file = Path.GetFileName(args[3]), sha256 = Hash(File.ReadAllBytes(args[3])), members = actual.Length, verified = true });
    string actualHash(int ordinal) {
        batch.Members[ordinal].Content.Position = 0;
        return HashStream(batch.Members[ordinal].Content);
    }
}
if (args.Length >= 7) {
    using var cab = ArchiveInspector.Open(args[5]);
    using var member = cab.OpenRead(0);
    Require(HashStream(member) == Hash(File.ReadAllBytes(args[6])), "Native makecab member mismatch");
    Require(cab.Verify(new[] { 0 }).Members.All(m => m.Status == ArchiveVerificationStatus.Verified), "Native CAB verification failed");
    using var copied = new MemoryStream(); cab.CopyTo(0, copied);
    Require(Hash(copied.ToArray()) == Hash(File.ReadAllBytes(args[6])), "Native CAB verified copy mismatch");
    results.Add(new { file = Path.GetFileName(args[5]), sha256 = Hash(File.ReadAllBytes(args[5])), verified = true });
}
File.WriteAllText(args[2], JsonSerializer.Serialize(new {
    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
    assembly = typeof(ArchiveInspector).Assembly.GetName().Version?.ToString(), results
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Qualified {results.Count} packed-consumer scenarios.");
return 0;

static string Hash(byte[] bytes) { using var sha = SHA256.Create(); return Hex(sha.ComputeHash(bytes)); }
static string HashStream(Stream stream) { using var sha = SHA256.Create(); return Hex(sha.ComputeHash(stream)); }
static string Hex(byte[] bytes) => BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
static string[] ExpectedMsiHashes(string path) {
    if (!File.Exists(path)) return Directory.GetFiles(path, "*", SearchOption.AllDirectories)
        .Select(p => Hash(File.ReadAllBytes(p))).OrderBy(h => h, StringComparer.Ordinal).ToArray();
    using var manifest = JsonDocument.Parse(File.ReadAllText(path));
    return manifest.RootElement.GetProperty("members").EnumerateArray().Select(m => m.GetProperty("sha256").GetString()!)
        .OrderBy(h => h, StringComparer.Ordinal).ToArray();
}
static void Require(bool value, string message) { if (!value) throw new InvalidDataException(message); }
