using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Archives;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;

if (args.Length > 0 && args[0] == "--large-input-evidence")
    return OfficeIMO.Archives.Qualification.LargeInputEvidence.Run(args.Skip(1).ToArray());
if (args.Length > 0 && args[0] == "--cab-read-evidence")
    return OfficeIMO.Archives.Qualification.CabinetReadEvidence.Run(args.Skip(1).ToArray());

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
    var expected = Directory.GetFiles(args[4], "*", SearchOption.AllDirectories).Select(p => Hash(File.ReadAllBytes(p))).OrderBy(h => h, StringComparer.Ordinal).ToArray();
    Require(actual.SequenceEqual(expected), "MSI payload differs from independent extraction");
    var extracted = ArchiveExtractor.Extract(args[3], Path.Combine(output, "msi"));
    var published = extracted.Files.Select(f => Hash(File.ReadAllBytes(Path.Combine(output, "msi", f.RelativePath)))).OrderBy(h => h, StringComparer.Ordinal);
    Require(published.SequenceEqual(expected), "Published MSI payload mismatch");
    results.Add(new { file = Path.GetFileName(args[3]), sha256 = Hash(File.ReadAllBytes(args[3])), members = actual.Length, verified = true });
}
if (args.Length >= 7) {
    using var cab = ArchiveInspector.Open(args[5]);
    using var member = cab.OpenRead(0);
    Require(HashStream(member) == Hash(File.ReadAllBytes(args[6])), "Native makecab member mismatch");
    Require(cab.Verify(new[] { 0 }).Members.All(m => m.Status == ArchiveVerificationStatus.Verified), "Native CAB verification failed");
    results.Add(new { file = Path.GetFileName(args[5]), sha256 = Hash(File.ReadAllBytes(args[5])), verified = true });
}
File.WriteAllText(args[2], JsonSerializer.Serialize(new {
    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
    assembly = typeof(ArchiveInspector).Assembly.GetName().Version?.ToString(), results
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Qualified {results.Count} packed-consumer scenarios.");
return 0;

static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
static string HashStream(Stream stream) => Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
static void Require(bool value, string message) { if (!value) throw new InvalidDataException(message); }
