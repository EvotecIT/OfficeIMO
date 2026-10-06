using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using OfficeIMO.Provenance.C2pa;

if (args.Length != 3) {
    Console.Error.WriteLine("Usage: <corpus-directory> <trusted-c2patool-path> <new-evidence.json>");
    return 2;
}
string corpusRoot = Path.GetFullPath(args[0]);
string toolPath = Path.GetFullPath(args[1]);
string outputPath = Path.GetFullPath(args[2]);
if (File.Exists(outputPath)) throw new IOException("Use a new evidence path; existing evidence is not overwritten.");
using JsonDocument corpus = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "corpus.json")));
var verifier = new C2paToolProvenanceVerifier(toolPath);
C2paToolAvailability availability = verifier.CheckAvailability();
if (!availability.Available) throw new InvalidOperationException(availability.Diagnostic);
var results = new List<object>();
int failures = 0;
string scratch = Path.Combine(Path.GetTempPath(), "officeimo-provenance-interop-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(scratch);
try {
    foreach (JsonElement entry in corpus.RootElement.GetProperty("cases").EnumerateArray()) {
        string name = entry.GetProperty("file").GetString()!;
        if (Path.GetFileName(name) != name || name.Contains('/') || name.Contains('\\')) throw new InvalidDataException("Invalid corpus filename.");
        string path = Path.Combine(corpusRoot, name);
        if (new FileInfo(path).Length > 8 * 1024 * 1024) throw new InvalidDataException("Oversized fixture.");
        byte[] input = File.ReadAllBytes(path);
        string inputHash = Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant();
        if (inputHash != entry.GetProperty("sha256").GetString()) throw new InvalidDataException("Fixture hash mismatch: " + name);
        try {
            var before = OfficeProvenanceInspector.Inspect(input, name);
            var verified = verifier.Verify(path);
            var removed = OfficeProvenanceRemover.Remove(input, name);
            byte[] cleaned = removed.ToArray();
            string cleanedPath = Path.Combine(scratch, name);
            File.WriteAllBytes(cleanedPath, cleaned);
            var after = OfficeProvenanceInspector.Inspect(cleaned, name);
            var afterVerified = verifier.Verify(cleanedPath);
            bool decoded = OfficeRasterImageDecoder.TryDecode(input, out var originalImage) &&
                OfficeRasterImageDecoder.TryDecode(cleaned, out var cleanedImage) &&
                originalImage!.Width == cleanedImage!.Width && originalImage.Height == cleanedImage.Height &&
                originalImage.GetPixels().AsSpan().SequenceEqual(cleanedImage.GetPixels());
            bool passed = before.HasC2paManifest == entry.GetProperty("hasManifest").GetBoolean() &&
                verified.Status.ToString() == entry.GetProperty("verification").GetString() &&
                !after.HasC2paManifest && afterVerified.Status == OfficeProvenanceVerificationStatus.NotPresent && decoded &&
                (!before.HasC2paManifest ? input.AsSpan().SequenceEqual(cleaned) : removed.WasChanged);
            if (!passed) failures++;
            results.Add(new { file = name, inputHash, outputHash = Convert.ToHexString(SHA256.HashData(cleaned)).ToLowerInvariant(),
                manifest = before.HasC2paManifest, structural = before.Evidence.Select(e => e.IsStructurallyValid),
                verification = verified.Status.ToString(), verificationFindings = verified.Findings,
                afterVerification = afterVerified.Status.ToString(), sameDecodedPixels = decoded, passed });
            Console.WriteLine($"{name}: {verified.Status}; after={afterVerified.Status}; pixels={decoded}; pass={passed}");
        } catch (Exception exception) {
            failures++;
            results.Add(new { file = name, inputHash, passed = false, error = exception.ToString() });
            Console.Error.WriteLine(name + ": " + exception.Message);
        }
    }
    string json = JsonSerializer.Serialize(new {
        schema = "officeimo.provenance.interop.evidence.v1", utc = DateTimeOffset.UtcNow,
        corpusRevision = corpus.RootElement.GetProperty("revision").GetString(),
        corpusLicense = corpus.RootElement.GetProperty("license").GetString(),
        toolVersion = availability.Version, toolSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(toolPath))).ToLowerInvariant(),
        platform = System.Runtime.InteropServices.RuntimeInformation.RuntimeIdentifier,
        results, failures
    }, new JsonSerializerOptions { WriteIndented = true });
    Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
    using var output = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write);
    using var writer = new StreamWriter(output);
    writer.Write(json);
} finally { Directory.Delete(scratch, recursive: true); }
return failures == 0 ? 0 : 1;
