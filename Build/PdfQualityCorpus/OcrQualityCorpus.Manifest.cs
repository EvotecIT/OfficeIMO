using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.PdfQualityCorpus;

internal static partial class OcrQualityCorpus {
    private static readonly JsonSerializerOptions Json = new() { PropertyNameCaseInsensitive = true, WriteIndented = true };

    private static async Task<(OcrQualityManifest Manifest, OcrQualityLabel[] Labels, string Hash)> ReadManifestAsync(
        string manifestPath, string root, CancellationToken token) {
        byte[] bytes = await ReadBoundedAsync(manifestPath, token);
        var manifest = JsonSerializer.Deserialize<OcrQualityManifest>(bytes, Json) ?? throw new InvalidDataException("Empty OCR manifest.");
        if (manifest.Version != 1 || manifest.Sources.Count is < 1 or > 100) throw new InvalidDataException("Unsupported OCR manifest.");
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (OcrQualitySource source in manifest.Sources) {
            if (!files.Add(source.File) || !Uri.TryCreate(source.Authority, UriKind.Absolute, out Uri? authority) ||
                authority.Scheme != "https" || string.IsNullOrWhiteSpace(source.License)) throw new InvalidDataException("Invalid OCR source provenance.");
            // Verify every source, not just those reached by a selected case.
            await ReadVerifiedAsync(root, source.File, source.Sha256, token);
        }
        byte[] labelBytes = await ReadVerifiedAsync(root, manifest.Labels, manifest.LabelsSha256, token);
        OcrQualityLabel[] labels = JsonSerializer.Deserialize<OcrQualityLabel[]>(labelBytes, Json) ?? throw new InvalidDataException("Empty OCR labels.");
        if (labels.Length is < 1 or > 200) throw new InvalidDataException("OCR label count is out of bounds.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (OcrQualityLabel label in labels) {
            if (!ids.Add(label.Id) || string.IsNullOrWhiteSpace(label.Id) || label.Id.Length > 64 ||
                !label.Id.All(c => char.IsAsciiLetterOrDigit(c) || c == '-') || !manifest.Sources.Any(source => source.File == label.File) ||
                string.IsNullOrWhiteSpace(label.DocumentClass) || string.IsNullOrWhiteSpace(label.Language) ||
                string.IsNullOrWhiteSpace(label.Expected) || label.Expected.Length > 50_000 || label.Page < 1 ||
                !Unit(label.MaximumCer) || !Unit(label.MaximumWer)) throw new InvalidDataException("Invalid OCR quality label.");
            if (!label.File.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) && label.Page != 1)
                throw new InvalidDataException("Raster qualification labels must target the first page.");
            if (label.Region != null) {
                if (label.Region.Length != 4) throw new InvalidDataException("An OCR label region needs four normalized coordinates.");
                _ = new PdfOcrPageRegion(label.Page, label.Region[0], label.Region[1], label.Region[2], label.Region[3]);
            }
        }
        return (manifest, labels, Hash(bytes));
    }

    private static string Resolve(string root, string file) {
        if (string.IsNullOrWhiteSpace(file) || Path.IsPathRooted(file) || file.Contains('\\') ||
            file.Split('/').Any(part => part is "" or "." or "..")) throw new InvalidDataException("OCR paths must stay inside the asset root.");
        string current = Path.GetFullPath(root);
        if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) throw new InvalidDataException("The asset root must not be a link.");
        foreach (string part in file.Split('/')) {
            current = Path.Combine(current, part);
            if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) throw new InvalidDataException("OCR assets must not traverse links.");
        }
        return current;
    }

    private static async Task<byte[]> ReadVerifiedAsync(string root, string file, string expected, CancellationToken token) {
        if (expected.Length != 64 || !expected.All(Uri.IsHexDigit)) throw new InvalidDataException("Invalid OCR asset digest.");
        byte[] bytes = await ReadBoundedAsync(Resolve(root, file), token);
        if (!Hash(bytes).Equals(expected, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("OCR asset digest mismatch.");
        return bytes;
    }

    private static async Task<byte[]> ReadBoundedAsync(string path, CancellationToken token) {
        const int maximum = 25 * 1024 * 1024;
        await using FileStream stream = File.OpenRead(path);
        if (stream.Length > maximum) throw new InvalidDataException("OCR corpus asset exceeds the byte budget.");
        using var result = new MemoryStream();
        byte[] buffer = new byte[81920];
        int count;
        while ((count = await stream.ReadAsync(buffer, token)) != 0) {
            if (result.Length + count > maximum) throw new InvalidDataException("OCR corpus asset exceeds the byte budget.");
            await result.WriteAsync(buffer.AsMemory(0, count), token);
        }
        return result.ToArray();
    }

    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static bool Unit(double number) => double.IsFinite(number) && number >= 0 && number <= 1;
}
