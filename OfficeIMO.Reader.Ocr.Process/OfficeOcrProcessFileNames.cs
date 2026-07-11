using OfficeIMO.Reader;

namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>Safe temporary filename helpers shared by process-based OCR providers.</summary>
public static class OfficeOcrProcessFileNames {
    /// <summary>Returns a bounded, filesystem-safe extension for an OCR asset.</summary>
    public static string GetSafeAssetExtension(OfficeDocumentAsset asset) {
        if (asset == null) throw new ArgumentNullException(nameof(asset));
        string extension = asset.Extension ?? Path.GetExtension(asset.FileName ?? string.Empty);
        if (string.IsNullOrWhiteSpace(extension)) extension = asset.MediaType switch {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/tiff" => ".tiff",
            "image/bmp" => ".bmp",
            "image/webp" => ".webp",
            _ => ".bin"
        };
        string safe = new string(extension.Where(static character => char.IsLetterOrDigit(character) || character == '.').ToArray());
        if (safe.Length == 0 || safe.Length > 12) return ".bin";
        return safe[0] == '.' ? safe : "." + safe;
    }
}
