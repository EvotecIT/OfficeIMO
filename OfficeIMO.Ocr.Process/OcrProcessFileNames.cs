namespace OfficeIMO.Ocr.Process;

/// <summary>Safe temporary filename helpers shared by process-based OCR providers.</summary>
public static class OcrProcessFileNames {
    /// <summary>Returns a bounded, filesystem-safe extension for an OCR input.</summary>
    public static string GetSafeExtension(string? fileName, string? mediaType) {
        string extension = Path.GetExtension(fileName ?? string.Empty);
        if (string.IsNullOrWhiteSpace(extension)) extension = mediaType switch {
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
