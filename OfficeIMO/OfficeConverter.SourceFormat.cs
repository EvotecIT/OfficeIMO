namespace OfficeIMO;

public static partial class OfficeConverter {
    /// <summary>Resolves an explicit codec without changing the physical source used for admission and publication checks.</summary>
    internal static string ConversionSourceName(string inputPath, OfficeConversionOptions settings) {
        string name = Path.GetFileName(inputPath);
        string? extension = settings.SourceExtension;
        if (extension == null && settings.SourceFormat is { } format && FormatForExtension(Path.GetExtension(name)) != format)
            extension = DefaultExtension(format);
        return extension == null ? name : Path.ChangeExtension(name, extension);
    }

    internal static string? NormalizeSourceExtension(string? extension) {
        if (extension == null) return null;
        string value = extension.Trim();
        if (!value.StartsWith(".", StringComparison.Ordinal)) value = "." + value;
        if (!OfficeConversionCapabilityCatalog.All.Any(route => route.SourceExtensions.Contains(value, StringComparer.OrdinalIgnoreCase)))
            throw new ArgumentException("Choose a source extension from the canonical conversion catalog.", nameof(extension));
        return value.ToLowerInvariant();
    }
}
