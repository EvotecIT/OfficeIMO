using OfficeIMO.Drawing;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word;

/// <summary>Machine-readable catalog of Word formats supported or classified by OfficeIMO.</summary>
public static class WordFormatCatalog {
    private static readonly IReadOnlyList<OfficeFormatDescriptor> Formats = Array.AsReadOnly(new[] {
        new OfficeFormatDescriptor("Word.Doc", ".doc", OfficeDocumentFamily.Word, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("Word.Dot", ".dot", OfficeDocumentFamily.Word, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("Word.Docx", ".docx", OfficeDocumentFamily.Word, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: false),
        new OfficeFormatDescriptor("Word.Docm", ".docm", OfficeDocumentFamily.Word, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true),
        new OfficeFormatDescriptor("Word.Dotx", ".dotx", OfficeDocumentFamily.Word, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: false),
        new OfficeFormatDescriptor("Word.Dotm", ".dotm", OfficeDocumentFamily.Word, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true)
    });

    /// <summary>Gets every Word format in the compatibility matrix.</summary>
    public static IReadOnlyList<OfficeFormatDescriptor> All => Formats;

    /// <summary>Gets the descriptor for a file extension.</summary>
    public static OfficeFormatDescriptor GetByExtension(string pathOrExtension) {
        if (!TryGetByExtension(pathOrExtension, out OfficeFormatDescriptor? format)) {
            throw new NotSupportedException($"Unsupported Word format extension '{GetExtension(pathOrExtension)}'.");
        }

        return format!;
    }

    /// <summary>Tries to resolve a file path or extension to a Word format descriptor.</summary>
    public static bool TryGetByExtension(string? pathOrExtension, out OfficeFormatDescriptor? format) {
        string extension = GetExtension(pathOrExtension);
        format = Formats.FirstOrDefault(candidate =>
            string.Equals(candidate.Extension, extension, StringComparison.OrdinalIgnoreCase));
        return format != null;
    }

    internal static OfficeFormatDescriptor GetDescriptor(WordFileFormat format, string? path = null) {
        if (TryGetByExtension(path, out OfficeFormatDescriptor? descriptor)
            && (format == WordFileFormat.Doc) == (descriptor!.Generation == OfficeFormatGeneration.Legacy)) {
            return descriptor;
        }

        return format == WordFileFormat.Doc ? Formats[0] : Formats[2];
    }

    internal static OfficeFormatDescriptor GetDescriptor(WordprocessingDocumentType documentType) => documentType switch {
        WordprocessingDocumentType.MacroEnabledDocument => Formats[3],
        WordprocessingDocumentType.Template => Formats[4],
        WordprocessingDocumentType.MacroEnabledTemplate => Formats[5],
        _ => Formats[2]
    };

    private static string GetExtension(string? pathOrExtension) {
        if (string.IsNullOrWhiteSpace(pathOrExtension)) return string.Empty;
        string value = pathOrExtension!.Trim();
        string extension = value.StartsWith(".", StringComparison.Ordinal) && value.IndexOfAny(new[] { '/', '\\' }) < 0
            ? value
            : Path.GetExtension(value);
        return extension.ToLowerInvariant();
    }
}
