using OfficeIMO.Drawing;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Excel;

/// <summary>Machine-readable catalog of Excel formats supported or classified by OfficeIMO.</summary>
public static class ExcelFormatCatalog {
    private static readonly IReadOnlyList<OfficeFormatDescriptor> Formats = Array.AsReadOnly(new[] {
        new OfficeFormatDescriptor("Excel.Xls", ".xls", OfficeDocumentFamily.Excel, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("Excel.Xlt", ".xlt", OfficeDocumentFamily.Excel, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("Excel.Xla", ".xla", OfficeDocumentFamily.Excel, OfficeDocumentKind.AddIn,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("Excel.Xlm", ".xlm", OfficeDocumentFamily.Excel, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("Excel.Xlw", ".xlw", OfficeDocumentFamily.Excel, OfficeDocumentKind.Workspace,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: false),
        new OfficeFormatDescriptor("Excel.Xlsx", ".xlsx", OfficeDocumentFamily.Excel, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: false),
        new OfficeFormatDescriptor("Excel.Xlsm", ".xlsm", OfficeDocumentFamily.Excel, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true),
        new OfficeFormatDescriptor("Excel.Xltx", ".xltx", OfficeDocumentFamily.Excel, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: false),
        new OfficeFormatDescriptor("Excel.Xltm", ".xltm", OfficeDocumentFamily.Excel, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true),
        new OfficeFormatDescriptor("Excel.Xlam", ".xlam", OfficeDocumentFamily.Excel, OfficeDocumentKind.AddIn,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true),
        new OfficeFormatDescriptor("Excel.Xlsb", ".xlsb", OfficeDocumentFamily.Excel, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.BinaryOpenXml, macroEnabled: true)
    });

    /// <summary>Gets every Excel format in the compatibility matrix.</summary>
    public static IReadOnlyList<OfficeFormatDescriptor> All => Formats;

    /// <summary>Gets the descriptor for a file extension.</summary>
    public static OfficeFormatDescriptor GetByExtension(string pathOrExtension) {
        if (!TryGetByExtension(pathOrExtension, out OfficeFormatDescriptor? format)) {
            throw new NotSupportedException($"Unsupported Excel format extension '{GetExtension(pathOrExtension)}'.");
        }

        return format!;
    }

    /// <summary>Tries to resolve a file path or extension to an Excel format descriptor.</summary>
    public static bool TryGetByExtension(string? pathOrExtension, out OfficeFormatDescriptor? format) {
        string extension = GetExtension(pathOrExtension);
        format = Formats.FirstOrDefault(candidate =>
            string.Equals(candidate.Extension, extension, StringComparison.OrdinalIgnoreCase));
        return format != null;
    }

    internal static OfficeFormatDescriptor GetDescriptor(ExcelFileFormat format, string? path = null) {
        if (TryGetByExtension(path, out OfficeFormatDescriptor? descriptor)) {
            if (format == ExcelFileFormat.Xls && descriptor!.Generation == OfficeFormatGeneration.Legacy) return descriptor;
            if (format == ExcelFileFormat.Xlsb && descriptor!.Encoding == OfficeFormatEncoding.BinaryOpenXml) return descriptor;
            if (format == ExcelFileFormat.Xlsx && descriptor!.Encoding == OfficeFormatEncoding.OpenXml) return descriptor;
        }

        return format switch {
            ExcelFileFormat.Xls => Formats[0],
            ExcelFileFormat.Xlsb => Formats[10],
            _ => Formats[5]
        };
    }

    internal static OfficeFormatDescriptor GetDescriptor(SpreadsheetDocumentType documentType) => documentType switch {
        SpreadsheetDocumentType.MacroEnabledWorkbook => Formats[6],
        SpreadsheetDocumentType.Template => Formats[7],
        SpreadsheetDocumentType.MacroEnabledTemplate => Formats[8],
        SpreadsheetDocumentType.AddIn => Formats[9],
        _ => Formats[5]
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
