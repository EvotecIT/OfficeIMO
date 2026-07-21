using OfficeIMO.Drawing;
using DocumentFormat.OpenXml;

namespace OfficeIMO.PowerPoint;

/// <summary>Machine-readable catalog of PowerPoint formats supported or classified by OfficeIMO.</summary>
public static class PowerPointFormatCatalog {
    private static readonly IReadOnlyList<OfficeFormatDescriptor> Formats = Array.AsReadOnly(new[] {
        new OfficeFormatDescriptor("PowerPoint.Ppt", ".ppt", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("PowerPoint.Pot", ".pot", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("PowerPoint.Pps", ".pps", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.SlideShow,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("PowerPoint.Ppa", ".ppa", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.AddIn,
            OfficeFormatGeneration.Legacy, OfficeFormatEncoding.CompoundBinary, macroEnabled: true),
        new OfficeFormatDescriptor("PowerPoint.Pptx", ".pptx", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: false),
        new OfficeFormatDescriptor("PowerPoint.Pptm", ".pptm", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true),
        new OfficeFormatDescriptor("PowerPoint.Potx", ".potx", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: false),
        new OfficeFormatDescriptor("PowerPoint.Potm", ".potm", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.Template,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true),
        new OfficeFormatDescriptor("PowerPoint.Ppsx", ".ppsx", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.SlideShow,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: false),
        new OfficeFormatDescriptor("PowerPoint.Ppsm", ".ppsm", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.SlideShow,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true),
        new OfficeFormatDescriptor("PowerPoint.Ppam", ".ppam", OfficeDocumentFamily.PowerPoint, OfficeDocumentKind.AddIn,
            OfficeFormatGeneration.Modern, OfficeFormatEncoding.OpenXml, macroEnabled: true)
    });

    /// <summary>Gets every PowerPoint format in the compatibility matrix.</summary>
    public static IReadOnlyList<OfficeFormatDescriptor> All => Formats;

    /// <summary>Gets the descriptor for a file extension.</summary>
    public static OfficeFormatDescriptor GetByExtension(string pathOrExtension) {
        if (!TryGetByExtension(pathOrExtension, out OfficeFormatDescriptor? format)) {
            throw new NotSupportedException($"Unsupported PowerPoint format extension '{GetExtension(pathOrExtension)}'.");
        }

        return format!;
    }

    /// <summary>Tries to resolve a file path or extension to a PowerPoint format descriptor.</summary>
    public static bool TryGetByExtension(string? pathOrExtension, out OfficeFormatDescriptor? format) {
        string extension = GetExtension(pathOrExtension);
        format = Formats.FirstOrDefault(candidate =>
            string.Equals(candidate.Extension, extension, StringComparison.OrdinalIgnoreCase));
        return format != null;
    }

    internal static OfficeFormatDescriptor GetDescriptor(PowerPointFileFormat format, string? path = null) {
        if (TryGetByExtension(path, out OfficeFormatDescriptor? descriptor)
            && (PowerPointPresentation.IsLegacyBinaryFormat(format))
                == (descriptor!.Generation == OfficeFormatGeneration.Legacy)) {
            return descriptor;
        }

        return format switch {
            PowerPointFileFormat.Ppt => Formats[0],
            PowerPointFileFormat.Pot => Formats[1],
            PowerPointFileFormat.Pps => Formats[2],
            _ => Formats[4]
        };
    }

    internal static OfficeFormatDescriptor GetDescriptor(PresentationDocumentType documentType) => documentType switch {
        PresentationDocumentType.MacroEnabledPresentation => Formats[5],
        PresentationDocumentType.Template => Formats[6],
        PresentationDocumentType.MacroEnabledTemplate => Formats[7],
        PresentationDocumentType.Slideshow => Formats[8],
        PresentationDocumentType.MacroEnabledSlideshow => Formats[9],
        PresentationDocumentType.AddIn => Formats[10],
        _ => Formats[4]
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
