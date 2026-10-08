using OfficeIMO.Drawing;

namespace OfficeIMO.ConversionConsistency;

/// <summary>Image export formats covered by the full-page comparison contract.</summary>
internal static class ComparedImageFormats {
    internal static IReadOnlyList<OfficeImageExportFormat> All { get; } = Array.AsReadOnly(new[] {
        OfficeImageExportFormat.Png,
        OfficeImageExportFormat.Svg,
        OfficeImageExportFormat.Jpeg,
        OfficeImageExportFormat.Tiff,
        OfficeImageExportFormat.Webp
    });
}
