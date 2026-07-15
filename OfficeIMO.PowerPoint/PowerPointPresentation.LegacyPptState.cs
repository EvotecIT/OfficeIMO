using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private LegacyPptImportDiagnostic[] _legacyPptImportDiagnostics = Array.Empty<LegacyPptImportDiagnostic>();
        private string? _legacyPptSourcePath;

        /// <summary>Gets the detected physical format of the presentation source.</summary>
        public PowerPointFileFormat SourceFormat { get; private set; } = PowerPointFileFormat.Pptx;

        /// <summary>Gets the original legacy source path, or the current Open XML file association.</summary>
        public string? SourcePath => IsLegacyBinaryFormat(SourceFormat)
            ? _legacyPptSourcePath
            : string.IsNullOrWhiteSpace(FilePath) ? null : FilePath;

        /// <summary>Gets diagnostics produced while importing a legacy binary presentation.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> LegacyPptImportDiagnostics => _legacyPptImportDiagnostics;

        internal void MarkLoadedFromLegacyPpt(string? sourcePath, LegacyPptPresentation legacy,
            PowerPointFileFormat sourceFormat) {
            _legacyPptSourcePath = sourcePath;
            _legacyPptImportDiagnostics = legacy.Diagnostics.ToArray();
            SourceFormat = sourceFormat;
        }

        internal void MarkLoadedFromOpenXml() {
            SourceFormat = PowerPointFileFormat.Pptx;
        }

        internal static bool IsLegacyBinaryFormat(PowerPointFileFormat format) =>
            format == PowerPointFileFormat.Ppt || format == PowerPointFileFormat.Pot || format == PowerPointFileFormat.Pps;
    }
}
