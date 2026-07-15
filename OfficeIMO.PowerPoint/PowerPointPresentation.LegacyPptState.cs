using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private LegacyPptImportDiagnostic[] _legacyPptImportDiagnostics = Array.Empty<LegacyPptImportDiagnostic>();
        private string? _legacyPptSourcePath;
        private LegacyPptPackage? _legacyPptPackage;
        private string? _legacyPptProjectionFingerprint;

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
            _legacyPptPackage = legacy.Package;
            _legacyPptProjectionFingerprint = CreatePackageFingerprint(_document!);
            SourceFormat = sourceFormat;
        }

        internal void MarkLoadedFromOpenXml() {
            _legacyPptPackage = null;
            _legacyPptProjectionFingerprint = null;
            SourceFormat = PowerPointFileFormat.Pptx;
        }

        internal bool CanPreserveOriginalLegacyPackage => _legacyPptPackage != null
            && _legacyPptProjectionFingerprint != null
            && string.Equals(_legacyPptProjectionFingerprint, CreatePackageFingerprint(_document!),
                StringComparison.Ordinal);

        internal bool TryCopyOriginalLegacyPackage(out byte[] bytes) {
            if (CanPreserveOriginalLegacyPackage) {
                bytes = _legacyPptPackage!.CopyOriginalBytes();
                return true;
            }
            bytes = Array.Empty<byte>();
            return false;
        }

        private void ClearLegacyPptPackageState() {
            _legacyPptPackage = null;
            _legacyPptProjectionFingerprint = null;
        }

        internal static bool IsLegacyBinaryFormat(PowerPointFileFormat format) =>
            format == PowerPointFileFormat.Ppt || format == PowerPointFileFormat.Pot || format == PowerPointFileFormat.Pps;
    }
}
