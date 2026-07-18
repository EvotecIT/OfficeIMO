using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private LegacyPptImportDiagnostic[] _legacyPptImportDiagnostics = Array.Empty<LegacyPptImportDiagnostic>();
        private string? _legacyPptSourcePath;
        private LegacyPptPackage? _legacyPptPackage;
        private LegacyPptProjectionMap? _legacyPptProjectionMap;
        private string? _legacyPptProjectionFingerprint;
        private LegacyPptProjectionFingerprint? _legacyPptPreservationFingerprint;
        private string[] _legacyPptLinkedOleDetails = Array.Empty<string>();
        private string[] _legacyPptActiveXDetails = Array.Empty<string>();
        private string[] _legacyPptMediaDetails = Array.Empty<string>();

        /// <summary>Gets the detected physical format of the presentation source.</summary>
        public PowerPointFileFormat SourceFormat { get; private set; } = PowerPointFileFormat.Pptx;

        /// <summary>Gets the original legacy source path, or the current Open XML file association.</summary>
        public string? SourcePath => IsLegacyBinaryFormat(SourceFormat)
            ? _legacyPptSourcePath
            : string.IsNullOrWhiteSpace(FilePath) ? null : FilePath;

        /// <summary>Gets diagnostics produced while importing a legacy binary presentation.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> LegacyPptImportDiagnostics => _legacyPptImportDiagnostics;

        internal void MarkLoadedFromLegacyPpt(string? sourcePath, LegacyPptPresentation legacy,
            LegacyPptProjectionMap projectionMap, PowerPointFileFormat sourceFormat) {
            _legacyPptSourcePath = sourcePath;
            _legacyPptImportDiagnostics = legacy.Diagnostics.ToArray();
            _legacyPptPackage = legacy.Package;
            _legacyPptProjectionMap = projectionMap;
            _legacyPptProjectionFingerprint = CreatePackageFingerprint(_document!);
            _legacyPptPreservationFingerprint = LegacyPptProjectionFingerprint.Create(_document!, projectionMap);
            _legacyPptLinkedOleDetails = legacy.LinkedOleObjects.Select(ole =>
                $"Linked OLE {ole.Id}: {ole.ProgId ?? "unknown class"}, {ole.UpdateMode}, {ole.Length} bytes"
                + (ole.WasCompressed ? ", compressed source" : string.Empty))
                .ToArray();
            _legacyPptActiveXDetails = legacy.ActiveXControls.Select(control =>
                $"ActiveX {control.Id}: {control.ProgId ?? "unknown class"}, {control.Length} bytes"
                + (control.WasCompressed ? ", compressed source" : string.Empty))
                .ToArray();
            _legacyPptMediaDetails = legacy.Media.Where(media =>
                    !media.HasProjectableAudio || media.Loop
                    || media.Rewind || media.Narration)
                .Select(media => $"{media.Kind} {media.Id}: "
                    + (media.Path ?? (media.SoundId.HasValue
                        ? $"sound {media.SoundId.Value}"
                        : "native device reference")))
                .ToArray();
            SourceFormat = sourceFormat;
        }

        internal void MarkLoadedFromOpenXml() {
            _legacyPptPackage = null;
            _legacyPptProjectionMap = null;
            _legacyPptProjectionFingerprint = null;
            _legacyPptPreservationFingerprint = null;
            _legacyPptLinkedOleDetails = Array.Empty<string>();
            _legacyPptActiveXDetails = Array.Empty<string>();
            _legacyPptMediaDetails = Array.Empty<string>();
            SourceFormat = PowerPointFileFormat.Pptx;
        }

        internal bool CanPreserveOriginalLegacyPackage => _legacyPptPackage != null
            && _legacyPptProjectionFingerprint != null
            && string.Equals(_legacyPptProjectionFingerprint, CreatePackageFingerprint(_document!),
                StringComparison.Ordinal);

        internal LegacyPptPackage? LegacyPptPackage => _legacyPptPackage;

        internal LegacyPptProjectionMap? LegacyPptProjectionMap => _legacyPptProjectionMap;

        internal IReadOnlyList<string> LegacyPptLinkedOleDetails =>
            _legacyPptLinkedOleDetails;

        internal IReadOnlyList<string> LegacyPptActiveXDetails =>
            _legacyPptActiveXDetails;

        internal IReadOnlyList<string> LegacyPptMediaDetails =>
            _legacyPptMediaDetails;

        internal bool HasOnlyLegacyPptPreservableChanges => _legacyPptProjectionMap != null
            && _legacyPptPreservationFingerprint != null
            && _legacyPptPreservationFingerprint.Matches(_document!, _legacyPptProjectionMap);

        internal bool TryCopyOriginalLegacyPackage(out byte[] bytes) {
            if (CanPreserveOriginalLegacyPackage) {
                bytes = _legacyPptPackage!.CopyOriginalBytes();
                return true;
            }
            bytes = Array.Empty<byte>();
            return false;
        }

        internal bool TryCopyOriginalEncryptedLegacyPackage(
            out byte[] bytes) {
            if (CanPreserveOriginalLegacyPackage
                && _legacyPptPackage!.TryCopyOriginalEncryptedBytes(
                    out bytes)) {
                return true;
            }
            bytes = Array.Empty<byte>();
            return false;
        }

        private void ClearLegacyPptPackageState() {
            _legacyPptPackage = null;
            _legacyPptProjectionMap = null;
            _legacyPptProjectionFingerprint = null;
            _legacyPptPreservationFingerprint = null;
            _legacyPptLinkedOleDetails = Array.Empty<string>();
            _legacyPptActiveXDetails = Array.Empty<string>();
            _legacyPptMediaDetails = Array.Empty<string>();
        }

        internal static bool IsLegacyBinaryFormat(PowerPointFileFormat format) =>
            format == PowerPointFileFormat.Ppt || format == PowerPointFileFormat.Pot || format == PowerPointFileFormat.Pps;
    }
}
