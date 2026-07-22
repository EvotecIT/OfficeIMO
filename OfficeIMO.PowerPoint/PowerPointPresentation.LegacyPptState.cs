using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System.Security.Cryptography;
using A = DocumentFormat.OpenXml.Drawing;

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
        private bool _legacyPptHasVbaContent;
        private bool _legacyPptHadProjectedVbaContent;
        private byte[]? _legacyPptProjectedVbaDigest;
        private bool _legacyPptHasEmbeddedOleContent;
        private bool _legacyPptHasLinkedOleContent;
        private bool _legacyPptHasActiveXContent;
        private bool _legacyPptHasExternalHyperlinkContent;
        private bool _legacyPptHasExternalMediaContent;
        private bool _legacyPptHasRunProgramContent;
        private byte[]? _openXmlOriginalPackageBytes;

        /// <summary>Gets the detected physical format of the presentation source.</summary>
        public PowerPointFileFormat SourceFormat { get; private set; } = PowerPointFileFormat.Pptx;

        /// <summary>Gets the concrete source format, including modern template, slideshow, add-in, and macro variants.</summary>
        public OfficeFormatDescriptor SourceFormatDescriptor => IsLegacyBinaryFormat(SourceFormat)
            ? PowerPointFormatCatalog.GetDescriptor(SourceFormat, _legacyPptSourcePath)
            : _document != null
                ? PowerPointFormatCatalog.GetDescriptor(_document.DocumentType)
                : PowerPointFormatCatalog.GetDescriptor(SourceFormat, FilePath);

        /// <summary>Gets the original legacy source path, or the current Open XML file association.</summary>
        public string? SourcePath => IsLegacyBinaryFormat(SourceFormat)
            ? _legacyPptSourcePath
            : string.IsNullOrWhiteSpace(FilePath) ? null : FilePath;

        /// <summary>Gets diagnostics produced while importing a legacy binary presentation.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> LegacyPptImportDiagnostics => _legacyPptImportDiagnostics;

        /// <summary>Tries to read an original source retained by a compatibility conversion.</summary>
        public bool TryGetCompatibilitySourcePayload(
            out OfficeCompatibilitySourcePayload? payload,
            out string? error) {
            if (IsLegacyBinaryFormat(SourceFormat)) {
                return OfficeCompatibilitySourceCarrier.TryRead(
                    _legacyPptPackage?.CompoundFile,
                    out payload,
                    out error);
            }

            if (_openXmlOriginalPackageBytes != null) {
                return OfficeCompatibilitySourceCarrier.TryReadPackage(
                    _openXmlOriginalPackageBytes,
                    out payload,
                    out error);
            }

            try {
                using var snapshot = new MemoryStream();
                using (_document!.Clone(snapshot)) { }
                return OfficeCompatibilitySourceCarrier.TryReadPackage(
                    snapshot.ToArray(),
                    out payload,
                    out error);
            } catch (Exception exception) when (
                exception is IOException
                || exception is InvalidDataException
                || exception is NotSupportedException) {
                payload = null;
                error = exception.Message;
                return false;
            }
        }

        internal void MarkLoadedFromLegacyPpt(string? sourcePath, LegacyPptPresentation legacy,
            LegacyPptProjectionMap projectionMap, PowerPointFileFormat sourceFormat) {
            _legacyPptSourcePath = sourcePath;
            _legacyPptImportDiagnostics = legacy.Diagnostics.ToArray();
            _legacyPptPackage = legacy.Package;
            _openXmlOriginalPackageBytes = null;
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
            _legacyPptHasVbaContent = legacy.HasVbaContent;
            _legacyPptHadProjectedVbaContent = legacy.VbaProject != null;
            _legacyPptProjectedVbaDigest = legacy.VbaProject == null
                ? null
                : ComputeSha256(legacy.VbaProject.GetBytes());
            _legacyPptHasEmbeddedOleContent = legacy.HasEmbeddedOleContent;
            _legacyPptHasLinkedOleContent = legacy.HasLinkedOleContent;
            _legacyPptHasActiveXContent = legacy.HasActiveXContent;
            _legacyPptHasExternalHyperlinkContent =
                legacy.HasExternalHyperlinkContent;
            _legacyPptHasExternalMediaContent =
                legacy.HasExternalMediaContent;
            _legacyPptHasRunProgramContent = legacy.HasRunProgramContent;
            SourceFormat = sourceFormat;
        }

        internal void MarkLoadedFromOpenXml(byte[] originalPackageBytes) {
            if (originalPackageBytes == null) throw new ArgumentNullException(nameof(originalPackageBytes));
            _legacyPptPackage = null;
            _openXmlOriginalPackageBytes = OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(originalPackageBytes)
                ? (byte[])originalPackageBytes.Clone()
                : null;
            _legacyPptProjectionMap = null;
            _legacyPptProjectionFingerprint = null;
            _legacyPptPreservationFingerprint = null;
            _legacyPptLinkedOleDetails = Array.Empty<string>();
            _legacyPptActiveXDetails = Array.Empty<string>();
            _legacyPptMediaDetails = Array.Empty<string>();
            ClearLegacyPptSecurityState();
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

        internal bool LegacyPptWillPreserveVbaContent =>
            _legacyPptHasVbaContent
            && (!_legacyPptHadProjectedVbaContent
                || IsProjectedVbaContentUnchanged());
        internal bool LegacyPptHasEmbeddedOleContent =>
            _legacyPptHasEmbeddedOleContent;
        internal bool LegacyPptHasLinkedOleContent =>
            _legacyPptHasLinkedOleContent;
        internal bool LegacyPptHasActiveXContent =>
            _legacyPptHasActiveXContent;
        internal bool LegacyPptWillPreserveExternalHyperlinkContent =>
            _legacyPptHasExternalHyperlinkContent
            && EnumerateReferencedHyperlinks().Any(item =>
                !string.Equals(item.Action?.Value, "ppaction://program",
                    StringComparison.OrdinalIgnoreCase));
        internal bool LegacyPptHasExternalMediaContent =>
            _legacyPptHasExternalMediaContent;
        internal bool LegacyPptWillPreserveRunProgramContent =>
            _legacyPptHasRunProgramContent
            && EnumerateReferencedHyperlinks().Any(item =>
                string.Equals(item.Action?.Value, "ppaction://program",
                    StringComparison.OrdinalIgnoreCase));

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
            _openXmlOriginalPackageBytes = null;
            _legacyPptProjectionMap = null;
            _legacyPptProjectionFingerprint = null;
            _legacyPptPreservationFingerprint = null;
            _legacyPptLinkedOleDetails = Array.Empty<string>();
            _legacyPptActiveXDetails = Array.Empty<string>();
            _legacyPptMediaDetails = Array.Empty<string>();
            ClearLegacyPptSecurityState();
        }

        private void ClearLegacyPptSecurityState() {
            _legacyPptHasVbaContent = false;
            _legacyPptHadProjectedVbaContent = false;
            _legacyPptProjectedVbaDigest = null;
            _legacyPptHasEmbeddedOleContent = false;
            _legacyPptHasLinkedOleContent = false;
            _legacyPptHasActiveXContent = false;
            _legacyPptHasExternalHyperlinkContent = false;
            _legacyPptHasExternalMediaContent = false;
            _legacyPptHasRunProgramContent = false;
        }

        private bool IsProjectedVbaContentUnchanged() {
            if (_legacyPptProjectedVbaDigest == null) return true;
            VbaProjectPart? part = _presentationPart.VbaProjectPart;
            if (part == null) return false;
            try {
                using Stream stream = part.GetStream(FileMode.Open,
                    FileAccess.Read);
                using SHA256 sha256 = SHA256.Create();
                return sha256.ComputeHash(stream)
                    .SequenceEqual(_legacyPptProjectedVbaDigest);
            } catch (Exception exception) when (
                exception is IOException
                || exception is UnauthorizedAccessException
                || exception is InvalidDataException) {
                // A separate package-part finding will reject unreadable VBA.
                // Remain conservative if that finding is bypassed.
                return true;
            }
        }

        private IEnumerable<A.HyperlinkType> EnumerateReferencedHyperlinks() {
            foreach (SlidePart slidePart in _presentationPart.SlideParts) {
                if (slidePart.Slide != null) {
                    foreach (A.HyperlinkType item in EnumerateReferencedHyperlinks(
                                 slidePart.Slide)) {
                        yield return item;
                    }
                }
                if (slidePart.NotesSlidePart?.NotesSlide != null) {
                    foreach (A.HyperlinkType item in EnumerateReferencedHyperlinks(
                                 slidePart.NotesSlidePart.NotesSlide)) {
                        yield return item;
                    }
                }
            }
            foreach (SlideMasterPart masterPart in _presentationPart.SlideMasterParts) {
                if (masterPart.SlideMaster != null) {
                    foreach (A.HyperlinkType item in EnumerateReferencedHyperlinks(
                                 masterPart.SlideMaster)) {
                        yield return item;
                    }
                }
                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts) {
                    if (layoutPart.SlideLayout == null) continue;
                    foreach (A.HyperlinkType item in EnumerateReferencedHyperlinks(
                                 layoutPart.SlideLayout)) {
                        yield return item;
                    }
                }
            }
            if (_presentationPart.NotesMasterPart?.NotesMaster != null) {
                foreach (A.HyperlinkType item in EnumerateReferencedHyperlinks(
                             _presentationPart.NotesMasterPart.NotesMaster)) {
                    yield return item;
                }
            }
            if (_presentationPart.HandoutMasterPart?.HandoutMaster != null) {
                foreach (A.HyperlinkType item in EnumerateReferencedHyperlinks(
                             _presentationPart.HandoutMasterPart.HandoutMaster)) {
                    yield return item;
                }
            }
        }

        private static IEnumerable<A.HyperlinkType> EnumerateReferencedHyperlinks(
            DocumentFormat.OpenXml.OpenXmlPartRootElement root) {
            foreach (A.HyperlinkOnClick item in root
                         .Descendants<A.HyperlinkOnClick>()) {
                    if (!string.IsNullOrEmpty(item.Id?.Value)) yield return item;
            }
            foreach (A.HyperlinkOnHover item in root
                         .Descendants<A.HyperlinkOnHover>()) {
                    if (!string.IsNullOrEmpty(item.Id?.Value)) yield return item;
            }
            foreach (A.HyperlinkOnMouseOver item in root
                         .Descendants<A.HyperlinkOnMouseOver>()) {
                    if (!string.IsNullOrEmpty(item.Id?.Value)) yield return item;
            }
        }

        private static byte[] ComputeSha256(byte[] bytes) {
            using SHA256 sha256 = SHA256.Create();
            return sha256.ComputeHash(bytes);
        }

        internal static bool IsLegacyBinaryFormat(PowerPointFileFormat format) =>
            format == PowerPointFileFormat.Ppt || format == PowerPointFileFormat.Pot || format == PowerPointFileFormat.Pps;
    }
}
