using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.Drawing;
using System.Security.Cryptography;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        // Field initializers would run in every constructor, the SDK one included, so these start null.
        private LegacyPptImportDiagnostic[]? _legacyPptImportDiagnostics;
        private string? _legacyPptSourcePath;
        private LegacyPptPackage? _legacyPptPackage;
        private LegacyPptProjectionMap? _legacyPptProjectionMap;
        private string? _legacyPptProjectionFingerprint;
        private LegacyPptProjectionFingerprint? _legacyPptPreservationFingerprint;
        private string[]? _legacyPptLinkedOleDetails;
        private string[]? _legacyPptActiveXDetails;
        private string[]? _legacyPptMediaDetails;
        private bool _legacyPptHasVbaContent;
        private bool _legacyPptHadProjectedVbaContent;
        private byte[]? _legacyPptProjectedVbaDigest;
        private bool _legacyPptHasEmbeddedOleContent;
        private bool _legacyPptHasLinkedOleContent;
        private bool _legacyPptHasActiveXContent;
        private bool _legacyPptHasExternalHyperlinkContent;
        private bool _legacyPptHasExternalMediaContent;
        private bool _legacyPptHasRunProgramContent;
        private bool _legacyPptHadProjectedRunProgramContent;
        private byte[]? _openXmlOriginalPackageBytes;

        /// <summary>Gets the detected physical format of the presentation source.</summary>
        public PowerPointFileFormat SourceFormat { get; private set; }

        /// <summary>Gets the concrete source format, including modern template, slideshow, add-in, and macro variants.</summary>
        public OfficeFormatDescriptor SourceFormatDescriptor => IsLegacyBinaryFormat(SourceFormat)
            ? PowerPointFormatCatalog.GetDescriptor(SourceFormat, _legacyPptSourcePath)
            : OpenXmlFormatDescriptor
                ?? PowerPointFormatCatalog.GetDescriptor(SourceFormat, FilePath);

        /// <summary>Gets the original legacy source path, or the current Open XML file association.</summary>
        public string? SourcePath => IsLegacyBinaryFormat(SourceFormat)
            ? _legacyPptSourcePath
            : string.IsNullOrWhiteSpace(FilePath) ? null : FilePath;

        /// <summary>Gets diagnostics produced while importing a legacy binary presentation.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> LegacyPptImportDiagnostics =>
            _legacyPptImportDiagnostics ?? Array.Empty<LegacyPptImportDiagnostic>();

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
                return OfficeCompatibilitySourceCarrier.TryReadPackage(
                    CapturePackageBytes(),
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

        internal void MarkLoadedFromLegacyPpt(string? sourcePath, LegacyPptDocument legacy,
            LegacyPptProjectionMap projectionMap, PowerPointFileFormat sourceFormat) {
            _legacyPptSourcePath = sourcePath;
            _legacyPptImportDiagnostics = legacy.Diagnostics.ToArray();
            _sourceLoadDiagnostics = legacy.CreateImportReport().Diagnostics.Select(d => new OfficeDiagnostic(d.Code,
                legacy.Diagnostics.FirstOrDefault(item => item.Code == d.Code)?.Severity switch {
                    LegacyPptDiagnosticSeverity.Information => OfficeDiagnosticSeverity.Information,
                    LegacyPptDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                    _ => OfficeDiagnosticSeverity.Warning
                }, d.Message, d.Location, d.Format, d.LossKind)).ToArray();
            _legacyPptPackage = legacy.Package;
            _openXmlOriginalPackageBytes = null;
            _legacyPptProjectionMap = projectionMap;
            // Both fingerprints read a copy of the package, so the presentation stays in whichever model holds it.
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(this);
            _legacyPptProjectionFingerprint = LegacyPptPackageFingerprint.Create(view);
            _legacyPptPreservationFingerprint = LegacyPptProjectionFingerprint.Create(view, projectionMap);
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
            _legacyPptHadProjectedRunProgramContent =
                EnumerateReferencedHyperlinks(view).Any(IsRunProgramHyperlink);
            SourceFormat = sourceFormat;
        }

        internal void MarkLoadedFromOpenXml(byte[] originalPackageBytes) {
            ArgumentNullException.ThrowIfNull(originalPackageBytes);
            MarkLoadedFromOpenXmlCore(OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(originalPackageBytes)
                ? (byte[])originalPackageBytes.Clone()
                : null);
        }

        /// <summary>Records an Open XML source read in place from the caller's stream.</summary>
        internal void MarkLoadedFromOpenXml(Stream originalPackage, CancellationToken cancellationToken) {
            ArgumentNullException.ThrowIfNull(originalPackage);
            MarkLoadedFromOpenXmlCore(OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(originalPackage)
                ? OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(originalPackage, cancellationToken)
                : null);
        }

        private void MarkLoadedFromOpenXmlCore(byte[]? carrierPackageBytes) {
            _legacyPptPackage = null;
            _openXmlOriginalPackageBytes = carrierPackageBytes;
            _legacyPptProjectionMap = null;
            _legacyPptProjectionFingerprint = null;
            _legacyPptPreservationFingerprint = null;
            _legacyPptLinkedOleDetails = Array.Empty<string>();
            _legacyPptActiveXDetails = Array.Empty<string>();
            _legacyPptMediaDetails = Array.Empty<string>();
            ClearLegacyPptSecurityState();
            SourceFormat = IsMacroEnabledOpenXmlSource
                ? PowerPointFileFormat.Pptm
                : PowerPointFileFormat.Pptx;
        }

        internal bool CanPreserveOriginalLegacyPackage {
            get {
                if (_legacyPptPackage == null || _legacyPptProjectionFingerprint == null) return false;
                using LegacyPptPackageView view = LegacyPptPackageView.Capture(this);
                return CanPreserveOriginalLegacyPackageIn(view);
            }
        }

        /// <summary>Whether the package read through <paramref name="view"/> is still the one the import projected.</summary>
        internal bool CanPreserveOriginalLegacyPackageIn(LegacyPptPackageView view) =>
            _legacyPptPackage != null
            && _legacyPptProjectionFingerprint != null
            && string.Equals(_legacyPptProjectionFingerprint, LegacyPptPackageFingerprint.Create(view),
                StringComparison.Ordinal);

        internal LegacyPptPackage? LegacyPptPackage => _legacyPptPackage;

        internal LegacyPptProjectionMap? LegacyPptProjectionMap => _legacyPptProjectionMap;

        internal IReadOnlyList<string> LegacyPptLinkedOleDetails =>
            _legacyPptLinkedOleDetails ?? Array.Empty<string>();

        internal IReadOnlyList<string> LegacyPptActiveXDetails =>
            _legacyPptActiveXDetails ?? Array.Empty<string>();

        internal IReadOnlyList<string> LegacyPptMediaDetails =>
            _legacyPptMediaDetails ?? Array.Empty<string>();

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
        internal bool LegacyPptWillPreserveExternalHyperlinkContent {
            get {
                if (!_legacyPptHasExternalHyperlinkContent) return false;
                using LegacyPptPackageView view = LegacyPptPackageView.Capture(this);
                return EnumerateReferencedHyperlinks(view).Any(item =>
                    IsExternalHyperlink(item.Part, item.Hyperlink));
            }
        }
        internal bool LegacyPptHasExternalMediaContent =>
            _legacyPptHasExternalMediaContent;
        internal bool LegacyPptWillPreserveRunProgramContent {
            get {
                if (!_legacyPptHasRunProgramContent) return false;
                if (!_legacyPptHadProjectedRunProgramContent) return true;
                using LegacyPptPackageView view = LegacyPptPackageView.Capture(this);
                return EnumerateReferencedHyperlinks(view).Any(IsRunProgramHyperlink);
            }
        }

        internal bool HasOnlyLegacyPptPreservableChanges {
            get {
                if (_legacyPptProjectionMap == null || _legacyPptPreservationFingerprint == null) return false;
                using LegacyPptPackageView view = LegacyPptPackageView.Capture(this);
                return HasOnlyLegacyPptPreservableChangesIn(view);
            }
        }

        /// <summary>Whether the package read through <paramref name="view"/> holds only edits the preserving writer encodes.</summary>
        internal bool HasOnlyLegacyPptPreservableChangesIn(LegacyPptPackageView view) =>
            _legacyPptProjectionMap != null
            && _legacyPptPreservationFingerprint != null
            && _legacyPptPreservationFingerprint.Matches(view, _legacyPptProjectionMap);

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
            _legacyPptHadProjectedRunProgramContent = false;
        }

        private bool IsProjectedVbaContentUnchanged() {
            if (_legacyPptProjectedVbaDigest == null) return true;
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(this);
            LegacyPptPartView part = view.Presentation.RelatedPart(
                LegacyPptPackageView.VbaProjectRelationshipType);
            if (part.IsNone) return false;
            try {
                return ComputeSha256(part.GetBytes())
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

        /// <summary>
        /// The hyperlinks with a relationship id on the slides, their notes, the masters, their layouts and the notes and
        /// handout masters, with the part that holds each.
        /// </summary>
        private static IEnumerable<(LegacyPptPartView Part, OfficeXmlNode Hyperlink)>
            EnumerateReferencedHyperlinks(LegacyPptPackageView view) {
            LegacyPptPartView presentation = view.Presentation;
            foreach (LegacyPptPartView slidePart in presentation.RelatedParts(
                         LegacyPptPackageView.SlideRelationshipType)) {
                foreach (var item in EnumerateReferencedHyperlinks(slidePart)) yield return item;
                LegacyPptPartView notesPart = slidePart.NotesSlide;
                if (notesPart.IsNone) continue;
                foreach (var item in EnumerateReferencedHyperlinks(notesPart)) yield return item;
            }
            foreach (LegacyPptPartView masterPart in presentation.RelatedParts(
                         LegacyPptPackageView.SlideMasterRelationshipType)) {
                foreach (var item in EnumerateReferencedHyperlinks(masterPart)) yield return item;
                foreach (LegacyPptPartView layoutPart in masterPart.RelatedParts(
                             LegacyPptPackageView.SlideLayoutRelationshipType)) {
                    foreach (var item in EnumerateReferencedHyperlinks(layoutPart)) yield return item;
                }
            }
            foreach (var item in EnumerateReferencedHyperlinks(view.NotesMaster)) yield return item;
            foreach (var item in EnumerateReferencedHyperlinks(view.HandoutMaster)) yield return item;
        }

        private static IEnumerable<(LegacyPptPartView Part, OfficeXmlNode Hyperlink)>
            EnumerateReferencedHyperlinks(LegacyPptPartView part) {
            if (part.IsNone) yield break;
            OfficeXmlNode root = part.Root;
            if (root.IsNone) yield break;
            foreach (string name in new[] { "hlinkClick", "hlinkHover", "hlinkMouseOver" }) {
                foreach (OfficeXmlNode item in root.Descendants(LegacyPptXml.A, name)) {
                    if (!string.IsNullOrEmpty(item.Attribute(LegacyPptXml.R, "id")))
                        yield return (part, item);
                }
            }
        }

        private static bool IsRunProgramHyperlink((LegacyPptPartView Part, OfficeXmlNode Hyperlink) item) =>
            string.Equals(item.Hyperlink.Attribute("action"), "ppaction://program",
                StringComparison.OrdinalIgnoreCase);

        private static bool IsExternalHyperlink(LegacyPptPartView part,
            OfficeXmlNode hyperlink) {
            string? relationshipId = hyperlink.Attribute(LegacyPptXml.R, "id");
            return !string.IsNullOrEmpty(relationshipId)
                && part.Relationships.Any(relationship =>
                    relationship.IsExternal
                    && relationship.IsType(LegacyPptPackageView.HyperlinkRelationshipType)
                    && string.Equals(relationship.Id, relationshipId,
                        StringComparison.Ordinal));
        }

        private static byte[] ComputeSha256(byte[] bytes) {
            using SHA256 sha256 = SHA256.Create();
            return sha256.ComputeHash(bytes);
        }

        /// <summary>
        /// Opens a separate read-only presentation over package bytes, for the binary writer's checks that read engine
        /// shapes, so they don't open the SDK on the presentation being imported.
        /// </summary>
        internal static PowerPointDocument OpenLegacyPptProjectionReader(byte[] packageBytes) {
            ArgumentNullException.ThrowIfNull(packageBytes);
            return LoadPackage(packageBytes, filePath: null, sourceStream: null,
                new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    PersistenceMode = DocumentPersistenceMode.Explicit,
                    MaxInputBytes = null
                },
                captureSourcePackageSecurity: false);
        }

        internal static bool IsLegacyBinaryFormat(PowerPointFileFormat format) =>
            format == PowerPointFileFormat.Ppt || format == PowerPointFileFormat.Pot || format == PowerPointFileFormat.Pps;
    }
}
