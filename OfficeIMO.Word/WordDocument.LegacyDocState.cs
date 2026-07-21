using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private LegacyDocImportDiagnostic[] _legacyDocImportDiagnostics = Array.Empty<LegacyDocImportDiagnostic>();
        private LegacyDocUnsupportedFeature[] _legacyDocUnsupportedFeatures = Array.Empty<LegacyDocUnsupportedFeature>();
        private LegacyDocPreservedFeature[] _legacyDocPreservedFeatures = Array.Empty<LegacyDocPreservedFeature>();
        private LegacyDocCompoundFeature[] _legacyDocCompoundFeatures = Array.Empty<LegacyDocCompoundFeature>();
        private OfficeCompoundFile? _legacyDocSourceCompoundFile;
        private string? _legacyDocSourcePath;
        private byte[]? _openXmlOriginalPackageBytes;

        /// <summary>
        /// Gets the detected physical format of the document source.
        /// </summary>
        public WordFileFormat SourceFormat { get; private set; } = WordFileFormat.Docx;

        /// <summary>Gets the concrete source format, including modern template and macro variants.</summary>
        public OfficeFormatDescriptor SourceFormatDescriptor => SourceFormat == WordFileFormat.Doc
            ? WordFormatCatalog.GetDescriptor(SourceFormat, _legacyDocSourcePath)
            : WordFormatCatalog.GetDescriptor(_wordprocessingDocument.DocumentType);

        /// <summary>Gets the original legacy source path, or the current Open XML file association.</summary>
        public string? SourcePath => SourceFormat == WordFileFormat.Doc
            ? _legacyDocSourcePath
            : string.IsNullOrWhiteSpace(FilePath) ? null : FilePath;

        /// <summary>
        /// Gets diagnostics produced while importing the legacy `.doc` document.
        /// </summary>
        public IReadOnlyList<LegacyDocImportDiagnostic> LegacyDocImportDiagnostics => _legacyDocImportDiagnostics;

        /// <summary>
        /// Gets unsupported or preserve-only features discovered while importing the legacy `.doc` document.
        /// </summary>
        public IReadOnlyList<LegacyDocUnsupportedFeature> LegacyDocUnsupportedFeatures => _legacyDocUnsupportedFeatures;

        /// <summary>
        /// Gets preserve-only non-compound feature metadata discovered while importing the legacy `.doc` document.
        /// </summary>
        public IReadOnlyList<LegacyDocPreservedFeature> LegacyDocPreservedFeatures => _legacyDocPreservedFeatures;

        /// <summary>
        /// Gets preserve-only compound storage discovered while importing the legacy `.doc` document.
        /// </summary>
        public IReadOnlyList<LegacyDocCompoundFeature> LegacyDocCompoundFeatures => _legacyDocCompoundFeatures;

        /// <summary>Tries to read an original source retained by a compatibility conversion.</summary>
        public bool TryGetCompatibilitySourcePayload(
            out OfficeCompatibilitySourcePayload? payload,
            out string? error) {
            if (SourceFormat == WordFileFormat.Doc) {
                return OfficeCompatibilitySourceCarrier.TryRead(
                    _legacyDocSourceCompoundFile,
                    out payload,
                    out error);
            }

            if (_openXmlOriginalPackageBytes != null) {
                return OfficeCompatibilitySourceCarrier.TryReadPackage(
                    _openXmlOriginalPackageBytes,
                    out payload,
                    out error);
            }

            payload = null;
            error = null;
            return false;
        }

        internal OfficeCompoundFile? LegacyDocSourceCompoundFile => _legacyDocSourceCompoundFile;

        internal void MarkLoadedFromLegacyDoc(string? sourcePath, LegacyDocDocument document, bool attachSourcePathForSave = false) {
            SourceFormat = WordFileFormat.Doc;
            _legacyDocSourcePath = sourcePath;
            _legacyDocImportDiagnostics = document.Diagnostics.ToArray();
            _legacyDocPreservedFeatures = document.PreservedFeatures.ToArray();
            _legacyDocCompoundFeatures = document.CompoundFeatures.ToArray();
            _legacyDocUnsupportedFeatures = document.UnsupportedFeatures.ToArray();
            _legacyDocSourceCompoundFile = document.SourceCompoundFile;
            _openXmlOriginalPackageBytes = null;
            FilePath = attachSourcePathForSave && sourcePath != null
                ? sourcePath
                : null;
        }

        private bool HasLossyLegacyDocImportState() {
            return _legacyDocUnsupportedFeatures.Length > 0
                || _legacyDocPreservedFeatures.Length > 0
                || _legacyDocCompoundFeatures.Length > 0;
        }

        private void EnsureLegacyDocSaveDoesNotDropImportedContent(WordSaveOptions? options) {
            if (SourceFormat != WordFileFormat.Doc
                || !HasLossyLegacyDocImportState()
                || options?.LossPolicy == WordConversionLossPolicy.Allow) {
                return;
            }

            string source = string.IsNullOrWhiteSpace(_legacyDocSourcePath)
                ? "a legacy binary DOC source"
                : $"legacy binary DOC source '{_legacyDocSourcePath}'";
            throw new NotSupportedException($"Saving is blocked because this document was loaded from {source} with unsupported or preserve-only content. Review LegacyDocUnsupportedFeatures, LegacyDocPreservedFeatures, and LegacyDocCompoundFeatures, or set WordSaveOptions.LossPolicy to WordConversionLossPolicy.Allow when that loss is intentional.");
        }
    }
}
