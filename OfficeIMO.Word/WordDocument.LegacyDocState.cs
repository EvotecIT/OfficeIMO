using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private LegacyDocImportDiagnostic[] _legacyDocImportDiagnostics = Array.Empty<LegacyDocImportDiagnostic>();
        private LegacyDocUnsupportedFeature[] _legacyDocUnsupportedFeatures = Array.Empty<LegacyDocUnsupportedFeature>();
        private LegacyDocPreservedFeature[] _legacyDocPreservedFeatures = Array.Empty<LegacyDocPreservedFeature>();
        private LegacyDocCompoundFeature[] _legacyDocCompoundFeatures = Array.Empty<LegacyDocCompoundFeature>();
        private OfficeCompoundFile? _legacyDocSourceCompoundFile;
        private string? _legacyDocSourcePath;

        /// <summary>
        /// Gets the detected physical format of the document source.
        /// </summary>
        public WordFileFormat SourceFormat { get; private set; } = WordFileFormat.Docx;

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

        internal OfficeCompoundFile? LegacyDocSourceCompoundFile => _legacyDocSourceCompoundFile;

        internal void MarkLoadedFromLegacyDoc(string? sourcePath, LegacyDocDocument document, bool attachSourcePathForSave = false) {
            SourceFormat = WordFileFormat.Doc;
            _legacyDocSourcePath = sourcePath;
            _legacyDocImportDiagnostics = document.Diagnostics.ToArray();
            _legacyDocPreservedFeatures = document.PreservedFeatures.ToArray();
            _legacyDocCompoundFeatures = document.CompoundFeatures.ToArray();
            _legacyDocUnsupportedFeatures = document.UnsupportedFeatures.ToArray();
            _legacyDocSourceCompoundFile = document.SourceCompoundFile;
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
