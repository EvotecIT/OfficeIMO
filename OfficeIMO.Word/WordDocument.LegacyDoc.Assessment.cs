using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    /// <summary>Reports whether the current document can be encoded by OfficeIMO's bounded native DOC writer.</summary>
    public sealed class LegacyDocWriteAssessment {
        internal LegacyDocWriteAssessment(bool isSupported, long? encodedByteCount, string diagnosticCode, string message) {
            IsSupported = isSupported;
            EncodedByteCount = encodedByteCount;
            DiagnosticCode = diagnosticCode;
            Message = message;
        }

        /// <summary>Gets whether the current document is inside the native DOC writer's supported subset.</summary>
        public bool IsSupported { get; }

        /// <summary>Gets the encoded size when assessment completed successfully.</summary>
        public long? EncodedByteCount { get; }

        /// <summary>Gets the stable assessment code.</summary>
        public string DiagnosticCode { get; }

        /// <summary>Gets the human-readable assessment detail.</summary>
        public string Message { get; }

        /// <summary>Throws when the document is outside the native DOC writer's supported subset.</summary>
        public LegacyDocWriteAssessment EnsureSupported() {
            if (!IsSupported) throw new NotSupportedException(Message);
            return this;
        }
    }

    public partial class WordDocument {
        /// <summary>
        /// Runs the real native DOC encoder without committing an artifact and returns a structured capability result.
        /// This allocates the candidate DOC bytes so the assessment cannot drift from the writer used by <see cref="Save(Stream, WordFileFormat, WordSaveOptions?)"/>.
        /// </summary>
        public LegacyDocWriteAssessment AssessLegacyDocWrite(WordSaveOptions? options = null) {
            if (_wordprocessingDocument.MainDocumentPart?.Document?.Body == null) {
                return new LegacyDocWriteAssessment(
                    false,
                    null,
                    "LegacyDocWriteUnsupported",
                    "Document couldn't be saved as legacy DOC because the main document part or document body is missing.");
            }
            using var cloneStream = new MemoryStream();
            using (var packageClone = _wordprocessingDocument.Clone(cloneStream, true)) { }
            cloneStream.Position = 0;
            using WordDocument writableClone = Load(cloneStream, new WordLoadOptions {
                MaxInputBytes = Math.Max(WordLoadOptions.DefaultMaxInputBytes, cloneStream.Length)
            });
            CopyLegacyDocAssessmentStateTo(writableClone);
            return writableClone.AssessLegacyDocWriteCore(options);
        }

        private void CopyLegacyDocAssessmentStateTo(WordDocument clone) {
            clone.SourceFormat = SourceFormat;
            clone.AutoUpdateToc = AutoUpdateToc;
            clone._legacyDocImportDiagnostics = _legacyDocImportDiagnostics;
            clone._legacyDocUnsupportedFeatures = _legacyDocUnsupportedFeatures;
            clone._legacyDocPreservedFeatures = _legacyDocPreservedFeatures;
            clone._legacyDocCompoundFeatures = _legacyDocCompoundFeatures;
            clone._legacyDocSourceCompoundFile = _legacyDocSourceCompoundFile;
            clone._legacyDocSourcePath = _legacyDocSourcePath;
            clone._openXmlOriginalPackageBytes = _openXmlOriginalPackageBytes;
        }

        private LegacyDocWriteAssessment AssessLegacyDocWriteCore(WordSaveOptions? options) {
            try {
                byte[] bytes = ToBytes(WordFileFormat.Doc, options);
                return new LegacyDocWriteAssessment(
                    true,
                    bytes.LongLength,
                    "LegacyDocWriteSupported",
                    "The document is inside OfficeIMO's tested native DOC writer subset.");
            } catch (NotSupportedException exception) {
                return new LegacyDocWriteAssessment(
                    false,
                    null,
                    "LegacyDocWriteUnsupported",
                    exception.Message);
            } catch (WordSignatureSavePolicyException exception) {
                return new LegacyDocWriteAssessment(
                    false,
                    null,
                    "LegacyDocWriteUnsupported",
                    exception.Message);
            } catch (InvalidOperationException exception) {
                return new LegacyDocWriteAssessment(
                    false,
                    null,
                    "LegacyDocWriteUnsupported",
                    exception.Message);
            }
        }
    }
}
