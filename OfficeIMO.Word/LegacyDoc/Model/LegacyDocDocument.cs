using OfficeIMO.Shared;
using OfficeIMO.Word.LegacyDoc.Diagnostics;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Neutral legacy binary Word document model for the supported import subset.
    /// </summary>
    public sealed class LegacyDocDocument {
        private readonly List<LegacyDocImportDiagnostic> _diagnostics = new();
        private readonly List<string> _paragraphs = new();
        private readonly List<IReadOnlyList<LegacyDocTextRun>> _paragraphTextRuns = new();
        private readonly List<LegacyDocUnsupportedFeature> _unsupportedFeatures = new();

        private LegacyDocDocument() {
        }

        /// <summary>Gets body text decoded from the Word piece table.</summary>
        public string Text { get; private set; } = string.Empty;

        /// <summary>Gets body paragraphs projected from Word paragraph marks.</summary>
        public IReadOnlyList<string> Paragraphs => _paragraphs;

        internal IReadOnlyList<IReadOnlyList<LegacyDocTextRun>> ParagraphTextRuns => _paragraphTextRuns;

        internal LegacyDocDocumentProperties DocumentProperties { get; } = new();

        /// <summary>Gets diagnostics produced while reading the legacy document.</summary>
        public IReadOnlyList<LegacyDocImportDiagnostic> Diagnostics => _diagnostics;

        /// <summary>Gets unsupported or preserve-only features discovered while reading the legacy document.</summary>
        public IReadOnlyList<LegacyDocUnsupportedFeature> UnsupportedFeatures => _unsupportedFeatures;

        /// <summary>
        /// Loads a legacy DOC model from a file path.
        /// </summary>
        public static LegacyDocDocument Load(string path, LegacyDocImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            return Load(File.ReadAllBytes(path), options);
        }

        /// <summary>
        /// Loads a legacy DOC model from a stream.
        /// </summary>
        public static LegacyDocDocument Load(Stream stream, LegacyDocImportOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            using var buffer = new MemoryStream();
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }
            stream.CopyTo(buffer);
            return Load(buffer.ToArray(), options);
        }

        /// <summary>
        /// Loads a legacy DOC model from compound document bytes.
        /// </summary>
        public static LegacyDocDocument Load(byte[] bytes, LegacyDocImportOptions? options = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            options ??= new LegacyDocImportOptions();

            var document = new LegacyDocDocument();
            if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compoundFile, out string? compoundError)) {
                document.AddError("DOC-COMPOUND-INVALID", compoundError ?? "The OLE compound document could not be read.");
                return document;
            }

            document.LoadFromCompound(compoundFile!, options);
            return document;
        }

        /// <summary>
        /// Creates a compact report from the decoded model and diagnostics.
        /// </summary>
        public LegacyDocImportReport CreateImportReport() {
            return new LegacyDocImportReport(this);
        }

        private void LoadFromCompound(OfficeCompoundFile compoundFile, LegacyDocImportOptions options) {
            if (!compoundFile.Streams.TryGetValue("WordDocument", out byte[]? wordDocumentStream)) {
                AddError("DOC-WORDDOCUMENT-MISSING", "The compound document does not contain a WordDocument stream.");
                return;
            }

            if (wordDocumentStream.Length > options.MaxWordDocumentStreamBytes) {
                AddError("DOC-WORDDOCUMENT-TOO-LARGE", $"The WordDocument stream is {wordDocumentStream.Length} bytes, which exceeds the configured limit of {options.MaxWordDocumentStreamBytes} bytes.");
                return;
            }

            LegacyDocFib fib;
            if (!LegacyDocFib.TryRead(wordDocumentStream, out fib, out string? fibError)) {
                AddError("DOC-FIB-INVALID", fibError ?? "The WordDocument stream does not contain a supported File Information Block.");
                return;
            }

            if (fib.IsEncrypted) {
                AddError("DOC-ENCRYPTED", "Password-protected binary .doc files are detected but are not imported by the dependency-free reader.");
                return;
            }

            if (options.ReportUnsupportedFeatures) {
                AddKnownUnsupportedFeatureDiagnostics(compoundFile);
            }

            LegacyDocOleDocumentPropertyReader.AddDocumentProperties(compoundFile, this, options);

            string tableStreamName = fib.UsesOneTableStream ? "1Table" : "0Table";
            if (!compoundFile.Streams.TryGetValue(tableStreamName, out byte[]? tableStream)) {
                string alternateName = fib.UsesOneTableStream ? "0Table" : "1Table";
                if (compoundFile.Streams.TryGetValue(alternateName, out tableStream)) {
                    AddWarning("DOC-TABLE-STREAM-FALLBACK", $"The FIB requested {tableStreamName}, but only {alternateName} was present. The available table stream was used.");
                } else {
                    AddError("DOC-TABLE-STREAM-MISSING", $"The compound document does not contain the {tableStreamName} table stream.");
                    return;
                }
            }

            if (!LegacyDocPieceTable.TryRead(wordDocumentStream, tableStream, fib, out LegacyDocTextContent textContent, out string? textError)) {
                AddError("DOC-PIECE-TABLE-INVALID", textError ?? "The legacy DOC piece table could not be decoded.");
                return;
            }

            IReadOnlyList<string> fontFamilies = LegacyDocFontTableReader.ReadFontFamilies(tableStream, fib, out string? fontTableWarning);
            if (fontTableWarning != null) {
                AddWarning("DOC-FONT-TABLE-INVALID", fontTableWarning);
            }

            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges = LegacyDocCharacterFormattingReader.ReadCharacterFormatting(wordDocumentStream, tableStream, fib, fontFamilies, out string? formattingWarning);
            if (formattingWarning != null) {
                AddWarning("DOC-CHPX-INVALID", formattingWarning);
            }

            Text = BuildFormattedParagraphs(textContent.Characters, formattingRanges);
            foreach (IReadOnlyList<LegacyDocTextRun> paragraphRuns in _paragraphTextRuns) {
                _paragraphs.Add(string.Concat(paragraphRuns.Select(run => run.Text)));
            }
        }

        private void AddKnownUnsupportedFeatureDiagnostics(OfficeCompoundFile compoundFile) {
            OfficeCompoundFileEntry? macroEntry = compoundFile.Entries.FirstOrDefault(entry => entry.Path.StartsWith("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                || entry.Path.StartsWith("Macros", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0);
            if (macroEntry != null) {
                AddUnsupportedFeature(
                    new LegacyDocUnsupportedFeature(
                        LegacyDocUnsupportedFeatureKind.VbaProject,
                        "DOC-MACROS-PRESENT",
                        "The legacy DOC contains VBA project storage. Macros are preserved in the source file but are not projected into the OfficeIMO document.",
                        macroEntry.Path,
                        "Compound:VbaProjectStorage"));
            }

            OfficeCompoundFileEntry? oleObjectEntry = compoundFile.Entries.FirstOrDefault(entry => entry.Path.IndexOf("ObjectPool", StringComparison.OrdinalIgnoreCase) >= 0);
            if (oleObjectEntry != null) {
                AddUnsupportedFeature(
                    new LegacyDocUnsupportedFeature(
                        LegacyDocUnsupportedFeatureKind.OleObject,
                        "DOC-OLE-OBJECTS-PRESENT",
                        "The legacy DOC contains embedded OLE object storage. Embedded objects are preserved in the source file but are not projected into the OfficeIMO document.",
                        oleObjectEntry.Path,
                        "Compound:OleObjectStorage"));
            }
        }

        private string BuildFormattedParagraphs(IReadOnlyList<LegacyDocTextCharacter> characters, IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges) {
            var bodyText = new System.Text.StringBuilder(characters.Count);
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            bool hasCurrentRun = false;

            foreach (LegacyDocTextCharacter textCharacter in characters) {
                char? normalized = NormalizeBodyCharacter(textCharacter.Character);
                if (normalized == null) {
                    continue;
                }

                LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset);
                if (normalized.Value == '\r') {
                    FlushRun();
                    _paragraphTextRuns.Add(currentRuns.ToArray());
                    currentRuns.Clear();
                    hasCurrentRun = false;
                    bodyText.Append('\r');
                    continue;
                }

                if (!hasCurrentRun || !format.Equals(currentFormat)) {
                    FlushRun();
                    currentFormat = format;
                    hasCurrentRun = true;
                }

                runText.Append(normalized.Value);
                bodyText.Append(normalized.Value);
            }

            FlushRun();
            if (currentRuns.Count > 0) {
                _paragraphTextRuns.Add(currentRuns.ToArray());
            }

            return bodyText.ToString();

            void FlushRun() {
                if (runText.Length == 0) {
                    return;
                }

                currentRuns.Add(new LegacyDocTextRun(
                    runText.ToString(),
                    currentFormat.Bold,
                    currentFormat.Italic,
                    currentFormat.Underline,
                    currentFormat.FontSizeHalfPoints,
                    currentFormat.ColorHex,
                    currentFormat.FontFamily));
                runText.Clear();
            }
        }

        private static char? NormalizeBodyCharacter(char character) {
            switch (character) {
                case '\0':
                case '\a':
                    return null;
                case '\v':
                case '\f':
                case '\n':
                    return '\r';
                default:
                    if (!char.IsControl(character) || character == '\t' || character == '\r' || character == '\n') {
                        return character;
                    }

                    return null;
            }
        }

        private static LegacyDocCharacterFormat GetFormatForFileOffset(IReadOnlyList<LegacyDocCharacterFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocCharacterFormat.Default;
        }

        internal void AddInfo(string code, string message) {
            _diagnostics.Add(new LegacyDocImportDiagnostic(code, LegacyDocDiagnosticSeverity.Info, message));
        }

        private void AddError(string code, string message) {
            _diagnostics.Add(new LegacyDocImportDiagnostic(code, LegacyDocDiagnosticSeverity.Error, message));
        }

        internal void AddWarning(string code, string message) {
            _diagnostics.Add(new LegacyDocImportDiagnostic(code, LegacyDocDiagnosticSeverity.Warning, message));
        }

        private void AddUnsupportedFeature(LegacyDocUnsupportedFeature feature) {
            _unsupportedFeatures.Add(feature);
            AddWarning(feature.Code, feature.Description);
        }
    }
}
