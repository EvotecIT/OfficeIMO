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
        private readonly List<LegacyDocParagraphFormat> _paragraphFormats = new();
        private readonly List<LegacyDocBodyBlock> _bodyBlocks = new();
        private readonly List<LegacyDocUnsupportedFeature> _unsupportedFeatures = new();

        private LegacyDocDocument() {
        }

        /// <summary>Gets body text decoded from the Word piece table.</summary>
        public string Text { get; private set; } = string.Empty;

        /// <summary>Gets body paragraphs projected from Word paragraph marks.</summary>
        public IReadOnlyList<string> Paragraphs => _paragraphs;

        internal IReadOnlyList<IReadOnlyList<LegacyDocTextRun>> ParagraphTextRuns => _paragraphTextRuns;

        internal IReadOnlyList<LegacyDocParagraphFormat> ParagraphFormats => _paragraphFormats;

        internal IReadOnlyList<LegacyDocBodyBlock> BodyBlocks => _bodyBlocks;

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

            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges = LegacyDocParagraphFormattingReader.ReadParagraphFormatting(wordDocumentStream, tableStream, fib, out string? paragraphFormattingWarning);
            if (paragraphFormattingWarning != null) {
                AddWarning("DOC-PAPX-INVALID", paragraphFormattingWarning);
            }

            Text = BuildFormattedParagraphs(textContent.Characters, formattingRanges, paragraphFormattingRanges);
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

        private string BuildFormattedParagraphs(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges) {
            var bodyText = new System.Text.StringBuilder(characters.Count);
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            bool hasCurrentRun = false;
            bool inTable = false;
            bool justClosedCell = false;
            var tableRows = new List<LegacyDocTableRow>();
            var currentTableRow = new List<LegacyDocTableCell>();

            foreach (LegacyDocTextCharacter textCharacter in characters) {
                if (textCharacter.Character == '\a') {
                    AddCurrentTextAsTableCell();
                    continue;
                }

                char? normalized = NormalizeBodyCharacter(textCharacter.Character);
                if (normalized == null) {
                    continue;
                }

                LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset);
                if (normalized.Value == '\r') {
                    if (inTable) {
                        FlushTable();
                    } else {
                        AddCurrentTextAsParagraph(GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset));
                    }

                    bodyText.Append('\r');
                    continue;
                }

                AppendRunCharacter(normalized.Value, format);
                bodyText.Append(normalized.Value);
            }

            if (inTable) {
                FlushTable();
            } else if (currentRuns.Count > 0 || runText.Length > 0) {
                AddCurrentTextAsParagraph(LegacyDocParagraphFormat.Default);
            }

            return bodyText.ToString();

            void AppendRunCharacter(char character, LegacyDocCharacterFormat format) {
                if (!hasCurrentRun || !format.Equals(currentFormat)) {
                    FlushRun();
                    currentFormat = format;
                    hasCurrentRun = true;
                }

                runText.Append(character);
                if (inTable) {
                    justClosedCell = false;
                }
            }

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

            void AddCurrentTextAsParagraph(LegacyDocParagraphFormat paragraphFormat) {
                FlushRun();
                IReadOnlyList<LegacyDocTextRun> runs = currentRuns.ToArray();
                _paragraphTextRuns.Add(runs);
                _paragraphFormats.Add(paragraphFormat);
                _paragraphs.Add(string.Concat(runs.Select(run => run.Text)));
                _bodyBlocks.Add(new LegacyDocParagraphBlock(runs, paragraphFormat));
                currentRuns.Clear();
                hasCurrentRun = false;
            }

            void AddCurrentTextAsTableCell() {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    justClosedCell = false;
                }

                if (currentRuns.Count == 0 && justClosedCell) {
                    if (currentTableRow.Count > 0) {
                        tableRows.Add(new LegacyDocTableRow(currentTableRow.ToArray()));
                        currentTableRow.Clear();
                    }

                    justClosedCell = false;
                    return;
                }

                currentTableRow.Add(new LegacyDocTableCell(string.Concat(currentRuns.Select(run => run.Text))));
                currentRuns.Clear();
                hasCurrentRun = false;
                justClosedCell = true;
            }

            void FlushTable() {
                FlushRun();
                if (currentRuns.Count > 0) {
                    currentTableRow.Add(new LegacyDocTableCell(string.Concat(currentRuns.Select(run => run.Text))));
                    currentRuns.Clear();
                }

                if (currentTableRow.Count > 0) {
                    tableRows.Add(new LegacyDocTableRow(currentTableRow.ToArray()));
                    currentTableRow.Clear();
                }

                if (tableRows.Count > 0) {
                    _bodyBlocks.Add(new LegacyDocTableBlock(tableRows.ToArray()));
                    tableRows.Clear();
                }

                hasCurrentRun = false;
                inTable = false;
                justClosedCell = false;
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

        private static LegacyDocParagraphFormat GetParagraphFormatForFileOffset(IReadOnlyList<LegacyDocParagraphFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocParagraphFormat.Default;
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
