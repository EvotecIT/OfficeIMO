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

        internal LegacyDocStyleSheet StyleSheet { get; private set; } = LegacyDocStyleSheet.Empty;

        internal LegacyDocSectionFormat SectionFormat { get; private set; } = LegacyDocSectionFormat.Default;

        internal IReadOnlyList<LegacyDocSection> Sections { get; private set; } = Array.Empty<LegacyDocSection>();

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
                AddKnownUnsupportedFeatureDiagnostics(compoundFile, fib);
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

            StyleSheet = LegacyDocStyleSheet.Read(tableStream, fib, fontFamilies, out string? styleSheetWarning);
            if (styleSheetWarning != null) {
                AddWarning("DOC-STYLESHEET-INVALID", styleSheetWarning);
            }

            Sections = LegacyDocSectionFormattingReader.ReadSections(wordDocumentStream, tableStream, fib, out string? sectionFormattingWarning);
            SectionFormat = Sections.Count == 0 ? LegacyDocSectionFormat.Default : Sections[0].Format;
            if (sectionFormattingWarning != null) {
                AddWarning("DOC-SEPX-INVALID", sectionFormattingWarning);
            }

            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges = LegacyDocCharacterFormattingReader.ReadCharacterFormatting(wordDocumentStream, tableStream, fib, fontFamilies, out string? formattingWarning);
            if (formattingWarning != null) {
                AddWarning("DOC-CHPX-INVALID", formattingWarning);
            }

            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges = LegacyDocParagraphFormattingReader.ReadParagraphFormatting(wordDocumentStream, tableStream, fib, out string? paragraphFormattingWarning);
            if (paragraphFormattingWarning != null) {
                AddWarning("DOC-PAPX-INVALID", paragraphFormattingWarning);
            }

            Text = BuildFormattedParagraphs(textContent.Characters, formattingRanges, paragraphFormattingRanges, Sections, options.ReportUnsupportedFeatures);
        }

        private void AddKnownUnsupportedFeatureDiagnostics(OfficeCompoundFile compoundFile, LegacyDocFib fib) {
            AddUnsupportedFibFlagFeatures(fib);

            AddUnsupportedCompoundEntryIfPresent(
                compoundFile,
                entry => entry.Path.StartsWith("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.StartsWith("Macros", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocUnsupportedFeatureKind.VbaProject,
                "DOC-MACROS-PRESENT",
                "The legacy DOC contains VBA project storage. Macros are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:VbaProjectStorage");

            AddUnsupportedCompoundEntryIfPresent(
                compoundFile,
                entry => entry.Path.IndexOf("ObjectPool", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocUnsupportedFeatureKind.OleObject,
                "DOC-OLE-OBJECTS-PRESENT",
                "The legacy DOC contains embedded OLE object storage. Embedded objects are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:OleObjectStorage");

            AddUnsupportedCompoundEntryIfPresent(
                compoundFile,
                entry => entry.Path.IndexOf("ActiveX", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.IndexOf("OCX", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocUnsupportedFeatureKind.ActiveXControl,
                "DOC-ACTIVEX-CONTROLS-PRESENT",
                "The legacy DOC contains ActiveX control storage. ActiveX controls are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:ActiveXControlStorage");

            AddUnsupportedCompoundEntryIfPresent(
                compoundFile,
                entry => string.Equals(entry.Name.TrimStart('\u0001'), "Ole10Native", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.IndexOf("Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.IndexOf("Package", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocUnsupportedFeatureKind.EmbeddedPackage,
                "DOC-EMBEDDED-PACKAGES-PRESENT",
                "The legacy DOC contains embedded package payload storage. Embedded packages are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:EmbeddedPackageStorage");

            AddUnsupportedDataStreamFeatureIfPresent(compoundFile);

            AddUnsupportedStoryFeatureIfPresent(
                fib.CcpHdd,
                LegacyDocUnsupportedFeatureKind.HeaderFooter,
                "DOC-HEADER-FOOTER-STORIES-PRESENT",
                "The legacy DOC contains header or footer story text. Headers and footers are preserved in the source file but are not projected into the OfficeIMO document.",
                "Fib:CcpHdd");
            AddUnsupportedStoryFeatureIfPresent(
                fib.CcpFtn,
                LegacyDocUnsupportedFeatureKind.Footnote,
                "DOC-FOOTNOTE-STORIES-PRESENT",
                "The legacy DOC contains footnote story text. Footnotes are preserved in the source file but are not projected into the OfficeIMO document.",
                "Fib:CcpFtn");
            AddUnsupportedStoryFeatureIfPresent(
                fib.CcpEdn,
                LegacyDocUnsupportedFeatureKind.Endnote,
                "DOC-ENDNOTE-STORIES-PRESENT",
                "The legacy DOC contains endnote story text. Endnotes are preserved in the source file but are not projected into the OfficeIMO document.",
                "Fib:CcpEdn");
            AddUnsupportedStoryFeatureIfPresent(
                fib.CcpAtn,
                LegacyDocUnsupportedFeatureKind.Comment,
                "DOC-COMMENT-STORIES-PRESENT",
                "The legacy DOC contains comment or annotation story text. Comments are preserved in the source file but are not projected into the OfficeIMO document.",
                "Fib:CcpAtn");
            AddUnsupportedStoryFeatureIfPresent(
                fib.CcpTxbx,
                LegacyDocUnsupportedFeatureKind.TextBox,
                "DOC-TEXTBOX-STORIES-PRESENT",
                "The legacy DOC contains text box story text. Text boxes are preserved in the source file but are not projected into the OfficeIMO document.",
                "Fib:CcpTxbx");
            AddUnsupportedStoryFeatureIfPresent(
                fib.CcpHdrTxbx,
                LegacyDocUnsupportedFeatureKind.TextBox,
                "DOC-HEADER-TEXTBOX-STORIES-PRESENT",
                "The legacy DOC contains header or footer text box story text. Header and footer text boxes are preserved in the source file but are not projected into the OfficeIMO document.",
                "Fib:CcpHdrTxbx");
        }

        private void AddUnsupportedCompoundEntryIfPresent(
            OfficeCompoundFile compoundFile,
            Func<OfficeCompoundFileEntry, bool> predicate,
            LegacyDocUnsupportedFeatureKind kind,
            string code,
            string description,
            string detailCode) {
            OfficeCompoundFileEntry? entry = compoundFile.Entries.FirstOrDefault(predicate);
            if (entry == null) {
                return;
            }

            AddUnsupportedFeature(new LegacyDocUnsupportedFeature(kind, code, description, entry.Path, detailCode));
        }

        private void AddUnsupportedFibFlagFeatures(LegacyDocFib fib) {
            if (fib.IsFastSaved || fib.QuickSaveCount > 0) {
                string detailCode = fib.IsFastSaved
                    ? "Fib:FComplex"
                    : "Fib:CQuickSaves";
                string description = fib.IsFastSaved
                    ? "The legacy DOC is marked as fast-saved or complex. Fast-save deltas are preserved in the source file but are not projected into the OfficeIMO document."
                    : $"The legacy DOC reports {fib.QuickSaveCount} quick-save revision(s). Quick-save history is preserved in the source file but is not projected into the OfficeIMO document.";
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.FastSave,
                    "DOC-FAST-SAVE-PRESENT",
                    description,
                    detailCode: detailCode));
            }

            if (fib.HasPictures) {
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.Picture,
                    "DOC-PICTURES-PRESENT",
                    "The legacy DOC FIB indicates picture payloads. Pictures are preserved in the source file but are not projected into the OfficeIMO document.",
                    detailCode: "Fib:FHasPic"));
            }
        }

        private void AddUnsupportedDataStreamFeatureIfPresent(OfficeCompoundFile compoundFile) {
            OfficeCompoundFileEntry? entry = compoundFile.Entries.FirstOrDefault(item =>
                item.IsStream && string.Equals(item.Name, "Data", StringComparison.OrdinalIgnoreCase));
            if (entry == null) {
                return;
            }

            if (!compoundFile.Streams.TryGetValue(entry.Path, out byte[]? dataStream)
                && !compoundFile.Streams.TryGetValue(entry.Name, out dataStream)) {
                return;
            }

            if (dataStream.Length == 0) {
                return;
            }

            AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                LegacyDocUnsupportedFeatureKind.BinaryData,
                "DOC-BINARY-DATA-STREAM-PRESENT",
                "The legacy DOC contains a binary Data stream used by pictures, drawings, form fields, or other payloads. These payloads are preserved in the source file but are not projected into the OfficeIMO document.",
                entry.Path,
                "Compound:BinaryDataStream"));
        }

        private void AddUnsupportedStoryFeatureIfPresent(
            int characterCount,
            LegacyDocUnsupportedFeatureKind kind,
            string code,
            string description,
            string detailCode) {
            if (characterCount <= 0) {
                return;
            }

            AddUnsupportedFeature(new LegacyDocUnsupportedFeature(kind, code, description, detailCode: detailCode));
        }

        private string BuildFormattedParagraphs(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            IReadOnlyList<LegacyDocSection> sections,
            bool reportUnsupportedFeatures) {
            var bodyText = new System.Text.StringBuilder(characters.Count);
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            bool hasCurrentRun = false;
            bool inTable = false;
            bool justClosedCell = false;
            var tableRows = new List<LegacyDocTableRow>();
            var currentTableRow = new List<LegacyDocTableCell>();
            int nextSectionIndex = sections.Count > 1 ? 1 : sections.Count;
            bool reportedUnprojectedSectionBoundary = false;

            foreach (LegacyDocTextCharacter textCharacter in characters) {
                if (textCharacter.Character == '\a') {
                    LegacyDocParagraphFormat paragraphFormat = GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset);
                    if (paragraphFormat.IsInTable == true) {
                        if (paragraphFormat.IsTableTerminatingParagraph == true) {
                            AddCurrentTableRow(paragraphFormat);
                        } else {
                            AddCurrentTextAsTableCell(paragraphFormat, allowHeuristicRowTerminator: false);
                        }
                    } else {
                        AddCurrentTextAsTableCell(paragraphFormat, allowHeuristicRowTerminator: true);
                    }

                    AddSectionBreaksAtBodyBoundary(textCharacter.CharacterPosition + 1);
                    continue;
                }

                char? normalized = NormalizeBodyCharacter(textCharacter.Character);
                if (normalized == null) {
                    continue;
                }

                LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset);
                if (normalized.Value == '\r') {
                    if (inTable) {
                        FlushTable(GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset));
                    } else {
                        AddCurrentTextAsParagraph(GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset));
                    }

                    bodyText.Append('\r');
                    AddSectionBreaksAtBodyBoundary(textCharacter.CharacterPosition + 1);
                    continue;
                }

                AppendRunCharacter(normalized.Value, format);
                bodyText.Append(normalized.Value);
            }

            if (inTable) {
                FlushTable(LegacyDocParagraphFormat.Default);
            } else if (currentRuns.Count > 0 || runText.Length > 0) {
                AddCurrentTextAsParagraph(LegacyDocParagraphFormat.Default);
            }

            ReportRemainingUnprojectedSectionBoundaries();
            return bodyText.ToString();

            void AddSectionBreaksAtBodyBoundary(int characterPosition) {
                while (nextSectionIndex < sections.Count && sections[nextSectionIndex].StartCharacter < characterPosition) {
                    ReportUnprojectedSectionBoundary(sections[nextSectionIndex].StartCharacter);
                    nextSectionIndex++;
                }

                while (nextSectionIndex < sections.Count && sections[nextSectionIndex].StartCharacter == characterPosition) {
                    if (characterPosition < characters.Count) {
                        _bodyBlocks.Add(new LegacyDocSectionBreakBlock(sections[nextSectionIndex].Format));
                    }

                    nextSectionIndex++;
                }
            }

            void ReportRemainingUnprojectedSectionBoundaries() {
                while (nextSectionIndex < sections.Count) {
                    if (sections[nextSectionIndex].StartCharacter < characters.Count) {
                        ReportUnprojectedSectionBoundary(sections[nextSectionIndex].StartCharacter);
                    }

                    nextSectionIndex++;
                }
            }

            void ReportUnprojectedSectionBoundary(int characterPosition) {
                if (!reportUnsupportedFeatures || reportedUnprojectedSectionBoundary) {
                    return;
                }

                reportedUnprojectedSectionBoundary = true;
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.Section,
                    "DOC-MULTIPLE-SECTIONS-PRESENT",
                    $"The legacy DOC contains a section boundary at character position {characterPosition} that does not align with a supported body-block boundary. That section is preserved in the source file but is not projected into the OfficeIMO document.",
                    detailCode: "Fib:PlcfSed"));
            }

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
                    currentFormat.Strike,
                    currentFormat.DoubleStrike,
                    currentFormat.Outline,
                    currentFormat.Shadow,
                    currentFormat.Emboss,
                    currentFormat.Imprint,
                    currentFormat.Hidden,
                    currentFormat.Caps,
                    currentFormat.VerticalPosition,
                    currentFormat.Underline,
                    currentFormat.Highlight,
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

            void AddCurrentTextAsTableCell(LegacyDocParagraphFormat paragraphFormat, bool allowHeuristicRowTerminator) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    justClosedCell = false;
                }

                if (allowHeuristicRowTerminator && currentRuns.Count == 0 && justClosedCell) {
                    if (currentTableRow.Count > 0) {
                        tableRows.Add(new LegacyDocTableRow(currentTableRow.ToArray()));
                        currentTableRow.Clear();
                    }

                    justClosedCell = false;
                    return;
                }

                currentTableRow.Add(new LegacyDocTableCell(currentRuns.ToArray(), paragraphFormat));
                currentRuns.Clear();
                hasCurrentRun = false;
                justClosedCell = true;
            }

            void AddCurrentTableRow(LegacyDocParagraphFormat paragraphFormat) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                }

                if (currentRuns.Count > 0 || (!justClosedCell && currentTableRow.Count == 0)) {
                    currentTableRow.Add(new LegacyDocTableCell(currentRuns.ToArray(), paragraphFormat));
                    currentRuns.Clear();
                }

                if (currentTableRow.Count > 0) {
                    tableRows.Add(new LegacyDocTableRow(currentTableRow.ToArray(), paragraphFormat.TableCellWidthsTwips));
                    currentTableRow.Clear();
                }

                hasCurrentRun = false;
                justClosedCell = false;
            }

            void FlushTable(LegacyDocParagraphFormat paragraphFormat) {
                FlushRun();
                if (currentRuns.Count > 0) {
                    currentTableRow.Add(new LegacyDocTableCell(currentRuns.ToArray(), paragraphFormat));
                    currentRuns.Clear();
                }

                if (currentTableRow.Count > 0) {
                    tableRows.Add(new LegacyDocTableRow(currentTableRow.ToArray(), paragraphFormat.TableCellWidthsTwips));
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
                    return character;
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
