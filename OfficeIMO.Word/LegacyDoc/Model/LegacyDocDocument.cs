using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml.Wordprocessing;
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
        private readonly List<LegacyDocPreservedFeature> _preservedFeatures = new();
        private readonly List<LegacyDocCompoundFeature> _compoundFeatures = new();
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

        internal IReadOnlyList<LegacyDocHeaderFooterStory> HeaderFooterStories { get; private set; } = Array.Empty<LegacyDocHeaderFooterStory>();

        internal IReadOnlyList<LegacyDocFootnote> Footnotes { get; private set; } = Array.Empty<LegacyDocFootnote>();

        internal IReadOnlyList<LegacyDocEndnote> Endnotes { get; private set; } = Array.Empty<LegacyDocEndnote>();

        internal IReadOnlyList<LegacyDocComment> Comments { get; private set; } = Array.Empty<LegacyDocComment>();

        internal IReadOnlyList<LegacyDocTextBoxStory> TextBoxStories { get; private set; } = Array.Empty<LegacyDocTextBoxStory>();

        internal IReadOnlyList<LegacyDocBookmark> Bookmarks { get; private set; } = Array.Empty<LegacyDocBookmark>();

        internal bool DifferentOddAndEvenPages { get; private set; }

        internal bool RevisionMarkingEnabled { get; private set; }

        internal bool LockedRevisionTrackingEnabled { get; private set; }

        /// <summary>Gets the private source container retained for same-format compound rewriting.</summary>
        internal OfficeCompoundFile? SourceCompoundFile { get; private set; }

        /// <summary>Gets diagnostics produced while reading the legacy document.</summary>
        public IReadOnlyList<LegacyDocImportDiagnostic> Diagnostics => _diagnostics;

        /// <summary>Gets unsupported or preserve-only features discovered while reading the legacy document.</summary>
        public IReadOnlyList<LegacyDocUnsupportedFeature> UnsupportedFeatures => _unsupportedFeatures;

        /// <summary>Gets preserve-only non-compound feature metadata discovered while reading the legacy document.</summary>
        public IReadOnlyList<LegacyDocPreservedFeature> PreservedFeatures => _preservedFeatures;

        /// <summary>Gets preserve-only compound storage discovered while reading the legacy document.</summary>
        public IReadOnlyList<LegacyDocCompoundFeature> CompoundFeatures => _compoundFeatures;

        /// <summary>
        /// Loads a legacy DOC model from a file path.
        /// </summary>
        public static LegacyDocDocument Load(string path, LegacyDocImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            options ??= new LegacyDocImportOptions();
            options.Validate();
            using FileStream stream = File.OpenRead(path);
            return Load(OfficeStreamReader.ReadAllBytes(stream, options.MaxInputBytes), options);
        }

        /// <summary>
        /// Loads a legacy DOC model from a stream.
        /// </summary>
        public static LegacyDocDocument Load(Stream stream, LegacyDocImportOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            options ??= new LegacyDocImportOptions();
            options.Validate();
            return Load(OfficeStreamReader.ReadAllBytes(stream, options.MaxInputBytes), options);
        }

        /// <summary>
        /// Loads a legacy DOC model from compound document bytes.
        /// </summary>
        public static LegacyDocDocument Load(byte[] bytes, LegacyDocImportOptions? options = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            options ??= new LegacyDocImportOptions();
            options.Validate();
            if (bytes.Length > options.MaxInputBytes) {
                throw new InvalidDataException($"Legacy DOC input exceeds the configured maximum size ({options.MaxInputBytes} bytes).");
            }

            var document = new LegacyDocDocument();
            if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compoundFile, out string? compoundError)) {
                document.AddError("DOC-COMPOUND-INVALID", compoundError ?? "The OLE compound document could not be read.");
                return document;
            }

            document.SourceCompoundFile = compoundFile;
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
            if (!TryGetRootStream(compoundFile, "WordDocument", out byte[]? wordDocumentStreamCandidate)) {
                AddError("DOC-WORDDOCUMENT-MISSING", "The compound document does not contain a WordDocument stream.");
                return;
            }

            byte[] wordDocumentStream = wordDocumentStreamCandidate!;
            if (wordDocumentStream.Length > options.MaxInputBytes) {
                AddError("DOC-WORDDOCUMENT-TOO-LARGE", $"The WordDocument stream is {wordDocumentStream.Length} bytes, which exceeds the configured limit of {options.MaxInputBytes} bytes.");
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

            LegacyDocOleDocumentPropertyReader.AddDocumentProperties(compoundFile, this, options);

            string tableStreamName = fib.UsesOneTableStream ? "1Table" : "0Table";
            if (!TryGetRootStream(compoundFile, tableStreamName, out byte[]? tableStreamCandidate)) {
                string alternateName = fib.UsesOneTableStream ? "0Table" : "1Table";
                if (TryGetRootStream(compoundFile, alternateName, out tableStreamCandidate)) {
                    AddWarning("DOC-TABLE-STREAM-FALLBACK", $"The FIB requested {tableStreamName}, but only {alternateName} was present. The available table stream was used.");
                } else {
                    AddError("DOC-TABLE-STREAM-MISSING", $"The compound document does not contain the {tableStreamName} table stream.");
                    return;
                }
            }

            byte[] tableStream = tableStreamCandidate!;
            DifferentOddAndEvenPages = ReadDopFacingPagesFlag(tableStream, fib);
            EndnotePositionValues? dopEndnotePosition = ReadDopEndnotePlacement(tableStream, fib);

            if (!LegacyDocPieceTable.TryRead(wordDocumentStream, tableStream, fib, out LegacyDocTextContent textContent, out string? textError)) {
                AddError("DOC-PIECE-TABLE-INVALID", textError ?? "The legacy DOC piece table could not be decoded.");
                return;
            }

            IReadOnlyList<string> fontFamilies = LegacyDocFontTableReader.ReadFontFamilies(tableStream, fib, out string? fontTableWarning);
            if (fontTableWarning != null) {
                AddWarning("DOC-FONT-TABLE-INVALID", fontTableWarning);
            }

            IReadOnlyList<string> revisionAuthors = LegacyDocRevisionAuthorReader.Read(tableStream, fib, out string? revisionAuthorWarning);
            if (revisionAuthorWarning != null) {
                AddWarning("DOC-REVISION-AUTHORS-INVALID", revisionAuthorWarning);
            }

            StyleSheet = LegacyDocStyleSheet.Read(tableStream, fib, fontFamilies, out string? styleSheetWarning);
            if (styleSheetWarning != null) {
                AddWarning("DOC-STYLESHEET-INVALID", styleSheetWarning);
            }

            Sections = ApplyDopEndnotePlacement(
                LegacyDocSectionFormattingReader.ReadSections(wordDocumentStream, tableStream, fib, out string? sectionFormattingWarning),
                fib,
                dopEndnotePosition);
            SectionFormat = Sections.Count == 0 ? LegacyDocSectionFormat.Default : Sections[0].Format;
            if (sectionFormattingWarning != null) {
                AddWarning("DOC-SEPX-INVALID", sectionFormattingWarning);
            }

            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges = LegacyDocCharacterFormattingReader.ReadCharacterFormatting(
                wordDocumentStream,
                tableStream,
                fib,
                fontFamilies,
                revisionAuthors,
                out string? formattingWarning);
            if (formattingWarning != null) {
                AddWarning("DOC-CHPX-INVALID", formattingWarning);
            }

            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges = LegacyDocParagraphFormattingReader.ReadParagraphFormatting(wordDocumentStream, tableStream, fib, out string? paragraphFormattingWarning);
            if (paragraphFormattingWarning != null) {
                AddWarning("DOC-PAPX-INVALID", paragraphFormattingWarning);
            }

            Bookmarks = LegacyDocBookmarkReader.Read(tableStream, fib, out string? bookmarkWarning);
            if (bookmarkWarning != null) {
                AddWarning("DOC-BOOKMARK-PLC-INVALID", bookmarkWarning);
            }
            var bookmarkProjection = new LegacyDocBookmarkProjectionTracker(Bookmarks);

            AddUnsupportedParagraphFormattingFeaturesIfPresent(paragraphFormattingRanges, options.ReportUnsupportedContent);

            byte[] dataStream = TryGetRootStream(compoundFile, "Data", out byte[]? dataStreamCandidate)
                ? dataStreamCandidate!
                : Array.Empty<byte>();
            LegacyDocPictureReader.LegacyDocPictureReadResult pictures = LegacyDocPictureReader.Read(
                dataStream,
                textContent.AllCharacters,
                formattingRanges,
                fib.CcpText + fib.CcpFtn + fib.CcpHdd + fib.CcpAtn + fib.CcpEdn,
                options.MaxDecodedImageBytes);
            if (pictures.Warning != null) {
                AddWarning("DOC-PICTURE-DATA-INVALID", pictures.Warning);
            }

            Text = BuildFormattedParagraphs(
                textContent.Characters,
                formattingRanges,
                paragraphFormattingRanges,
                Sections,
                bookmarkProjection,
                pictures.PicturesByCharacterPosition,
                options.ReportUnsupportedContent);
            Footnotes = LegacyDocFootnoteReader.Read(tableStream, textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition, out string? footnoteWarning);
            if (footnoteWarning != null) {
                AddWarning("DOC-FOOTNOTE-PLC-INVALID", footnoteWarning);
            }

            Endnotes = LegacyDocFootnoteReader.ReadEndnotes(tableStream, textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition, out string? endnoteWarning);
            if (endnoteWarning != null) {
                AddWarning("DOC-ENDNOTE-PLC-INVALID", endnoteWarning);
            }

            Comments = LegacyDocCommentReader.Read(tableStream, textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition, out string? commentWarning);
            if (commentWarning != null) {
                AddWarning("DOC-COMMENT-PLC-INVALID", commentWarning);
            }

            TextBoxStories = LegacyDocTextBoxStoryReader.Read(textContent, fib, formattingRanges, bookmarkProjection);
            AddKnownUnsupportedFeatureDiagnostics(
                compoundFile,
                tableStream,
                fib,
                pictures.FullyProjectsDataStream,
                options.ReportUnsupportedContent);

            HeaderFooterStories = LegacyDocHeaderFooterReader.Read(tableStream, textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition, out string? headerFooterWarning);
            if (headerFooterWarning != null) {
                AddWarning("DOC-PLCFHDD-INVALID", headerFooterWarning);
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.HeaderFooter,
                    "DOC-HEADER-FOOTER-STORIES-PRESENT",
                    "The legacy DOC contains header or footer story text with an unsupported header/footer story PLC. Headers and footers are preserved in the source file but are not projected into the OfficeIMO document.",
                    detailCode: "Fib:PlcfHdd"),
                    options.ReportUnsupportedContent);
            }

            ReportRemainingUnprojectedBookmarks(bookmarkProjection);
        }

        private static bool TryGetRootStream(OfficeCompoundFile compoundFile, string name, out byte[]? bytes) {
            OfficeCompoundFileEntry? entry = compoundFile.Entries.FirstOrDefault(item =>
                item.IsStream && string.Equals(item.Path, name, StringComparison.OrdinalIgnoreCase));
            if (entry != null && compoundFile.Streams.TryGetValue(entry.Path, out bytes)) {
                return true;
            }

            return compoundFile.Streams.TryGetValue(name, out bytes);
        }

        private void AddKnownUnsupportedFeatureDiagnostics(OfficeCompoundFile compoundFile, byte[] tableStream, LegacyDocFib fib, bool pictureDataFullyProjected, bool reportDiagnostic) {
            AddUnsupportedFibFlagFeatures(fib, pictureDataFullyProjected, reportDiagnostic);

            AddCompoundFeatureIfPresent(
                compoundFile,
                entry => entry.Path.StartsWith("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.StartsWith("Macros", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocCompoundFeatureKind.VbaProject,
                "DOC-MACROS-PRESENT",
                "The legacy DOC contains VBA project storage. Macros are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:VbaProjectStorage");

            AddCompoundFeatureIfPresent(
                compoundFile,
                entry => entry.Path.IndexOf("ObjectPool", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocCompoundFeatureKind.OleObject,
                "DOC-OLE-OBJECTS-PRESENT",
                "The legacy DOC contains embedded OLE object storage. Embedded objects are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:OleObjectStorage");

            AddCompoundFeatureIfPresent(
                compoundFile,
                entry => entry.Path.IndexOf("ActiveX", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.IndexOf("OCX", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocCompoundFeatureKind.ActiveXControl,
                "DOC-ACTIVEX-CONTROLS-PRESENT",
                "The legacy DOC contains ActiveX control storage. ActiveX controls are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:ActiveXControlStorage");

            AddCompoundFeatureIfPresent(
                compoundFile,
                entry => string.Equals(entry.Name.TrimStart('\u0001'), "Ole10Native", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.IndexOf("Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.IndexOf("Package", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocCompoundFeatureKind.EmbeddedPackage,
                "DOC-EMBEDDED-PACKAGES-PRESENT",
                "The legacy DOC contains embedded package payload storage. Embedded packages are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:EmbeddedPackageStorage");

            AddDataStreamCompoundFeatureIfPresent(compoundFile, pictureDataFullyProjected);

            AddCompoundFeatureIfPresent(
                compoundFile,
                IsDigitalSignatureEntry,
                LegacyDocCompoundFeatureKind.DigitalSignature,
                "DOC-DIGITAL-SIGNATURE-PRESENT",
                "The legacy DOC contains digital-signature storage. Rewriting the document invalidates the existing signature.",
                "Compound:DigitalSignature");
            AddUnsupportedCompoundEntryIfPresent(
                compoundFile,
                IsDigitalSignatureEntry,
                LegacyDocUnsupportedFeatureKind.DigitalSignature,
                "DOC-DIGITAL-SIGNATURE-PRESENT",
                "The legacy DOC contains digital-signature storage. Saving is blocked by default because rewriting the document invalidates the existing signature.",
                "Compound:DigitalSignature",
                reportDiagnostic);

            if (fib.CcpHdd > 0 && fib.LcbPlcfHdd == 0) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpHdd,
                    LegacyDocUnsupportedFeatureKind.HeaderFooter,
                    "DOC-HEADER-FOOTER-STORIES-PRESENT",
                    "The legacy DOC contains header or footer story text without a supported header/footer story PLC. Headers and footers are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpHdd",
                    reportDiagnostic);
            }
            if (!LegacyDocFootnoteReader.HasReadableFootnoteTables(tableStream, fib)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpFtn,
                    LegacyDocUnsupportedFeatureKind.Footnote,
                    "DOC-FOOTNOTE-STORIES-PRESENT",
                    "The legacy DOC contains footnote story text. Footnotes are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpFtn",
                    reportDiagnostic);
            }
            if (!LegacyDocFootnoteReader.HasReadableEndnoteTables(tableStream, fib)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpEdn,
                    LegacyDocUnsupportedFeatureKind.Endnote,
                    "DOC-ENDNOTE-STORIES-PRESENT",
                    "The legacy DOC contains endnote story text. Endnotes are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpEdn",
                    reportDiagnostic);
            }
            if (!LegacyDocCommentReader.HasReadableCommentTables(tableStream, fib)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpAtn,
                    LegacyDocUnsupportedFeatureKind.Comment,
                    "DOC-COMMENT-STORIES-PRESENT",
                    "The legacy DOC contains comment or annotation story text. Comments are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpAtn",
                    reportDiagnostic);
            }
            ReadRevisionTrackingState(tableStream, fib);
            if (fib.CcpTxbx > 0 && !TextBoxStories.Any(story => !story.IsHeaderFooterTextBox)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpTxbx,
                    LegacyDocUnsupportedFeatureKind.TextBox,
                    "DOC-TEXTBOX-STORIES-PRESENT",
                    "The legacy DOC contains text box story text that could not be read from the piece table. Text boxes are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpTxbx",
                    reportDiagnostic);
            }

            if (fib.CcpHdrTxbx > 0) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpHdrTxbx,
                    LegacyDocUnsupportedFeatureKind.TextBox,
                    "DOC-HEADER-TEXTBOX-STORIES-PRESENT",
                    "The legacy DOC contains header or footer text box story text that could not be read from the piece table. Header and footer text boxes are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpHdrTxbx",
                    reportDiagnostic);
            }
        }

        private void AddCompoundFeatureIfPresent(
            OfficeCompoundFile compoundFile,
            Func<OfficeCompoundFileEntry, bool> predicate,
            LegacyDocCompoundFeatureKind kind,
            string code,
            string description,
            string detailCode) {
            OfficeCompoundFileEntry[] entries = compoundFile.Entries
                .Where(predicate)
                .OrderBy(entry => entry.Path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (entries.Length == 0) {
                return;
            }

            _compoundFeatures.Add(new LegacyDocCompoundFeature(
                kind,
                code,
                description,
                entries[0].Path,
                detailCode,
                entries.Length,
                entries.Sum(entry => entry.Size)));
        }

        private void AddUnsupportedCompoundEntryIfPresent(
            OfficeCompoundFile compoundFile,
            Func<OfficeCompoundFileEntry, bool> predicate,
            LegacyDocUnsupportedFeatureKind kind,
            string code,
            string description,
            string detailCode,
            bool reportDiagnostic) {
            OfficeCompoundFileEntry? entry = compoundFile.Entries.FirstOrDefault(predicate);
            if (entry == null) {
                return;
            }

            AddUnsupportedFeature(new LegacyDocUnsupportedFeature(kind, code, description, entry.Path, detailCode), reportDiagnostic);
        }

        private void AddUnsupportedFibFlagFeatures(LegacyDocFib fib, bool pictureDataFullyProjected, bool reportDiagnostic) {
            if (fib.IsFastSaved) {
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.FastSave,
                    "DOC-FAST-SAVE-PRESENT",
                    "The legacy DOC is marked as fast-saved or complex. Fast-save deltas are preserved in the source file but are not projected into the OfficeIMO document.",
                    detailCode: "Fib:FComplex"),
                    reportDiagnostic);
            }

            if (fib.QuickSaveCount > 0 && reportDiagnostic) {
                AddWarning(
                    "DOC-QUICK-SAVE-HISTORY-PRESENT",
                    $"The legacy DOC reports {fib.QuickSaveCount} quick-save revision(s). The projected document content is readable; quick-save history is not represented as editable revision history.");
            }

            if (fib.HasPictures && pictureDataFullyProjected) {
                AddInfo(
                    "DOC-PICTURES-PROJECTED",
                    "Supported inline picture payloads were projected from the legacy DOC Data stream into editable Open XML images.");
            } else if (fib.HasPictures) {
                _preservedFeatures.Add(new LegacyDocPreservedFeature(
                    LegacyDocPreservedFeatureKind.Picture,
                    "DOC-PICTURES-PRESENT",
                    "The legacy DOC FIB indicates picture payloads. Pictures are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:FHasPic"));
            }
        }

        private void AddDataStreamCompoundFeatureIfPresent(OfficeCompoundFile compoundFile, bool pictureDataFullyProjected) {
            if (pictureDataFullyProjected) {
                return;
            }

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

            _compoundFeatures.Add(new LegacyDocCompoundFeature(
                LegacyDocCompoundFeatureKind.BinaryData,
                "DOC-BINARY-DATA-STREAM-PRESENT",
                "The legacy DOC contains a binary Data stream used by pictures, drawings, form fields, or other payloads. These payloads are preserved in the source file but are not projected into the OfficeIMO document.",
                entry.Path,
                "Compound:BinaryDataStream",
                entryCount: 1,
                totalBytes: dataStream.Length));
        }

        private static bool IsDigitalSignatureEntry(OfficeCompoundFileEntry entry) {
            return entry.Name.Equals("_signatures", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("_xmlsignatures", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/_xmlsignatures", StringComparison.OrdinalIgnoreCase);
        }

        private void ReadRevisionTrackingState(byte[] tableStream, LegacyDocFib fib) {
            if (fib.LcbDop < 8 || fib.FcDop < 0 || fib.FcDop > tableStream.Length - fib.LcbDop) {
                return;
            }

            const uint revisionMarkingFlag = 0x00008000;
            const uint lockRevisionFlag = 0x40000000;
            uint dopSecondFlags = unchecked((uint)LegacyDocFib.ReadInt32(tableStream, fib.FcDop + 4));
            bool hasRevisionMarking = (dopSecondFlags & revisionMarkingFlag) != 0;
            bool hasLockedRevisionTracking = (dopSecondFlags & lockRevisionFlag) != 0;
            RevisionMarkingEnabled = hasRevisionMarking;
            LockedRevisionTrackingEnabled = hasLockedRevisionTracking;
        }

        private static bool ReadDopFacingPagesFlag(byte[] tableStream, LegacyDocFib fib) {
            const ushort facingPagesFlag = 0x0001;
            if (fib.LcbDop < 2 || fib.FcDop < 0 || fib.FcDop > tableStream.Length - fib.LcbDop) {
                return false;
            }

            ushort dopFlags = LegacyDocFib.ReadUInt16(tableStream, fib.FcDop);
            return (dopFlags & facingPagesFlag) != 0;
        }

        private static EndnotePositionValues? ReadDopEndnotePlacement(byte[] tableStream, LegacyDocFib fib) {
            const int endnotePlacementOffset = 52;
            const int minimumDopLength = endnotePlacementOffset + 4;
            const int endnotePlacementShift = 16;
            const uint endnotePlacementMask = 0x3;
            if (fib.LcbDop < minimumDopLength || fib.FcDop < 0 || fib.FcDop > tableStream.Length - fib.LcbDop) {
                return null;
            }

            uint row = unchecked((uint)LegacyDocFib.ReadInt32(tableStream, fib.FcDop + endnotePlacementOffset));
            uint placement = (row >> endnotePlacementShift) & endnotePlacementMask;
            switch (placement) {
                case 0:
                    return EndnotePositionValues.SectionEnd;
                case 3:
                    return EndnotePositionValues.DocumentEnd;
                default:
                    return null;
            }
        }

        private static IReadOnlyList<LegacyDocSection> ApplyDopEndnotePlacement(IReadOnlyList<LegacyDocSection> sections, LegacyDocFib fib, EndnotePositionValues? endnotePosition) {
            if (endnotePosition == null) {
                return sections;
            }

            if (sections.Count == 0) {
                return new[] {
                    new LegacyDocSection(
                        0,
                        Math.Max(0, fib.CcpText),
                        LegacyDocSectionFormat.Default.WithEndnotePosition(endnotePosition))
                };
            }

            var projected = new LegacyDocSection[sections.Count];
            for (int index = 0; index < sections.Count; index++) {
                LegacyDocSection section = sections[index];
                projected[index] = new LegacyDocSection(
                    section.StartCharacter,
                    section.EndCharacter,
                    section.Format.WithEndnotePosition(endnotePosition));
            }

            return projected;
        }

        private void AddUnsupportedStoryFeatureIfPresent(
            int characterCount,
            LegacyDocUnsupportedFeatureKind kind,
            string code,
            string description,
            string detailCode,
            bool reportDiagnostic) {
            if (characterCount <= 0) {
                return;
            }

            AddUnsupportedFeature(new LegacyDocUnsupportedFeature(kind, code, description, detailCode: detailCode), reportDiagnostic);
        }

        private void AddUnsupportedParagraphFormattingFeaturesIfPresent(IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges, bool reportUnsupportedFeatures) {
            bool reportedNestedTable = false;
            for (int index = 0; index < paragraphFormattingRanges.Count; index++) {
                LegacyDocParagraphFormat format = paragraphFormattingRanges[index].Format;
                if (!reportedNestedTable && format.MaximumTableDepth > 2) {
                    AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                        LegacyDocUnsupportedFeatureKind.NestedTable,
                        "DOC-NESTED-TABLES-PRESENT",
                        BuildNestedTableDescription(format),
                        detailCode: "PAPX:sprmPItap"),
                        reportUnsupportedFeatures);
                    reportedNestedTable = true;
                }

                if (format.HasMergedTableCells) {
                    AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                        LegacyDocUnsupportedFeatureKind.MergedTableCell,
                        "DOC-MERGED-TABLE-CELLS-PRESENT",
                        "The legacy DOC contains unsupported or conflicting merged table cell descriptors. That table structure is preserved in the source file but cannot be safely projected into the OfficeIMO table model yet.",
                        detailCode: "PAPX:sprmTDefTable"),
                        reportUnsupportedFeatures);
                    return;
                }
            }
        }

        private static string BuildNestedTableDescription(LegacyDocParagraphFormat format) {
            string depth = format.MaximumTableDepth > 1
                ? $"maximum table depth {format.MaximumTableDepth}"
                : "nested table depth marker";
            var markers = new List<string>();
            if (format.HasInnerTableCellMarker) {
                markers.Add("inner table cell marker");
            }

            if (format.HasInnerTableTerminatingParagraphMarker) {
                markers.Add("inner table row marker");
            }

            string markerText = markers.Count == 0
                ? string.Empty
                : $"; {string.Join(" and ", markers)} present";
            return $"The legacy DOC contains nested table descriptors ({depth}{markerText}). Nested tables are preserved in the source file but cannot be safely projected into the OfficeIMO table model yet.";
        }

        private void ReportRemainingUnprojectedBookmarks(LegacyDocBookmarkProjectionTracker bookmarkProjection) {
            LegacyDocBookmark? bookmark = bookmarkProjection.GetUnprojectedBookmarks().FirstOrDefault();
            if (bookmark == null) {
                return;
            }

            _preservedFeatures.Add(new LegacyDocPreservedFeature(
                LegacyDocPreservedFeatureKind.Bookmark,
                "DOC-BOOKMARK-RANGE-PRESENT",
                $"The legacy DOC contains bookmark '{bookmark.Name}' at character range {bookmark.StartCharacter}-{bookmark.EndCharacter} outside the currently supported body, table-cell, header/footer, footnote, and endnote paragraph bookmark projection. The bookmark is preserved in the source file but is not projected into the OfficeIMO document.",
                "Fib:PlcfBkf"));
        }

        private string BuildFormattedParagraphs(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            IReadOnlyList<LegacyDocSection> sections,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
            IReadOnlyDictionary<int, LegacyDocPicture> picturesByCharacterPosition,
            bool reportUnsupportedFeatures) {
            var bodyText = new System.Text.StringBuilder(characters.Count);
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder();
            var runCharacterPositions = new List<int>();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            LegacyDocHyperlinkTarget currentHyperlinkTarget = default;
            bool hasCurrentRun = false;
            bool inTable = false;
            bool justClosedCell = false;
            var tableRows = new List<LegacyDocTableRow>();
            var currentTableRow = new List<LegacyDocTableCell>();
            var currentTableCellParagraphs = new List<LegacyDocTableCellParagraph>();
            int nextSectionIndex = sections.Count > 1 ? 1 : sections.Count;
            bool reportedUnprojectedSectionBoundary = false;
            int currentParagraphStartCharacter = 0;
            int? currentTableStartCharacter = null;
            int? currentTableRowStartCharacter = null;
            IReadOnlyList<LegacyDocBookmark>? currentTableRowBoundaryBookmarks = null;

            for (int characterIndex = 0; characterIndex < characters.Count; characterIndex++) {
                LegacyDocTextCharacter textCharacter = characters[characterIndex];
                if (LegacyDocPictureReader.TryCreatePictureRun(
                    textCharacter,
                    picturesByCharacterPosition,
                    GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset),
                    out LegacyDocTextRun? pictureRun)) {
                    FlushRun();
                    currentRuns.Add(pictureRun!);
                    if (inTable) {
                        justClosedCell = false;
                    }

                    continue;
                }

                if (LegacyDocField.TryReadHyperlink(
                    characters,
                    characterIndex,
                    out LegacyDocHyperlinkTarget hyperlinkTarget,
                    out int resultStartIndex,
                    out int resultEndIndex,
                    out int fieldEndIndex)) {
                    AppendHyperlinkResult(hyperlinkTarget, resultStartIndex, resultEndIndex);
                    characterIndex = fieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadPageNumber(
                    characters,
                    characterIndex,
                    out int pageNumberResultStartIndex,
                    out int pageNumberResultEndIndex,
                    out int pageNumberFieldEndIndex)) {
                    AppendPageNumberResult(pageNumberResultStartIndex, pageNumberResultEndIndex);
                    characterIndex = pageNumberFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadNumberOfPages(
                    characters,
                    characterIndex,
                    out int numberOfPagesResultStartIndex,
                    out int numberOfPagesResultEndIndex,
                    out int numberOfPagesFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.NumPages, fieldInstruction: null, numberOfPagesResultStartIndex, numberOfPagesResultEndIndex);
                    characterIndex = numberOfPagesFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadDateTimeField(
                    characters,
                    characterIndex,
                    out LegacyDocFieldKind dateTimeFieldKind,
                    out string dateInstruction,
                    out int dateResultStartIndex,
                    out int dateResultEndIndex,
                    out int dateFieldEndIndex)) {
                    AppendFieldResult(dateTimeFieldKind, dateInstruction, dateResultStartIndex, dateResultEndIndex);
                    characterIndex = dateFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadDocumentPropertyField(
                    characters,
                    characterIndex,
                    out string documentPropertyInstruction,
                    out int documentPropertyResultStartIndex,
                    out int documentPropertyResultEndIndex,
                    out int documentPropertyFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.DocumentProperty, documentPropertyInstruction, documentPropertyResultStartIndex, documentPropertyResultEndIndex);
                    characterIndex = documentPropertyFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadEquationField(
                    characters,
                    characterIndex,
                    out string equationInstruction,
                    out int equationResultStartIndex,
                    out int equationResultEndIndex,
                    out int equationFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.Equation, equationInstruction, equationResultStartIndex, equationResultEndIndex);
                    characterIndex = equationFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadDisplayField(
                    characters,
                    characterIndex,
                    out int fallbackFieldResultStartIndex,
                    out int fallbackFieldResultEndIndex,
                    out int fallbackFieldEndIndex)) {
                    AppendFieldDisplayResult(fallbackFieldResultStartIndex, fallbackFieldResultEndIndex);
                    characterIndex = fallbackFieldEndIndex;
                    continue;
                }

                if (textCharacter.Character == '\a') {
                    LegacyDocParagraphFormat paragraphFormat = GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset);
                    LegacyDocCharacterFormat paragraphMarkFormat = GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset);
                    if (paragraphMarkFormat.HasFormatting) {
                        paragraphFormat = paragraphFormat.WithParagraphMarkFormat(paragraphMarkFormat);
                    }

                    if (paragraphFormat.IsInTable == true) {
                        if (paragraphFormat.IsTableTerminatingParagraph == true) {
                            AddCurrentTableRow(paragraphFormat, textCharacter.CharacterPosition + 1);
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
                    LegacyDocParagraphFormat paragraphFormat = GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset);
                    if (format.HasFormatting) {
                        paragraphFormat = paragraphFormat.WithParagraphMarkFormat(format);
                    }

                    if (paragraphFormat.IsInTable == true && paragraphFormat.IsTableTerminatingParagraph != true) {
                        AddCurrentTextAsTableCellParagraph(paragraphFormat);
                    } else if (inTable) {
                        FlushTable(GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset), textCharacter.CharacterPosition + 1);
                    } else {
                        AddCurrentTextAsParagraph(paragraphFormat);
                    }

                    bodyText.Append('\r');
                    AddSectionBreaksAtBodyBoundary(textCharacter.CharacterPosition + 1);
                    currentParagraphStartCharacter = textCharacter.CharacterPosition + 1;
                    continue;
                }

                AppendRunCharacter(normalized.Value, format, textCharacter.CharacterPosition);
                bodyText.Append(normalized.Value);
            }

            if (inTable) {
                FlushTable(LegacyDocParagraphFormat.Default, characters.Count == 0 ? 0 : characters[characters.Count - 1].CharacterPosition + 1);
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
                    if (IsInsideActiveTable()) {
                        ReportUnprojectedSectionBoundary(sections[nextSectionIndex].StartCharacter);
                    } else if (characterPosition < characters.Count) {
                        _bodyBlocks.Add(new LegacyDocSectionBreakBlock(sections[nextSectionIndex].Format));
                    }

                    nextSectionIndex++;
                }
            }

            bool IsInsideActiveTable() =>
                inTable
                || tableRows.Count > 0
                || currentTableRow.Count > 0
                || currentTableCellParagraphs.Count > 0;

            void ReportRemainingUnprojectedSectionBoundaries() {
                while (nextSectionIndex < sections.Count) {
                    if (sections[nextSectionIndex].StartCharacter < characters.Count) {
                        ReportUnprojectedSectionBoundary(sections[nextSectionIndex].StartCharacter);
                    }

                    nextSectionIndex++;
                }
            }

            void ReportUnprojectedSectionBoundary(int characterPosition) {
                if (reportedUnprojectedSectionBoundary) {
                    return;
                }

                reportedUnprojectedSectionBoundary = true;
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.Section,
                    "DOC-MULTIPLE-SECTIONS-PRESENT",
                    $"The legacy DOC contains a section boundary at character position {characterPosition} that does not align with a supported body-block boundary. That section is preserved in the source file but is not projected into the OfficeIMO document.",
                    detailCode: "Fib:PlcfSed"),
                    reportUnsupportedFeatures);
            }

            void AppendHyperlinkResult(LegacyDocHyperlinkTarget hyperlinkTarget, int resultStartIndex, int resultEndIndex) {
                for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                    if (LegacyDocField.TryReadEquationField(
                        characters,
                        resultIndex,
                        out string equationInstruction,
                        out int equationResultStartIndex,
                        out int equationResultEndIndex,
                        out int equationFieldEndIndex) &&
                        equationFieldEndIndex < resultEndIndex) {
                        FlushRun();
                        LegacyDocCharacterFormat equationFormat = LegacyDocCharacterFormat.Default;
                        var equationPositions = new List<int>();
                        var equationText = new System.Text.StringBuilder();
                        foreach (int equationResultIndex in LegacyDocField.EnumerateVisibleResultIndexes(characters, equationResultStartIndex, equationResultEndIndex)) {
                            LegacyDocTextCharacter equationCharacter = characters[equationResultIndex];
                            char? normalizedEquationCharacter = NormalizeBodyCharacter(equationCharacter.Character);
                            if (normalizedEquationCharacter == null) continue;
                            if (equationPositions.Count == 0) {
                                equationFormat = GetFormatForFileOffset(formattingRanges, equationCharacter.FileOffset);
                            }
                            equationText.Append(normalizedEquationCharacter.Value);
                            equationPositions.Add(equationCharacter.CharacterPosition);
                            bodyText.Append(normalizedEquationCharacter.Value);
                        }
                        currentRuns.Add(LegacyDocTextRunFactory.CreateFieldRun(
                            equationText.ToString(),
                            LegacyDocFieldKind.Equation,
                            equationInstruction,
                            equationFormat,
                            equationPositions,
                            hyperlinkTarget));
                        resultIndex = equationFieldEndIndex;
                        continue;
                    }

                    if (characters[resultIndex].Character == LegacyDocField.Begin &&
                        LegacyDocField.TryReadNestedFieldResult(
                            characters,
                            resultIndex,
                            resultEndIndex,
                            out int nestedResultStartIndex,
                            out int nestedResultEndIndex,
                            out int nestedFieldEndIndex)) {
                        foreach (int visibleResultIndex in LegacyDocField.EnumerateVisibleResultIndexes(
                            characters,
                            nestedResultStartIndex,
                            nestedResultEndIndex)) {
                            LegacyDocTextCharacter visibleCharacter = characters[visibleResultIndex];
                            char? normalizedVisibleCharacter = NormalizeBodyCharacter(visibleCharacter.Character);
                            if (normalizedVisibleCharacter == null) continue;
                            LegacyDocCharacterFormat visibleFormat = GetFormatForFileOffset(formattingRanges, visibleCharacter.FileOffset);
                            AppendRunCharacter(normalizedVisibleCharacter.Value, visibleFormat, visibleCharacter.CharacterPosition, hyperlinkTarget);
                            bodyText.Append(normalizedVisibleCharacter.Value);
                        }
                        resultIndex = nestedFieldEndIndex;
                        continue;
                    }

                    LegacyDocTextCharacter resultCharacter = characters[resultIndex];
                    char? normalized = NormalizeBodyCharacter(resultCharacter.Character);
                    if (normalized == null) {
                        continue;
                    }

                    LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset);
                    AppendRunCharacter(normalized.Value, format, resultCharacter.CharacterPosition, hyperlinkTarget);
                    bodyText.Append(normalized.Value);
                }
            }

            void AppendPageNumberResult(int resultStartIndex, int resultEndIndex) {
                AppendFieldResult(LegacyDocFieldKind.Page, fieldInstruction: null, resultStartIndex, resultEndIndex);
            }

            void AppendFieldDisplayResult(int resultStartIndex, int resultEndIndex) {
                for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                    LegacyDocTextCharacter resultCharacter = characters[resultIndex];
                    char? normalized = NormalizeBodyCharacter(resultCharacter.Character);
                    if (normalized == null) {
                        continue;
                    }

                    LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset);
                    AppendRunCharacter(normalized.Value, format, resultCharacter.CharacterPosition);
                    bodyText.Append(normalized.Value);
                }
            }

            void AppendFieldResult(LegacyDocFieldKind fieldKind, string? fieldInstruction, int resultStartIndex, int resultEndIndex) {
                FlushRun();
                LegacyDocCharacterFormat format = LegacyDocCharacterFormat.Default;
                var positions = new List<int>();
                var resultText = new System.Text.StringBuilder();
                for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                    LegacyDocTextCharacter resultCharacter = characters[resultIndex];
                    char? normalized = NormalizeBodyCharacter(resultCharacter.Character);
                    if (normalized == null) {
                        continue;
                    }

                    if (positions.Count == 0) {
                        format = GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset);
                    }

                    resultText.Append(normalized.Value);
                    positions.Add(resultCharacter.CharacterPosition);
                }

                currentRuns.Add(LegacyDocTextRunFactory.CreateFieldRun(
                    fieldKind == LegacyDocFieldKind.Page ? string.Empty : resultText.ToString(),
                    fieldKind,
                    fieldInstruction,
                    format,
                    positions));
                if (inTable) {
                    justClosedCell = false;
                }
            }

            void AppendRunCharacter(char character, LegacyDocCharacterFormat format, int characterPosition, LegacyDocHyperlinkTarget hyperlinkTarget = default) {
                if (!hasCurrentRun
                    || !format.Equals(currentFormat)
                    || hyperlinkTarget != currentHyperlinkTarget) {
                    FlushRun();
                    currentFormat = format;
                    currentHyperlinkTarget = hyperlinkTarget;
                    hasCurrentRun = true;
                }

                runText.Append(character);
                runCharacterPositions.Add(characterPosition);
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
                    currentFormat.NoProof,
                    currentFormat.Caps,
                    currentFormat.VerticalPosition,
                    currentFormat.Underline,
                    currentFormat.Highlight,
                    currentFormat.FontSizeHalfPoints,
                    currentFormat.ColorHex,
                    currentFormat.FontFamily,
                    runCharacterPositions,
                    currentHyperlinkTarget.Uri,
                    currentHyperlinkTarget.Anchor,
                    specified: currentFormat.Specified,
                    characterSpacingTwips: currentFormat.CharacterSpacingTwips,
                    language: currentFormat.Language,
                    eastAsiaLanguage: currentFormat.EastAsiaLanguage,
                    revision: currentFormat.Revision));
                runText.Clear();
                runCharacterPositions.Clear();
                currentHyperlinkTarget = default;
            }

            void AddCurrentTextAsParagraph(LegacyDocParagraphFormat paragraphFormat) {
                FlushRun();
                IReadOnlyList<LegacyDocTextRun> runs = currentRuns.ToArray();
                _paragraphTextRuns.Add(runs);
                _paragraphFormats.Add(paragraphFormat);
                _paragraphs.Add(string.Concat(runs.Select(run => run.Text)));
                int paragraphEndCharacter = Math.Max(currentParagraphStartCharacter + runs.Sum(run => run.Text.Length), GetRunEndCharacter(runs));
                _bodyBlocks.Add(new LegacyDocParagraphBlock(
                    runs,
                    paragraphFormat,
                    currentParagraphStartCharacter,
                    paragraphEndCharacter,
                    bookmarkProjection.ExtractProjectedParagraphBookmarks(currentParagraphStartCharacter, paragraphEndCharacter)));
                currentRuns.Clear();
                hasCurrentRun = false;
            }

            void AddCurrentTextAsTableCellParagraph(LegacyDocParagraphFormat paragraphFormat) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    justClosedCell = false;
                    currentTableStartCharacter = currentParagraphStartCharacter;
                    currentTableRowStartCharacter = currentParagraphStartCharacter;
                }

                currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                currentRuns.Clear();
                hasCurrentRun = false;
                justClosedCell = false;
            }

            void AddCurrentTextAsTableCell(LegacyDocParagraphFormat paragraphFormat, bool allowHeuristicRowTerminator) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    justClosedCell = false;
                    currentTableStartCharacter = currentParagraphStartCharacter;
                    currentTableRowStartCharacter = currentParagraphStartCharacter;
                }

                if (allowHeuristicRowTerminator && currentRuns.Count == 0 && justClosedCell) {
                    if (currentTableRow.Count > 0) {
                        tableRows.Add(new LegacyDocTableRow(currentTableRow.ToArray(), bookmarksBefore: currentTableRowBoundaryBookmarks ?? ExtractCurrentTableRowBoundaryBookmarks()));
                        currentTableRow.Clear();
                    }

                    currentTableRowBoundaryBookmarks = null;
                    justClosedCell = false;
                    return;
                }

                if (currentRuns.Count > 0 || currentTableCellParagraphs.Count == 0) {
                    currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                }

                currentTableRow.Add(new LegacyDocTableCell(currentTableCellParagraphs.ToArray()));
                currentTableCellParagraphs.Clear();
                currentRuns.Clear();
                hasCurrentRun = false;
                justClosedCell = true;
            }

            void AddCurrentTableRow(LegacyDocParagraphFormat paragraphFormat, int rowEndCharacter) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    currentTableStartCharacter = currentParagraphStartCharacter;
                    currentTableRowStartCharacter = currentParagraphStartCharacter;
                }

                if (currentRuns.Count > 0 || currentTableCellParagraphs.Count > 0 || (!justClosedCell && currentTableRow.Count == 0)) {
                    if (currentRuns.Count > 0 || currentTableCellParagraphs.Count == 0) {
                        currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                    }

                    currentTableRow.Add(new LegacyDocTableCell(currentTableCellParagraphs.ToArray()));
                    currentTableCellParagraphs.Clear();
                    currentRuns.Clear();
                }

                if (currentTableRow.Count > 0) {
                    tableRows.Add(new LegacyDocTableRow(
                        currentTableRow.ToArray(),
                        paragraphFormat.TableCellWidthsTwips,
                        paragraphFormat.TableLeftIndentTwips,
                        paragraphFormat.TableRowHeightTwips,
                        paragraphFormat.TableRowHeightIsExact,
                        paragraphFormat.TableRowCantSplit,
                        paragraphFormat.TableRowIsHeader,
                        paragraphFormat.TableAlignment,
                        paragraphFormat.TableCellHorizontalMerges,
                        paragraphFormat.TableCellVerticalMerges,
                        paragraphFormat.TableCellVerticalAlignments,
                        paragraphFormat.TableCellTextDirections,
                        paragraphFormat.TableCellFitTexts,
                        paragraphFormat.TableCellNoWraps,
                        paragraphFormat.TableCellHideMarks,
                        paragraphFormat.GetTableCellMarginsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellShadingsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellBordersForCellCount(currentTableRow.Count),
                        paragraphFormat.DefaultTableCellSpacingTwips,
                        paragraphFormat.TablePreferredWidth,
                        paragraphFormat.TableAutofit,
                        currentTableRowBoundaryBookmarks ?? ExtractCurrentTableRowBoundaryBookmarks()));
                    currentTableRow.Clear();
                }

                hasCurrentRun = false;
                justClosedCell = false;
                currentTableRowBoundaryBookmarks = null;
                currentTableRowStartCharacter = rowEndCharacter;
            }

            void FlushTable(LegacyDocParagraphFormat paragraphFormat, int tableEndCharacter) {
                FlushRun();
                if (currentRuns.Count > 0 || currentTableCellParagraphs.Count > 0) {
                    if (currentRuns.Count > 0 || currentTableCellParagraphs.Count == 0) {
                        currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                    }

                    currentTableRow.Add(new LegacyDocTableCell(currentTableCellParagraphs.ToArray()));
                    currentTableCellParagraphs.Clear();
                    currentRuns.Clear();
                }

                if (currentTableRow.Count > 0) {
                    tableRows.Add(new LegacyDocTableRow(
                        currentTableRow.ToArray(),
                        paragraphFormat.TableCellWidthsTwips,
                        paragraphFormat.TableLeftIndentTwips,
                        paragraphFormat.TableRowHeightTwips,
                        paragraphFormat.TableRowHeightIsExact,
                        paragraphFormat.TableRowCantSplit,
                        paragraphFormat.TableRowIsHeader,
                        paragraphFormat.TableAlignment,
                        paragraphFormat.TableCellHorizontalMerges,
                        paragraphFormat.TableCellVerticalMerges,
                        paragraphFormat.TableCellVerticalAlignments,
                        paragraphFormat.TableCellTextDirections,
                        paragraphFormat.TableCellFitTexts,
                        paragraphFormat.TableCellNoWraps,
                        paragraphFormat.TableCellHideMarks,
                        paragraphFormat.GetTableCellMarginsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellShadingsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellBordersForCellCount(currentTableRow.Count),
                        paragraphFormat.DefaultTableCellSpacingTwips,
                        paragraphFormat.TablePreferredWidth,
                        paragraphFormat.TableAutofit,
                        currentTableRowBoundaryBookmarks ?? ExtractCurrentTableRowBoundaryBookmarks()));
                    currentTableRow.Clear();
                }

                if (tableRows.Count > 0) {
                    int tableStartCharacter = currentTableStartCharacter ?? GetTableStartCharacter(tableRows);
                    _bodyBlocks.Add(new LegacyDocTableBlock(
                        tableRows.ToArray(),
                        tableStartCharacter,
                        tableEndCharacter,
                        bookmarkProjection.ExtractUnprojectedBlockBookmarks(tableStartCharacter, tableEndCharacter)));
                    tableRows.Clear();
                }

                hasCurrentRun = false;
                inTable = false;
                justClosedCell = false;
                currentTableStartCharacter = null;
                currentTableRowStartCharacter = null;
                currentTableRowBoundaryBookmarks = null;
            }

            LegacyDocTableCellParagraph CreateCurrentTableCellParagraph(LegacyDocParagraphFormat paragraphFormat) {
                IReadOnlyList<LegacyDocTextRun> runs = currentRuns.ToArray();
                int paragraphStartCharacter = GetRunStartCharacter(runs);
                int paragraphEndCharacter = GetRunEndCharacter(runs);
                if (currentTableRowBoundaryBookmarks == null
                    && currentTableRowStartCharacter.HasValue
                    && paragraphStartCharacter == currentTableRowStartCharacter.Value) {
                    currentTableRowBoundaryBookmarks = ExtractCurrentTableRowBoundaryBookmarks();
                }

                return new LegacyDocTableCellParagraph(
                    runs,
                    paragraphFormat,
                    paragraphStartCharacter,
                    paragraphEndCharacter,
                    bookmarkProjection.ExtractProjectedParagraphBookmarks(paragraphStartCharacter, paragraphEndCharacter));
            }

            IReadOnlyList<LegacyDocBookmark> ExtractCurrentTableRowBoundaryBookmarks() {
                int rowStartCharacter = currentTableRowStartCharacter ?? currentTableStartCharacter ?? currentParagraphStartCharacter;
                return rowStartCharacter != currentTableStartCharacter
                    ? bookmarkProjection.ExtractZeroLengthBoundaryBookmarks(rowStartCharacter)
                    : Array.Empty<LegacyDocBookmark>();
            }

            int GetTableStartCharacter(IReadOnlyList<LegacyDocTableRow> rows) {
                foreach (LegacyDocTableRow row in rows) {
                    foreach (LegacyDocTableCell cell in row.Cells) {
                        foreach (LegacyDocTableCellParagraph paragraph in cell.Paragraphs) {
                            return paragraph.StartCharacter;
                        }
                    }
                }

                return currentParagraphStartCharacter;
            }

            int GetRunStartCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
                foreach (LegacyDocTextRun run in runs) {
                    if (run.CharacterPositions.Count > 0) {
                        return run.CharacterPositions[0];
                    }
                }

                return currentParagraphStartCharacter;
            }

            int GetRunEndCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
                for (int index = runs.Count - 1; index >= 0; index--) {
                    IReadOnlyList<int> positions = runs[index].CharacterPositions;
                    if (positions.Count > 0) {
                        return positions[positions.Count - 1] + 1;
                    }
                }

                return currentParagraphStartCharacter;
            }
        }

        private static char? NormalizeBodyCharacter(char character) {
            switch (character) {
                case '\0':
                case '\a':
                    return null;
                case LegacyDocFootnoteReader.FootnoteReferenceCharacter:
                case LegacyDocCommentReader.CommentReferenceCharacter:
                case LegacyDocSpecialCharacters.TextWrappingBreak:
                case LegacyDocSpecialCharacters.PageBreak:
                case LegacyDocSpecialCharacters.ColumnBreak:
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

        internal void AddUnsupportedFeature(LegacyDocUnsupportedFeature feature, bool reportDiagnostic = true) {
            _unsupportedFeatures.Add(feature);
            if (reportDiagnostic) {
                AddWarning(feature.Code, feature.Description);
            }
        }
    }
}
