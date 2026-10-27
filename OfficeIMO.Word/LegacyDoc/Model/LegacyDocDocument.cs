using OfficeIMO.Core.Internal;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Diagnostics;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Neutral legacy binary Word document model for the supported import subset.
    /// </summary>
    public sealed partial class LegacyDocDocument {
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

        internal LegacyDocNumbering Numbering { get; private set; } = LegacyDocNumbering.Empty;

        internal LegacyDocSectionFormat SectionFormat { get; private set; } = LegacyDocSectionFormat.Default;

        internal IReadOnlyList<LegacyDocSection> Sections { get; private set; } = Array.Empty<LegacyDocSection>();

        internal IReadOnlyList<LegacyDocHeaderFooterStory> HeaderFooterStories { get; private set; } = Array.Empty<LegacyDocHeaderFooterStory>();

        internal IReadOnlyList<LegacyDocFootnote> Footnotes { get; private set; } = Array.Empty<LegacyDocFootnote>();

        internal IReadOnlyList<LegacyDocEndnote> Endnotes { get; private set; } = Array.Empty<LegacyDocEndnote>();

        internal IReadOnlyList<LegacyDocComment> Comments { get; private set; } = Array.Empty<LegacyDocComment>();

        internal IReadOnlyList<LegacyDocTextBoxStory> TextBoxStories { get; private set; } = Array.Empty<LegacyDocTextBoxStory>();

        internal IReadOnlyList<LegacyDocBookmark> Bookmarks { get; private set; } = Array.Empty<LegacyDocBookmark>();

        internal bool DifferentOddAndEvenPages { get; private set; }

        internal bool MirrorMargins { get; private set; }

        internal bool GutterAtTop { get; private set; }

        internal bool NoColumnBalance { get; private set; }

        internal int? DefaultTabStop { get; private set; }

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
        private static LegacyDocDocument LoadCore(byte[] bytes, LegacyDocImportOptions? options) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            options ??= new LegacyDocImportOptions();
            options.Validate();
            if (bytes.Length > options.MaxInputBytes) {
                throw new InvalidDataException($"Legacy DOC input exceeds the configured maximum size ({options.MaxInputBytes} bytes).");
            }

            var document = new LegacyDocDocument();
            if (!OfficeCompoundFileReader.TryRead(bytes, new OfficeCompoundReadOptions(preserveResourceLimits: true), out OfficeCompoundFile? compoundFile, out string? compoundError)) {
                document.AddError("doc-compound-invalid", compoundError ?? "The OLE compound document could not be read.");
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
                AddError("doc-worddocument-missing", "The compound document does not contain a WordDocument stream.");
                return;
            }

            byte[] wordDocumentStream = wordDocumentStreamCandidate!;
            if (wordDocumentStream.Length > options.MaxInputBytes) {
                AddError("doc-worddocument-too-large", $"The WordDocument stream is {wordDocumentStream.Length} bytes, which exceeds the configured limit of {options.MaxInputBytes} bytes.");
                return;
            }

            LegacyDocFib fib;
            if (!LegacyDocFib.TryRead(wordDocumentStream, out fib, out string? fibError)) {
                AddError("doc-fib-invalid", fibError ?? "The WordDocument stream does not contain a supported File Information Block.");
                return;
            }

            if (fib.IsEncrypted) {
                AddError("doc-encrypted", "Password-protected binary .doc files are detected but are not imported by the dependency-free reader.");
                return;
            }

            LegacyDocOleDocumentPropertyReader.AddDocumentProperties(compoundFile, this, options);

            string tableStreamName = fib.UsesOneTableStream ? "1Table" : "0Table";
            if (!TryGetRootStream(compoundFile, tableStreamName, out byte[]? tableStreamCandidate)) {
                string alternateName = fib.UsesOneTableStream ? "0Table" : "1Table";
                if (TryGetRootStream(compoundFile, alternateName, out tableStreamCandidate)) {
                    AddWarning("doc-table-stream-fallback", $"The FIB requested {tableStreamName}, but only {alternateName} was present. The available table stream was used.");
                } else {
                    AddError("doc-table-stream-missing", $"The compound document does not contain the {tableStreamName} table stream.");
                    return;
                }
            }

            byte[] tableStream = tableStreamCandidate!;
            DifferentOddAndEvenPages = ReadDopFacingPagesFlag(tableStream, fib);
            ReadDopMarginSettings(tableStream, fib);
            EndnotePositionValues? dopEndnotePosition = ReadDopEndnotePlacement(tableStream, fib);

            if (!LegacyDocPieceTable.TryRead(wordDocumentStream, tableStream, fib, options.MaxDecodedCharacters, out LegacyDocTextContent textContent, out string? textError)) {
                AddError("doc-piece-table-invalid", textError ?? "The legacy DOC piece table could not be decoded.");
                return;
            }

            IReadOnlyList<string> fontFamilies = LegacyDocFontTableReader.ReadFontFamilies(tableStream, fib, out string? fontTableWarning);
            if (fontTableWarning != null) {
                AddWarning("doc-font-table-invalid", fontTableWarning);
            }

            IReadOnlyList<string> revisionAuthors = LegacyDocRevisionAuthorReader.Read(tableStream, fib, out string? revisionAuthorWarning);
            if (revisionAuthorWarning != null) {
                AddWarning("doc-revision-authors-invalid", revisionAuthorWarning);
            }

            StyleSheet = LegacyDocStyleSheet.Read(tableStream, fib, fontFamilies, out string? styleSheetWarning);
            if (styleSheetWarning != null) {
                AddWarning("doc-stylesheet-invalid", styleSheetWarning);
            }

            Numbering = LegacyDocNumberingReader.Read(tableStream, fib, fontFamilies, out string? numberingWarning);
            if (numberingWarning != null) {
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(LegacyDocUnsupportedFeatureKind.Numbering,
                    "doc-numbering-invalid", numberingWarning, detailCode: "PlfLst/PlfLfo"), options.ReportUnsupportedContent);
            }

            Sections = ApplyDopEndnotePlacement(
                LegacyDocSectionFormattingReader.ReadSections(wordDocumentStream, tableStream, fib, out string? sectionFormattingWarning),
                fib,
                dopEndnotePosition);
            Sections = ApplyDopNoteSettings(Sections, tableStream, fib);
            SectionFormat = Sections.Count == 0 ? LegacyDocSectionFormat.Default : Sections[0].Format;
            if (sectionFormattingWarning != null) {
                AddWarning("doc-sepx-invalid", sectionFormattingWarning);
            }

            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges = LegacyDocCharacterFormattingReader.ReadCharacterFormatting(
                wordDocumentStream,
                tableStream,
                fib,
                fontFamilies,
                revisionAuthors,
                out string? formattingWarning);
            if (formattingWarning != null) {
                AddWarning("doc-chpx-invalid", formattingWarning);
            }

            byte[] dataStream = TryGetRootStream(compoundFile, "Data", out byte[]? dataStreamCandidate)
                ? dataStreamCandidate!
                : Array.Empty<byte>();
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges = LegacyDocParagraphFormattingReader.ReadParagraphFormatting(
                wordDocumentStream, tableStream, fib, out string? paragraphFormattingWarning, dataStream, options);
            if (paragraphFormattingWarning != null) {
                AddWarning("doc-papx-invalid", paragraphFormattingWarning);
            }

            if (numberingWarning == null && paragraphFormattingRanges.Select(item => item.Format)
                .Concat(StyleSheet.ParagraphStyles.Select(item => item.ParagraphFormat))
                .Any(format => !Numbering.ContainsReference(format))) {
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(LegacyDocUnsupportedFeatureKind.Numbering,
                    "doc-numbering-reference-invalid", "A paragraph or style references a missing native list instance or level.",
                    detailCode: "sprmPIlfo/sprmPIlvl"), options.ReportUnsupportedContent);
            }

            Bookmarks = LegacyDocBookmarkReader.Read(tableStream, fib, out string? bookmarkWarning);
            if (bookmarkWarning != null) {
                AddWarning("doc-bookmark-plc-invalid", bookmarkWarning);
            }
            var bookmarkProjection = new LegacyDocBookmarkProjectionTracker(Bookmarks);

            AddUnsupportedParagraphFormattingFeaturesIfPresent(paragraphFormattingRanges, options.ReportUnsupportedContent);

            LegacyDocPictureReader.LegacyDocPictureReadResult pictures = LegacyDocPictureReader.Read(
                dataStream,
                textContent.AllCharacters,
                formattingRanges,
                fib.CcpText + fib.CcpFtn + fib.CcpHdd + fib.CcpAtn + fib.CcpEdn,
                options.MaxDecodedImageBytes);
            if (pictures.Warning != null) {
                AddWarning("doc-picture-data-invalid", pictures.Warning);
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
                AddWarning("doc-footnote-plc-invalid", footnoteWarning);
            }

            Endnotes = LegacyDocFootnoteReader.ReadEndnotes(tableStream, textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition, out string? endnoteWarning);
            if (endnoteWarning != null) {
                AddWarning("doc-endnote-plc-invalid", endnoteWarning);
            }

            Comments = LegacyDocCommentReader.Read(tableStream, textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition, out string? commentWarning);
            if (commentWarning != null) {
                AddWarning("doc-comment-plc-invalid", commentWarning);
            }

            TextBoxStories = LegacyDocTextBoxStoryReader.Read(textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition);
            AddKnownUnsupportedFeatureDiagnostics(
                compoundFile,
                tableStream,
                fib,
                pictures.FullyProjectsDataStream,
                options.ReportUnsupportedContent);

            HeaderFooterStories = LegacyDocHeaderFooterReader.Read(tableStream, textContent, fib, formattingRanges, paragraphFormattingRanges, bookmarkProjection, pictures.PicturesByCharacterPosition, out string? headerFooterWarning);
            if (headerFooterWarning != null) {
                AddWarning("doc-plcfhdd-invalid", headerFooterWarning);
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.HeaderFooter,
                    "doc-header-footer-stories-present",
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
                "doc-macros-present",
                "The legacy DOC contains VBA project storage. Macros are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:VbaProjectStorage");

            AddCompoundFeatureIfPresent(
                compoundFile,
                entry => entry.Path.IndexOf("ObjectPool", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocCompoundFeatureKind.OleObject,
                "doc-ole-objects-present",
                "The legacy DOC contains embedded OLE object storage. Embedded objects are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:OleObjectStorage");

            AddCompoundFeatureIfPresent(
                compoundFile,
                entry => entry.Path.IndexOf("ActiveX", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.IndexOf("OCX", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocCompoundFeatureKind.ActiveXControl,
                "doc-activex-controls-present",
                "The legacy DOC contains ActiveX control storage. ActiveX controls are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:ActiveXControlStorage");

            AddCompoundFeatureIfPresent(
                compoundFile,
                entry => string.Equals(entry.Name.TrimStart('\u0001'), "Ole10Native", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.IndexOf("Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.IndexOf("Package", StringComparison.OrdinalIgnoreCase) >= 0,
                LegacyDocCompoundFeatureKind.EmbeddedPackage,
                "doc-embedded-packages-present",
                "The legacy DOC contains embedded package payload storage. Embedded packages are preserved in the source file but are not projected into the OfficeIMO document.",
                "Compound:EmbeddedPackageStorage");

            AddDataStreamCompoundFeatureIfPresent(compoundFile, pictureDataFullyProjected);

            AddCompoundFeatureIfPresent(
                compoundFile,
                IsDigitalSignatureEntry,
                LegacyDocCompoundFeatureKind.DigitalSignature,
                "doc-digital-signature-present",
                "The legacy DOC contains digital-signature storage. Rewriting the document invalidates the existing signature.",
                "Compound:DigitalSignature");
            AddUnsupportedCompoundEntryIfPresent(
                compoundFile,
                IsDigitalSignatureEntry,
                LegacyDocUnsupportedFeatureKind.DigitalSignature,
                "doc-digital-signature-present",
                "The legacy DOC contains digital-signature storage. Saving is blocked by default because rewriting the document invalidates the existing signature.",
                "Compound:DigitalSignature",
                reportDiagnostic);

            if (fib.CcpHdd > 0 && fib.LcbPlcfHdd == 0) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpHdd,
                    LegacyDocUnsupportedFeatureKind.HeaderFooter,
                    "doc-header-footer-stories-present",
                    "The legacy DOC contains header or footer story text without a supported header/footer story PLC. Headers and footers are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpHdd",
                    reportDiagnostic);
            }
            if (!LegacyDocFootnoteReader.HasReadableFootnoteTables(tableStream, fib)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpFtn,
                    LegacyDocUnsupportedFeatureKind.Footnote,
                    "doc-footnote-stories-present",
                    "The legacy DOC contains footnote story text. Footnotes are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpFtn",
                    reportDiagnostic);
            }
            if (!LegacyDocFootnoteReader.HasReadableEndnoteTables(tableStream, fib)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpEdn,
                    LegacyDocUnsupportedFeatureKind.Endnote,
                    "doc-endnote-stories-present",
                    "The legacy DOC contains endnote story text. Endnotes are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpEdn",
                    reportDiagnostic);
            }
            if (!LegacyDocCommentReader.HasReadableCommentTables(tableStream, fib)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpAtn,
                    LegacyDocUnsupportedFeatureKind.Comment,
                    "doc-comment-stories-present",
                    "The legacy DOC contains comment or annotation story text. Comments are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpAtn",
                    reportDiagnostic);
            }
            ReadRevisionTrackingState(tableStream, fib);
            if (fib.CcpTxbx > 0 && !TextBoxStories.Any(story => !story.IsHeaderFooterTextBox)) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpTxbx,
                    LegacyDocUnsupportedFeatureKind.TextBox,
                    "doc-textbox-stories-present",
                    "The legacy DOC contains text box story text that could not be read from the piece table. Text boxes are preserved in the source file but are not projected into the OfficeIMO document.",
                    "Fib:CcpTxbx",
                    reportDiagnostic);
            }

            if (fib.CcpHdrTxbx > 0) {
                AddUnsupportedStoryFeatureIfPresent(
                    fib.CcpHdrTxbx,
                    LegacyDocUnsupportedFeatureKind.TextBox,
                    "doc-header-textbox-stories-present",
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
                    "doc-fast-save-present",
                    "The legacy DOC is marked as fast-saved or complex. Fast-save deltas are preserved in the source file but are not projected into the OfficeIMO document.",
                    detailCode: "Fib:FComplex"),
                    reportDiagnostic);
            }

            if (fib.QuickSaveCount > 0 && reportDiagnostic) {
                AddWarning(
                    "doc-quick-save-history-present",
                    $"The legacy DOC reports {fib.QuickSaveCount} quick-save revision(s). The projected document content is readable; quick-save history is not represented as editable revision history.");
            }

            if (fib.HasPictures && pictureDataFullyProjected) {
                AddInfo(
                    "doc-pictures-projected",
                    "Supported inline picture payloads were projected from the legacy DOC Data stream into editable Open XML images.");
            } else if (fib.HasPictures) {
                _preservedFeatures.Add(new LegacyDocPreservedFeature(
                    LegacyDocPreservedFeatureKind.Picture,
                    "doc-pictures-present",
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
                "doc-binary-data-stream-present",
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
                        "doc-nested-tables-present",
                        BuildNestedTableDescription(format),
                        detailCode: "PAPX:sprmPItap"),
                        reportUnsupportedFeatures);
                    reportedNestedTable = true;
                }

                if (format.HasMergedTableCells) {
                    AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                        LegacyDocUnsupportedFeatureKind.MergedTableCell,
                        "doc-merged-table-cells-present",
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
                "doc-bookmark-range-present",
                $"The legacy DOC contains bookmark '{bookmark.Name}' at character range {bookmark.StartCharacter}-{bookmark.EndCharacter} outside the currently supported body, table-cell, header/footer, footnote, and endnote paragraph bookmark projection. The bookmark is preserved in the source file but is not projected into the OfficeIMO document.",
                "Fib:PlcfBkf"));
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
