using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int FibLength = 0x1AA;
        private const int TextOffset = 0x800;
        private const int OleSectorSize = 512;
        private const int OleMiniStreamCutoffSize = 4096;
        private const int ClxLength = 21;
        private const int SedLength = 12;
        private const uint CompressedTextFlag = 0x40000000;
        private const ushort SprmCFBold = 0x0835;
        private const ushort SprmCFItalic = 0x0836;
        private const ushort SprmCFStrike = 0x0837;
        private const ushort SprmCFOutline = 0x0838;
        private const ushort SprmCFShadow = 0x0839;
        private const ushort SprmCFImprint = 0x0854;
        private const ushort SprmCFSmallCaps = 0x083A;
        private const ushort SprmCFCaps = 0x083B;
        private const ushort SprmCFVanish = 0x083C;
        private const ushort SprmCFEmboss = 0x0858;
        private const ushort SprmCFSpec = 0x0855;
        private const ushort SprmCPicLocation = 0x6A03;
        private const ushort HasPicturesFibFlag = 0x0008;
        private const ushort TemplateFibFlag = 0x0001;
        private const ushort SprmCFNoProof = 0x0875;
        private const ushort SprmCHighlight = 0x2A0C;
        private const ushort SprmCKul = 0x2A3E;
        private const ushort SprmCDxaSpace = 0x8840;
        private const ushort SprmCCharScale = 0x4852;
        private const ushort SprmCIss = 0x2A48;
        private const ushort SprmCHps = 0x4A43;
        private const ushort SprmCRgLid0 = 0x486D;
        private const ushort SprmCRgLid1 = 0x486E;
        private const ushort SprmCRgFtc0 = 0x4A4F;
        private const ushort SprmCFDStrike = 0x2A53;
        private const ushort SprmCCv = 0x6870;
        private const ushort SprmCFRMarkDel = 0x0800;
        private const ushort SprmCFRMarkIns = 0x0801;
        private const ushort SprmCIbstRMark = 0x4804;
        private const ushort SprmCDttmRMark = 0x6805;
        private const ushort SprmCIbstRMarkDel = 0x4863;
        private const ushort SprmCDttmRMarkDel = 0x6864;
        private const ushort SprmPJc = 0x2461;
        private const ushort DefaultPcdFlags = 0x0310;
        private const ushort FootnotePcdFlags = 0x0330;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const ushort Word97FibVersion = 0x00C1;
        private const ushort Word97FibBackVersion = 0x00BF;
        private const ushort DefaultLanguageId = 0x0409;
        private const ushort FibRgW97WordCount = 0x000E;
        private const ushort FibRgLw97DwordCount = 0x0016;
        private const ushort OneTableStreamFlag = 0x0200;
        private const ushort ExtendedCharacterFlag = 0x1000;
        private const ushort DefaultFibFlags = 0x1200;
        private const int FcPlcfSedOffset = 0xCA;
        private const int LcbPlcfSedOffset = 0xCE;
        private const int FcPlcffndRefOffset = 0xAA;
        private const int LcbPlcffndRefOffset = 0xAE;
        private const int FcPlcffndTxtOffset = 0xB2;
        private const int LcbPlcffndTxtOffset = 0xB6;
        private const int FcPlcfandRefOffset = 0xBA;
        private const int LcbPlcfandRefOffset = 0xBE;
        private const int FcPlcfandTxtOffset = 0xC2;
        private const int LcbPlcfandTxtOffset = 0xC6;
        private const int FcPlcfendRefOffset = 0x20A;
        private const int LcbPlcfendRefOffset = 0x20E;
        private const int FcPlcfendTxtOffset = 0x212;
        private const int LcbPlcfendTxtOffset = 0x216;
        private const int FcStshfOffset = 0xA2;
        private const int LcbStshfOffset = 0xA6;
        private const int FcDopOffset = 0x192;
        private const int LcbDopOffset = 0x196;
        private const int FcSttbfAssocOffset = 0x19A;
        private const int LcbSttbfAssocOffset = 0x19E;
        private const int FcPlcfBtePapxOffset = 0x102;
        private const int LcbPlcfBtePapxOffset = 0x106;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;
        private const int FcSttbfBkmkOffset = 0x142;
        private const int LcbSttbfBkmkOffset = 0x146;
        private const int FcSttbfRMarkOffset = 0x232;
        private const int LcbSttbfRMarkOffset = 0x236;
        private const int FcRmdThreadingOffset = 0x38A;
        private const int LcbRmdThreadingOffset = 0x38E;
        private const int FcPlcfBkfOffset = 0x14A;
        private const int LcbPlcfBkfOffset = 0x14E;
        private const int FcPlcfBklOffset = 0x152;
        private const int LcbPlcfBklOffset = 0x156;
        private const int FcPlcfHddOffset = 0xF2;
        private const int LcbPlcfHddOffset = 0xF6;
        private const ushort NoteTextParagraphStyleIndex = 0x0023;
        private static readonly byte[] PlainParagraphPapx = { 0x00, 0x01, 0x00, 0x00 };
        private static readonly byte[] FootnoteTextParagraphPapx = { 0x00, 0x01, (byte)NoteTextParagraphStyleIndex, (byte)(NoteTextParagraphStyleIndex >> 8) };

        // The legacy DOC model carries WordprocessingML values as its own enums; the writer reads them as schema tokens.
        private static TModel? ToModel<TModel>(string? value) where TModel : struct, Enum =>
            value == null ? null : LegacyDocWordValues.TryParseXml<TModel>(value);

        private static byte? ToNfc(string format) =>
            ToModel<LegacyDocNumberFormat>(format) is { } model ? LegacyDocNumberFormatMapper.ToNfc(model) : null;

        private static void ThrowIfUnsupportedLegacyDocImportState(WordDocument document, WordSaveOptions? options) {
            bool hasLegacyDigitalSignature = document.LegacyDocCompoundFeatures.Any(feature =>
                feature.Kind == LegacyDocCompoundFeatureKind.DigitalSignature);
            if (hasLegacyDigitalSignature
                && (options?.SignatureMutationPolicy ?? OfficeSignatureMutationPolicy.BlockSave) == OfficeSignatureMutationPolicy.BlockSave) {
                throw new NotSupportedException(
                    "Native DOC saving is blocked because the imported legacy DOC contains digital-signature metadata. "
                    + "Rewriting the document invalidates that signature. Set WordSaveOptions.SignatureMutationPolicy to "
                    + "RemoveInvalidatedSignatures or PreserveSignatureMarkup and explicitly allow conversion loss to continue.");
            }

            if (document.SourceFormat != WordFileFormat.Doc
                || (document.LegacyDocUnsupportedFeatures.Count == 0
                    && document.LegacyDocPreservedFeatures.Count == 0
                    && document.LegacyDocCompoundFeatures.Count == 0)
                || options?.LossPolicy == OfficeConversionLossPolicy.Allow) {
                return;
            }

            string codes = string.Join(
                ", ",
                document.LegacyDocUnsupportedFeatures
                    .Select(feature => feature.Code)
                    .Concat(document.LegacyDocPreservedFeatures.Select(feature => feature.Code))
                    .Concat(document.LegacyDocCompoundFeatures.Select(feature => feature.Code))
                    .Where(code => !string.IsNullOrWhiteSpace(code))
                    .Distinct(StringComparer.Ordinal)
                    .Take(5));
            string detail = string.IsNullOrWhiteSpace(codes)
                ? "unsupported or preserve-only features"
                : $"unsupported or preserve-only features ({codes})";

            throw new NotSupportedException($"Native DOC saving is blocked because this document was imported from a legacy DOC with {detail}. Save as DOCX after reviewing LegacyDocUnsupportedFeatures, LegacyDocPreservedFeatures, and LegacyDocCompoundFeatures, or remove and recreate the unsupported content before saving as DOC.");
        }

        private static byte[] PadToRegularOleStream(byte[] bytes) {
            if (bytes.Length >= OleMiniStreamCutoffSize) {
                return bytes;
            }

            byte[] padded = new byte[OleMiniStreamCutoffSize];
            Buffer.BlockCopy(bytes, 0, padded, 0, bytes.Length);
            return padded;
        }

        private static LegacyDocWritableBody BuildBody(WordDocument document, LegacyDocSourcePackage source) {
            LegacyDocElement body = source.Body;
            if (body.IsNone) {
                throw new InvalidOperationException("Document couldn't be saved as legacy DOC because the document body is missing.");
            }

            ThrowIfUnsupportedDocumentParts(source);

            var text = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var paragraphFormats = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            var bodyBoundaries = new BodyBlockBoundaryCache();
            var pictures = new LegacyDocWritablePictures(document, source);
            LegacyDocWritableStyleSheet styleSheet = CreateWritableStyleSheet(source, body);
            LegacyDocWritableFootnotes footnotes = ReadSupportedFootnotes(source, pictures, styleSheet.StyleIndexes);
            LegacyDocWritableEndnotes endnotes = ReadSupportedEndnotes(source, pictures, styleSheet.StyleIndexes);
            LegacyDocWritableComments comments = ReadSupportedComments(source, pictures, styleSheet.StyleIndexes);
            LegacyDocTableStyleDefinitions tableStyleDefinitions = ReadTableStyleDefinitions(source);
            LegacyDocSectionFormat finalSectionFormat = LegacyDocSectionFormat.Default;
            var sections = new List<LegacyDocWritableSection>();
            int bodyContentCount = 0;
            foreach (LegacyDocElement child in body.Elements()) {
                AppendBodyChild(
                    text,
                    runs,
                    paragraphFormats,
                    bookmarks,
                    bodyBoundaries,
                    child,
                    source,
                    pictures,
                    styleSheet.StyleIndexes,
                    tableStyleDefinitions,
                    footnotes,
                    endnotes,
                    sections,
                    ref finalSectionFormat,
                    ref bodyContentCount,
                    "body");
            }

            // A final table still needs a body paragraph. A following authored
            // paragraph already supplies that boundary and must not get an extra blank.
            if (bodyContentCount == 0 || text[text.Length - 1] == '\a') {
                text.Append('\r');
            }

            AddSection(sections, text.Length, finalSectionFormat);
            footnotes.ThrowIfUnreferencedFootnotesRemain();
            endnotes.ThrowIfUnreferencedEndnotesRemain();
            comments.BindBodyReferences(body, text.ToString());
            bool hasNoteReferences = footnotes.HasReferences || endnotes.HasReferences;
            LegacyDocWritableHeaderFooterStories headerFooterStories = BuildHeaderFooterStories(source, pictures, hasNoteReferences, styleSheet.StyleIndexes);
            int terminalCharacterPadding = hasNoteReferences ? 1 : 0;
            LegacyDocWritableFootnoteStories footnoteStories = footnotes.CreateStories(text.Length, headerFooterStories.Text.Length, terminalCharacterPadding);
            bookmarks.AddRange(footnoteStories.Bookmarks, text.Length);
            bookmarks.AddRange(headerFooterStories.Bookmarks, text.Length + footnoteStories.Text.Length);
            LegacyDocWritableCommentStories commentStories = comments.CreateStories();
            int commentStoryStart = text.Length + footnoteStories.Text.Length + headerFooterStories.Text.Length;
            bookmarks.AddRange(commentStories.Bookmarks, commentStoryStart);
            LegacyDocWritableEndnoteStories endnoteStories = endnotes.CreateStories(
                text.Length,
                footnoteStories.Text.Length,
                headerFooterStories.Text.Length,
                commentStories.Text.Length,
                terminalCharacterPadding);
            bookmarks.AddRange(endnoteStories.Bookmarks, commentStoryStart + commentStories.Text.Length);
            LegacyDocElement settings = source.Settings;
            LegacyDocElement compatibility = settings.Child("compat");
            LegacyDocElement splitMark = compatibility.Child("splitPgBreakAndParaMark");
            if (splitMark.Exists && !IsOnOffEnabled(splitMark)) {
                throw new NotSupportedException("Native DOC layout always moves a trailing page-break paragraph mark onto the next page. Save as DOCX to retain SplitPageBreakAndParagraphMark=false.");
            }
            LegacyDocElement numberingTab = compatibility.Child("doNotUseIndentAsNumberingTabStop");
            if (numberingTab.Exists && !IsOnOffEnabled(numberingTab)) {
                throw new NotSupportedException("Native DOC numbering uses an authored or default tab stop rather than the hanging indent. Save as DOCX to retain DoNotUseIndentAsNumberingTabStop=false.");
            }
            bool trackRevisions = AnyOnOffEnabled(settings, "trackRevisions");
            bool lockRevisionTracking = IsLockedRevisionTracking(settings);
            var writableBody = new LegacyDocWritableBody(
                text.ToString(),
                runs,
                paragraphFormats,
                bookmarks.Create(),
                sections,
                styleSheet,
                source.Numbering,
                footnoteStories,
                endnoteStories,
                headerFooterStories,
                commentStories,
                pictures.DataBytes,
                pictures.HasPictures,
                HasEvenAndOddHeaders(source),
                AnyOnOffEnabled(settings, "mirrorMargins"),
                AnyOnOffEnabled(settings, "gutterAtTop"),
                AnyOnOffEnabled(compatibility, "noColumnBalance"),
                checked((ushort)(ReadTwipValue(settings.Child("defaultTabStop").Val, 720, "default tab interval") ?? 720)),
                ReadDocumentEndnotePosition(settings),
                trackRevisions || lockRevisionTracking,
                lockRevisionTracking);
            ValidateNativeNumberingReferences(source, styleSheet.StyleIndexes);
            return writableBody;
        }

        /// <summary>Whether any child of <paramref name="parent"/> with this name is an enabled on/off property.</summary>
        private static bool AnyOnOffEnabled(LegacyDocElement parent, string localName) {
            foreach (LegacyDocElement element in parent.Elements(localName)) {
                if (IsOnOffEnabled(element)) return true;
            }
            return false;
        }

        private static bool HasEvenAndOddHeaders(LegacyDocSourcePackage source) => AnyOnOffEnabled(source.Settings, "evenAndOddHeaders");

        private static bool IsLockedRevisionTracking(LegacyDocElement settings) {
            LegacyDocElement protection = settings.Child("documentProtection");
            return protection.Attr("edit") == "trackedChanges"
                && protection.Attr("enforcement") is { } enforcement
                && LegacyDocValues.ParseOnOff(enforcement);
        }

        private static string ReadDocumentEndnotePosition(LegacyDocElement settings) {
            LegacyDocElement properties = settings.Child("endnotePr");
            return ReadEndnotePosition(properties.Child("pos").Val) ?? "docEnd";
        }

        private static void ThrowIfUnsupportedDocumentParts(LegacyDocSourcePackage source) {
            ThrowIfUnsupportedReviewMarkup(source);

            if (source.ReachesRelationshipType(type => type == LegacyDocSourcePackage.RelationshipTypePrefix + "chart")) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Charts are not supported yet.");
            }

            if (source.ReachesRelationshipType(type => type == LegacyDocSourcePackage.RelationshipTypePrefix + "diagramData"
                    || type == LegacyDocSourcePackage.RelationshipTypePrefix + "diagramLayout"
                    || type == LegacyDocSourcePackage.RelationshipTypePrefix + "diagramQuickStyle"
                    || type == LegacyDocSourcePackage.RelationshipTypePrefix + "diagramColors")) {
                throw new NotSupportedException("Native DOC saving currently supports text only. SmartArt diagrams are not supported yet.");
            }

            if (source.ReachesRelationshipType(type => type == LegacyDocSourcePackage.RelationshipTypePrefix + "oleObject"
                    || type == LegacyDocSourcePackage.RelationshipTypePrefix + "package")) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Embedded objects and packages are not supported yet.");
            }

            if (source.MainPart(LegacyDocSourcePackage.VbaProjectRelationshipType) != null) {
                throw new NotSupportedException("Native DOC saving currently does not support macro or VBA projects. Remove macros or save as DOCM/DOCX before saving as DOC.");
            }
        }

        private const string Word2012Namespace = "http://schemas.microsoft.com/office/word/2012/wordml";

        private static void ThrowIfUnsupportedReviewMarkup(LegacyDocSourcePackage source) {
            IReadOnlyList<LegacyDocElement> storyRoots = GetReviewMarkupStoryRoots(source);
            if (storyRoots.Any(HasMoveRevisionMarkup)) {
                throw new NotSupportedException("Native DOC saving currently supports tracked insertions and deletions, but not tracked move markup. Accept or reject tracked moves, or save as DOCX before saving as DOC.");
            }

            LegacyDocElement commentsEx = source.Root(source.MainPart(LegacyDocSourcePackage.CommentsExRelationshipType));
            if (commentsEx.Is(Word2012Namespace, "commentsEx")) {
                foreach (LegacyDocElement comment in commentsEx.Elements()) {
                    if (!comment.Is(Word2012Namespace, "commentEx")) continue;
                    if (!string.IsNullOrWhiteSpace(comment.AttrNs(Word2012Namespace, "paraIdParent"))
                        || comment.AttrNs(Word2012Namespace, "done") != null) {
                        throw new NotSupportedException(
                            "Native DOC saving currently does not support threaded replies or resolved-state comment metadata. "
                            + "Flatten the comment threads and remove resolved-state metadata, or save as DOCX before saving as DOC.");
                    }
                }
            }

            if (source.MainPart(LegacyDocSourcePackage.CommentsIdsRelationshipType) != null
                || source.MainPart(LegacyDocSourcePackage.PeopleRelationshipType) != null
                || source.MainRelatedParts().Any(part => IsUnsupportedModernCommentMetadataPart(source, part.PartName))) {
                throw new NotSupportedException(
                    "Native DOC saving currently does not support modern comment identity, people, or extensible metadata. "
                    + "Remove that metadata, or save as DOCX before saving as DOC.");
            }

            LegacyDocElement body = source.Body;
            if (storyRoots
                .Where(storyRoot => storyRoot != body)
                .Any(HasCommentMarkers)) {
                throw new NotSupportedException(
                    "Native DOC saving currently supports comment references in the main document body only. "
                    + "Remove comments from headers, footers, footnotes, and endnotes, or save as DOCX before saving as DOC.");
            }
        }

        private static bool IsUnsupportedModernCommentMetadataPart(LegacyDocSourcePackage source, string partName) {
            string contentType = source.GetContentType(partName) ?? string.Empty;
            return partName.IndexOf("commentsExtensible", StringComparison.OrdinalIgnoreCase) >= 0
                || contentType.IndexOf("commentsExtensible", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static IReadOnlyList<LegacyDocElement> GetReviewMarkupStoryRoots(LegacyDocSourcePackage source) {
            var roots = new List<LegacyDocElement>();
            if (source.Body.Exists) {
                roots.Add(source.Body);
            }

            foreach (LegacyDocRelatedPart headerPart in source.HeaderParts) {
                LegacyDocElement header = source.Root(headerPart.PartName);
                if (header.Exists) {
                    roots.Add(header);
                }
            }

            foreach (LegacyDocRelatedPart footerPart in source.FooterParts) {
                LegacyDocElement footer = source.Root(footerPart.PartName);
                if (footer.Exists) {
                    roots.Add(footer);
                }
            }

            if (source.Footnotes.Exists) {
                roots.Add(source.Footnotes);
            }

            if (source.Endnotes.Exists) {
                roots.Add(source.Endnotes);
            }

            return roots;
        }

        private static bool HasMoveRevisionMarkup(LegacyDocElement storyRoot) {
            return storyRoot.HasDescendant("moveFrom")
                || storyRoot.HasDescendant("moveTo");
        }

        private static bool HasCommentMarkers(LegacyDocElement storyRoot) {
            return storyRoot.HasDescendant("commentRangeStart")
                || storyRoot.HasDescendant("commentRangeEnd")
                || storyRoot.HasDescendant("commentReference");
        }

        private static bool IsUserEndnote(LegacyDocElement endnote) {
            return endnote.Attr("type") is not { } type || type == "normal";
        }

        private static void AddSection(List<LegacyDocWritableSection> sections, int endCharacter, LegacyDocSectionFormat format) {
            if (sections.Count > 0 && endCharacter < sections[sections.Count - 1].EndCharacter) {
                throw new NotSupportedException("Native DOC saving cannot write sections with non-monotonic text ranges.");
            }

            if (sections.Count > 0 && endCharacter == sections[sections.Count - 1].EndCharacter) {
                sections[sections.Count - 1] = new LegacyDocWritableSection(endCharacter, format);
                return;
            }

            sections.Add(new LegacyDocWritableSection(endCharacter, format));
        }

        private static void AppendBodyChild(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> paragraphFormats,
            LegacyDocWritableBookmarksBuilder bookmarks,
            BodyBlockBoundaryCache bodyBoundaries,
            LegacyDocElement child,
            LegacyDocSourcePackage source,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            LegacyDocTableStyleDefinitions tableStyleDefinitions,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            List<LegacyDocWritableSection> sections,
            ref LegacyDocSectionFormat finalSectionFormat,
            ref int bodyContentCount,
            string containerDescription) {
            if (child.Is("bookmarkStart") && text.Length > 0 && text[text.Length - 1] == '\a' &&
                bodyBoundaries.FindFollowing(child).Is("tbl")) {
                // A range starting at the next table excludes its required separator.
                // Ends closing the preceding table keep their original position.
                text.Append('\r');
            }
            switch (child.WName) {
                case "p":
                    AppendParagraph(text, runs, paragraphFormats, bookmarks, child, source, pictures, styleIndexes, footnotes, endnotes);
                    bodyContentCount++;

                    LegacyDocElement paragraphSectionProperties = child.Child("pPr").Child("sectPr");
                    if (paragraphSectionProperties.Exists) {
                        LegacyDocSectionFormat paragraphSectionFormat = ReadSupportedSectionProperties(paragraphSectionProperties);
                        // A section mark terminates this paragraph. The PLC and SEPX retain
                        // this section's own formatting, including its start type.
                        text[text.Length - 1] = LegacyDocSpecialCharacters.PageBreak;
                        AddSection(sections, text.Length, paragraphSectionFormat);
                    }

                    break;
                case "tbl":
                    AppendTable(text, runs, paragraphFormats, bookmarks, child, source, pictures, styleIndexes, tableStyleDefinitions, footnotes, endnotes);
                    bodyContentCount++;
                    break;
                case "sdt":
                    AppendBodyContentControl(
                        text,
                        runs,
                        paragraphFormats,
                        bookmarks,
                        bodyBoundaries,
                        child,
                        source,
                        pictures,
                        styleIndexes,
                        tableStyleDefinitions,
                        footnotes,
                        endnotes,
                        sections,
                        ref finalSectionFormat,
                        ref bodyContentCount);
                    break;
                case "bookmarkStart":
                    bookmarks.AddStart(child, text.Length);
                    break;
                case "bookmarkEnd":
                    int? precedingTableEnd = text.Length > 1 && text[text.Length - 1] == '\r' && text[text.Length - 2] == '\a' &&
                        bodyBoundaries.FindPreceding(child).Is("tbl") && bodyBoundaries.FindFollowing(child).Is("tbl")
                        ? text.Length - 1
                        : null;
                    bookmarks.AddEnd(child, text.Length, precedingTableEnd);
                    break;
                case "sectPr":
                    finalSectionFormat = ReadSupportedSectionProperties(child);
                    break;
                default:
                    throw new NotSupportedException($"Native DOC saving currently supports body paragraphs, simple body content controls, and simple tables with bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, imprint, hidden text, proofing exclusion, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family text runs. Unsupported {containerDescription} element: {child.LocalName}.");
            }
        }

        private static void AppendBodyContentControl(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> paragraphFormats,
            LegacyDocWritableBookmarksBuilder bookmarks,
            BodyBlockBoundaryCache bodyBoundaries,
            LegacyDocElement sdtBlock,
            LegacyDocSourcePackage source,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            LegacyDocTableStyleDefinitions tableStyleDefinitions,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            List<LegacyDocWritableSection> sections,
            ref LegacyDocSectionFormat finalSectionFormat,
            ref int bodyContentCount) {
            LegacyDocElement contentBlock = sdtBlock.Child("sdtContent");
            if (contentBlock.IsNone) {
                throw new NotSupportedException("Native DOC saving supports body content controls only when they contain simple body content.");
            }

            foreach (LegacyDocElement child in contentBlock.Elements()) {
                AppendBodyChild(
                    text,
                    runs,
                    paragraphFormats,
                    bookmarks,
                    bodyBoundaries,
                    child,
                    source,
                    pictures,
                    styleIndexes,
                    tableStyleDefinitions,
                    footnotes,
                    endnotes,
                    sections,
                    ref finalSectionFormat,
                    ref bodyContentCount,
                    "body content control");
            }
        }

        private static void AppendParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, LegacyDocWritableBookmarksBuilder bookmarks, LegacyDocElement paragraph, LegacyDocSourcePackage source, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            LegacyDocElement paragraphProperties = paragraph.Child("pPr");
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedBodyParagraphFormatting(paragraphProperties, styleIndexes);
            LegacyDocWritableFormatting paragraphMarkFormatting = ReadSupportedParagraphMarkRunFormatting(paragraphProperties);
            int paragraphStart = text.Length;

            LegacyDocElement[] children = paragraph.Elements().ToArray();
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationField(text, runs, child, LegacyDocWritableFormatting.Plain);
                    continue;
                }

                switch (child.WName) {
                    case "pPr":
                        break;
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendSupportedComplexPageNumberField(children, ref index, text, runs, bookmarks, LegacyDocWritableFormatting.Plain);
                        } else {
                            AppendSupportedRunText(text, runs, child, footnotes, endnotes, pictures, source.Main);
                        }

                        break;
                    case "ins":
                        AppendSupportedRevisionText(text, runs, child, LegacyDocRevisionKind.Inserted, footnotes, endnotes, LegacyDocWritableFormatting.Plain, pictures, source.Main);
                        break;
                    case "del":
                        AppendSupportedRevisionText(text, runs, child, LegacyDocRevisionKind.Deleted, footnotes, endnotes, LegacyDocWritableFormatting.Plain, pictures, source.Main);
                        break;
                    case "hyperlink":
                        AppendSupportedHyperlinkText(text, runs, bookmarks, child, source.Main, footnotes, endnotes);
                        break;
                    case "fldSimple":
                        AppendSupportedPageNumberFieldFromSimpleField(text, runs, bookmarks, child, LegacyDocWritableFormatting.Plain);
                        break;
                    case "sdt":
                        AppendSupportedInlineContentControlText(text, runs, bookmarks, child, source.Main, pictures, footnotes, endnotes, LegacyDocWritableFormatting.Plain, "body paragraph inline content control");
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, text.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving currently supports only text runs, {SupportedFieldNames} simple fields, bookmarks, inline content controls, and simple hyperlinks with bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, imprint, hidden text, proofing exclusion, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family formatting. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            text.Append('\r');
            AddParagraphMarkRunFormatting(runs, text.Length - 1, paragraphMarkFormatting);
            if (paragraphFormatting.HasFormatting) {
                paragraphFormats.Add(new LegacyDocWritableParagraph(paragraphStart, text.Length - paragraphStart, paragraphFormatting));
            }
        }

        private static bool IsIgnorableParagraphMarkup(LegacyDocElement element) {
            return element.Is("proofErr")
                || element.Is("commentRangeStart")
                || element.Is("commentRangeEnd");
        }


        private static bool CanWriteCompressedText(string text) {
            foreach (char character in text) {
                if (character > 0x7F) {
                    return false;
                }
            }

            return true;
        }

        private static byte[] EncodeCompressedText(string text) {
            byte[] bytes = new byte[text.Length];
            for (int i = 0; i < text.Length; i++) {
                bytes[i] = (byte)text[i];
            }

            return bytes;
        }

        private static int AlignToSector(int value) {
            return ((value + OleSectorSize - 1) / OleSectorSize) * OleSectorSize;
        }

        private static int AlignToEven(int value) {
            return value % 2 == 0 ? value : value + 1;
        }

        private static int GetSectionDataOffset(LegacyDocWritableBody body, int textByteLength, int chpxFkpOffset, int chpxPageCount, int papxFkpOffset, int papxPageCount) {
            return AlignToEven(body.HasParagraphFormatting
                ? papxFkpOffset + (papxPageCount * OleSectorSize)
                : body.HasCharacterFormatting
                    ? chpxFkpOffset + (chpxPageCount * OleSectorSize)
                    : TextOffset + textByteLength);
        }

        private static IReadOnlyList<LegacyDocWritableSectionRecord> CreateSectionRecords(LegacyDocWritableBody body, int firstSepxOffset) {
            var records = new List<LegacyDocWritableSectionRecord>(body.Sections.Count);
            int sepxOffset = firstSepxOffset;
            foreach (LegacyDocWritableSection section in body.Sections) {
                // Even default sections need a valid record; offset zero points into the FIB, not an empty SEP.
                byte[] sepx = CreateSepx(section.Format);
                int recordSepxOffset = sepxOffset;
                sepxOffset = AlignToEven(sepxOffset + sepx.Length);

                records.Add(new LegacyDocWritableSectionRecord(section.EndCharacter, recordSepxOffset, sepx));
            }

            return records;
        }

        private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xFF));
            stream.WriteByte((byte)(value >> 8));
        }

        private static void WriteInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private readonly struct LegacyDocWritableSection {
            internal LegacyDocWritableSection(int endCharacter, LegacyDocSectionFormat format) {
                EndCharacter = endCharacter;
                Format = format;
            }

            internal int EndCharacter { get; }

            internal LegacyDocSectionFormat Format { get; }
        }

        private readonly struct LegacyDocWritableSectionRecord {
            internal LegacyDocWritableSectionRecord(int endCharacter, int sepxOffset, byte[] sepx) {
                EndCharacter = endCharacter;
                SepxOffset = sepxOffset;
                Sepx = sepx;
            }

            internal int EndCharacter { get; }

            internal int SepxOffset { get; }

            internal byte[] Sepx { get; }

            internal int EndOffset => Sepx.Length == 0 ? SepxOffset : SepxOffset + Sepx.Length;
        }

    }
}
