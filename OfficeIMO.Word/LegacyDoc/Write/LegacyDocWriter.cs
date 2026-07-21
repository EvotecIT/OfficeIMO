using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing.Internal;
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
        private const ushort FibRgFcLcb97Size = 0x00B7;
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
        private const int FcPlcfBtePapxOffset = 0x102;
        private const int LcbPlcfBtePapxOffset = 0x106;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;
        private const int FcSttbfBkmkOffset = 0x142;
        private const int LcbSttbfBkmkOffset = 0x146;
        private const int FcSttbfRMarkOffset = 0x232;
        private const int LcbSttbfRMarkOffset = 0x236;
        private const int FcPlcfBkfOffset = 0x14A;
        private const int LcbPlcfBkfOffset = 0x14E;
        private const int FcPlcfBklOffset = 0x152;
        private const int LcbPlcfBklOffset = 0x156;
        private const int FcPlcfHddOffset = 0xF2;
        private const int LcbPlcfHddOffset = 0xF6;
        private const int DopBaseLength = 8;
        private const int DopBaseEndnotePlacementLength = 56;
        private const int DopBaseEndnotePlacementOffset = 52;
        private const int DopBaseEndnotePlacementShift = 16;
        private const ushort FacingPagesDopFlag = 0x0001;
        private const ushort NoteTextParagraphStyleIndex = 0x0023;
        private static readonly byte[] PlainParagraphPapx = { 0x00, 0x01, 0x00, 0x00 };
        private static readonly byte[] FootnoteTextParagraphPapx = { 0x00, 0x01, (byte)NoteTextParagraphStyleIndex, (byte)(NoteTextParagraphStyleIndex >> 8) };

        internal static byte[] WriteDocument(
            WordDocument document,
            WordSaveOptions? options = null,
            bool isTemplate = false) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            ThrowIfUnsupportedLegacyDocImportState(document, options);

            LegacyDocWritableBody body = BuildBody(document);
            byte[] wordDocumentStream = PadToRegularOleStream(CreateWordDocumentStream(body, isTemplate));
            byte[] tableStream = PadToRegularOleStream(CreateTableStream(body));
            IReadOnlyList<OfficeCompoundStream> propertyStreams = LegacyDocPropertySetWriter.CreateDocumentPropertyStreams(document);
            var streams = new List<OfficeCompoundStream>(propertyStreams.Count + 2) {
                new OfficeCompoundStream("WordDocument", wordDocumentStream),
                new OfficeCompoundStream("1Table", tableStream)
            };
            if (body.HasPictures) {
                streams.Add(new OfficeCompoundStream("Data", body.PictureData));
            }
            foreach (OfficeCompoundStream propertyStream in propertyStreams) {
                streams.Add(new OfficeCompoundStream(propertyStream.Name, PadToRegularOleStream(propertyStream.Bytes)));
            }

            OfficeCompoundFile? sourceCompoundFile = document.LegacyDocSourceCompoundFile;
            if (sourceCompoundFile == null) {
                return OfficeCompoundFileWriter.Write(streams);
            }

            return OfficeCompoundFileWriter.Rewrite(
                sourceCompoundFile,
                streams.ToDictionary(stream => stream.Name, stream => stream.Bytes, StringComparer.OrdinalIgnoreCase));
        }

        private static void ThrowIfUnsupportedLegacyDocImportState(WordDocument document, WordSaveOptions? options) {
            bool hasLegacyDigitalSignature = document.LegacyDocCompoundFeatures.Any(feature =>
                feature.Kind == LegacyDocCompoundFeatureKind.DigitalSignature);
            if (hasLegacyDigitalSignature
                && options?.SignedDocumentPolicy != WordSignedDocumentSavePolicy.AllowSignatureInvalidation) {
                throw new NotSupportedException(
                    "Native DOC saving is blocked because the imported legacy DOC contains digital-signature metadata. "
                    + "Rewriting the document invalidates that signature. Set WordSaveOptions.SignedDocumentPolicy to "
                    + "WordSignedDocumentSavePolicy.AllowSignatureInvalidation and explicitly allow conversion loss to continue.");
            }

            if (document.SourceFormat != WordFileFormat.Doc
                || (document.LegacyDocUnsupportedFeatures.Count == 0
                    && document.LegacyDocPreservedFeatures.Count == 0
                    && document.LegacyDocCompoundFeatures.Count == 0)
                || options?.LossPolicy == WordConversionLossPolicy.Allow) {
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

        private static LegacyDocWritableBody BuildBody(WordDocument document) {
            var wordDocument = document._wordprocessingDocument;
            if (wordDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as legacy DOC because WordDocument wasn't provided.");
            }

            wordDocument.Save();
            var mainPart = wordDocument.MainDocumentPart;
            var body = mainPart?.Document?.Body;
            if (body == null) {
                throw new InvalidOperationException("Document couldn't be saved as legacy DOC because the document body is missing.");
            }

            ThrowIfUnsupportedDocumentParts(document, mainPart);

            var text = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var paragraphFormats = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            var pictures = new LegacyDocWritablePictures(document);
            LegacyDocWritableFootnotes footnotes = ReadSupportedFootnotes(mainPart!, pictures);
            LegacyDocWritableEndnotes endnotes = ReadSupportedEndnotes(mainPart!, pictures);
            LegacyDocWritableStyleSheet styleSheet = CreateWritableStyleSheet(mainPart!, body);
            LegacyDocWritableComments comments = ReadSupportedComments(mainPart!, pictures, styleSheet.StyleIndexes);
            IReadOnlyDictionary<string, Style> tableStyleDefinitions = ReadTableStyleDefinitions(mainPart!);
            LegacyDocSectionFormat finalSectionFormat = LegacyDocSectionFormat.Default;
            var sections = new List<LegacyDocWritableSection>();
            SectionMarkValues? pendingSectionBreakType = null;
            int bodyContentCount = 0;
            foreach (OpenXmlElement child in body.ChildElements) {
                AppendBodyChild(
                    text,
                    runs,
                    paragraphFormats,
                    bookmarks,
                    child,
                    mainPart!,
                    pictures,
                    styleSheet.StyleIndexes,
                    tableStyleDefinitions,
                    footnotes,
                    endnotes,
                    sections,
                    ref finalSectionFormat,
                    ref pendingSectionBreakType,
                    ref bodyContentCount,
                    "body");
            }

            if (bodyContentCount == 0) {
                text.Append('\r');
            }

            AddSection(sections, text.Length, finalSectionFormat.WithSectionBreakType(pendingSectionBreakType));
            footnotes.ThrowIfUnreferencedFootnotesRemain();
            endnotes.ThrowIfUnreferencedEndnotesRemain();
            comments.BindBodyReferences(body, text.ToString());
            bool hasNoteReferences = footnotes.HasReferences || endnotes.HasReferences;
            LegacyDocWritableHeaderFooterStories headerFooterStories = BuildHeaderFooterStories(document, mainPart!, pictures, hasNoteReferences, styleSheet.StyleIndexes);
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
            Settings? settings = mainPart!.DocumentSettingsPart?.Settings;
            bool trackRevisions = settings?.Elements<TrackRevisions>().Any(IsOnOffEnabled) == true;
            bool lockRevisionTracking = IsLockedRevisionTracking(settings);
            return new LegacyDocWritableBody(
                text.ToString(),
                runs,
                paragraphFormats,
                bookmarks.Create(),
                sections,
                styleSheet,
                footnoteStories,
                endnoteStories,
                headerFooterStories,
                commentStories,
                pictures.DataBytes,
                pictures.HasPictures,
                HasEvenAndOddHeaders(mainPart),
                ReadDocumentEndnotePosition(sections),
                trackRevisions || lockRevisionTracking,
                lockRevisionTracking);
        }

        private static bool HasEvenAndOddHeaders(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart) {
            Settings? settings = mainPart.DocumentSettingsPart?.Settings;
            return settings?.Elements<EvenAndOddHeaders>().Any(IsOnOffEnabled) == true;
        }

        private static bool IsLockedRevisionTracking(Settings? settings) {
            DocumentProtection? protection = settings?.Elements<DocumentProtection>().FirstOrDefault();
            return protection?.Edit?.Value == DocumentProtectionValues.TrackedChanges
                && protection.Enforcement != null
                && protection.Enforcement.Value;
        }

        private static EndnotePositionValues? ReadDocumentEndnotePosition(IReadOnlyList<LegacyDocWritableSection> sections) {
            EndnotePositionValues? position = null;
            foreach (LegacyDocWritableSection section in sections) {
                EndnotePositionValues? sectionPosition = section.Format.EndnotePosition;
                if (sectionPosition == null) {
                    continue;
                }

                if (position != null && position.Value != sectionPosition.Value) {
                    throw new NotSupportedException("Native DOC saving supports only one endnote placement for the whole document.");
                }

                position = sectionPosition;
            }

            return position;
        }

        private static void ThrowIfUnsupportedDocumentParts(WordDocument document, DocumentFormat.OpenXml.Packaging.MainDocumentPart? mainPart) {
            if (mainPart == null) {
                return;
            }

            ThrowIfUnsupportedReviewMarkup(mainPart);

            if (HasRelatedPart<ChartPart>(mainPart)) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Charts are not supported yet.");
            }

            if (HasRelatedPart<DiagramDataPart>(mainPart)
                || HasRelatedPart<DiagramLayoutDefinitionPart>(mainPart)
                || HasRelatedPart<DiagramStylePart>(mainPart)
                || HasRelatedPart<DiagramColorsPart>(mainPart)) {
                throw new NotSupportedException("Native DOC saving currently supports text only. SmartArt diagrams are not supported yet.");
            }

            if (HasRelatedPart<EmbeddedObjectPart>(mainPart)
                || HasRelatedPart<EmbeddedPackagePart>(mainPart)) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Embedded objects and packages are not supported yet.");
            }

            if (mainPart.VbaProjectPart != null) {
                throw new NotSupportedException("Native DOC saving currently does not support macro or VBA projects. Remove macros or save as DOCM/DOCX before saving as DOC.");
            }
        }

        private static bool HasRelatedPart<TPart>(OpenXmlPartContainer container)
            where TPart : OpenXmlPart {
            return HasRelatedPart<TPart>(container, new HashSet<OpenXmlPart>());
        }

        private static bool HasRelatedPart<TPart>(OpenXmlPartContainer container, HashSet<OpenXmlPart> visited)
            where TPart : OpenXmlPart {
            foreach (IdPartPair relationship in container.Parts) {
                OpenXmlPart part = relationship.OpenXmlPart;
                if (part is TPart) {
                    return true;
                }

                if (visited.Add(part) && HasRelatedPart<TPart>(part, visited)) {
                    return true;
                }
            }

            return false;
        }

        private static void ThrowIfUnsupportedReviewMarkup(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart) {
            IReadOnlyList<OpenXmlElement> storyRoots = GetReviewMarkupStoryRoots(mainPart);
            if (storyRoots.Any(HasMoveRevisionMarkup)) {
                throw new NotSupportedException("Native DOC saving currently supports tracked insertions and deletions, but not tracked move markup. Accept or reject tracked moves, or save as DOCX before saving as DOC.");
            }

            CommentsEx? commentsEx = mainPart.WordprocessingCommentsExPart?.CommentsEx;
            if (commentsEx?.Elements<CommentEx>().Any(comment =>
                    !string.IsNullOrWhiteSpace(comment.ParaIdParent?.Value)
                    || comment.Done != null) == true) {
                throw new NotSupportedException(
                    "Native DOC saving currently does not support threaded replies or resolved-state comment metadata. "
                    + "Flatten the comment threads and remove resolved-state metadata, or save as DOCX before saving as DOC.");
            }

            if (mainPart.WordprocessingCommentsIdsPart != null
                || mainPart.WordprocessingPeoplePart != null
                || mainPart.Parts.Select(pair => pair.OpenXmlPart).Any(IsUnsupportedModernCommentMetadataPart)) {
                throw new NotSupportedException(
                    "Native DOC saving currently does not support modern comment identity, people, or extensible metadata. "
                    + "Remove that metadata, or save as DOCX before saving as DOC.");
            }

            if (storyRoots
                .Where(storyRoot => !ReferenceEquals(storyRoot, mainPart.Document?.Body))
                .Any(HasCommentMarkers)) {
                throw new NotSupportedException(
                    "Native DOC saving currently supports comment references in the main document body only. "
                    + "Remove comments from headers, footers, footnotes, and endnotes, or save as DOCX before saving as DOC.");
            }
        }

        private static bool IsUnsupportedModernCommentMetadataPart(OpenXmlPart part) {
            string uri = part.Uri.OriginalString;
            string contentType = part.ContentType;
            return uri.IndexOf("commentsExtensible", StringComparison.OrdinalIgnoreCase) >= 0
                || contentType.IndexOf("commentsExtensible", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static IReadOnlyList<OpenXmlElement> GetReviewMarkupStoryRoots(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart) {
            var roots = new List<OpenXmlElement>();
            if (mainPart.Document?.Body != null) {
                roots.Add(mainPart.Document.Body);
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                if (headerPart.Header != null) {
                    roots.Add(headerPart.Header);
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts) {
                if (footerPart.Footer != null) {
                    roots.Add(footerPart.Footer);
                }
            }

            if (mainPart.FootnotesPart?.Footnotes != null) {
                roots.Add(mainPart.FootnotesPart.Footnotes);
            }

            if (mainPart.EndnotesPart?.Endnotes != null) {
                roots.Add(mainPart.EndnotesPart.Endnotes);
            }

            return roots;
        }

        private static bool HasMoveRevisionMarkup(OpenXmlElement storyRoot) {
            return storyRoot.Descendants<MoveFromRun>().Any()
                || storyRoot.Descendants<MoveToRun>().Any();
        }

        private static bool HasCommentMarkers(OpenXmlElement storyRoot) {
            return storyRoot.Descendants<CommentRangeStart>().Any()
                || storyRoot.Descendants<CommentRangeEnd>().Any()
                || storyRoot.Descendants<CommentReference>().Any();
        }

        private static bool IsUserEndnote(Endnote endnote) {
            return endnote.Type == null || endnote.Type.Value == FootnoteEndnoteValues.Normal;
        }

        private static bool IsPureSectionBreakParagraph(Paragraph paragraph) {
            ParagraphProperties? paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
            if (paragraphProperties?.GetFirstChild<SectionProperties>() == null) {
                return false;
            }

            return paragraph.ChildElements.All(element => element is ParagraphProperties)
                && paragraphProperties.ChildElements.All(element => element is SectionProperties);
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
            OpenXmlElement child,
            MainDocumentPart mainPart,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            IReadOnlyDictionary<string, Style> tableStyleDefinitions,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            List<LegacyDocWritableSection> sections,
            ref LegacyDocSectionFormat finalSectionFormat,
            ref SectionMarkValues? pendingSectionBreakType,
            ref int bodyContentCount,
            string containerDescription) {
            switch (child) {
                case Paragraph paragraph:
                    if (!IsPureSectionBreakParagraph(paragraph)) {
                        AppendParagraph(text, runs, paragraphFormats, bookmarks, paragraph, mainPart, pictures, styleIndexes, footnotes, endnotes);
                        bodyContentCount++;
                    }

                    SectionProperties? paragraphSectionProperties = paragraph.GetFirstChild<ParagraphProperties>()?.GetFirstChild<SectionProperties>();
                    if (paragraphSectionProperties != null) {
                        LegacyDocSectionFormat paragraphSectionFormat = ReadSupportedSectionProperties(paragraphSectionProperties);
                        AddSection(sections, text.Length, paragraphSectionFormat.WithSectionBreakType(pendingSectionBreakType));
                        pendingSectionBreakType = paragraphSectionFormat.SectionBreakType;
                    }

                    break;
                case Table table:
                    AppendTable(text, runs, paragraphFormats, bookmarks, table, mainPart, pictures, styleIndexes, tableStyleDefinitions, footnotes, endnotes);
                    bodyContentCount++;
                    break;
                case SdtBlock sdtBlock:
                    AppendBodyContentControl(
                        text,
                        runs,
                        paragraphFormats,
                        bookmarks,
                        sdtBlock,
                        mainPart,
                        pictures,
                        styleIndexes,
                        tableStyleDefinitions,
                        footnotes,
                        endnotes,
                        sections,
                        ref finalSectionFormat,
                        ref pendingSectionBreakType,
                        ref bodyContentCount);
                    break;
                case BookmarkStart bookmarkStart:
                    bookmarks.AddStart(bookmarkStart, text.Length);
                    break;
                case BookmarkEnd bookmarkEnd:
                    bookmarks.AddEnd(bookmarkEnd, text.Length);
                    break;
                case SectionProperties sectionProperties:
                    finalSectionFormat = ReadSupportedSectionProperties(sectionProperties);
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
            SdtBlock sdtBlock,
            MainDocumentPart mainPart,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            IReadOnlyDictionary<string, Style> tableStyleDefinitions,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            List<LegacyDocWritableSection> sections,
            ref LegacyDocSectionFormat finalSectionFormat,
            ref SectionMarkValues? pendingSectionBreakType,
            ref int bodyContentCount) {
            SdtContentBlock? contentBlock = sdtBlock.SdtContentBlock;
            if (contentBlock == null) {
                throw new NotSupportedException("Native DOC saving supports body content controls only when they contain simple body content.");
            }

            foreach (OpenXmlElement child in contentBlock.ChildElements) {
                AppendBodyChild(
                    text,
                    runs,
                    paragraphFormats,
                    bookmarks,
                    child,
                    mainPart,
                    pictures,
                    styleIndexes,
                    tableStyleDefinitions,
                    footnotes,
                    endnotes,
                    sections,
                    ref finalSectionFormat,
                    ref pendingSectionBreakType,
                    ref bodyContentCount,
                    "body content control");
            }
        }

        private static void AppendParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, LegacyDocWritableBookmarksBuilder bookmarks, Paragraph paragraph, MainDocumentPart mainPart, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            ParagraphProperties? paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedBodyParagraphFormatting(paragraphProperties, styleIndexes);
            LegacyDocWritableFormatting paragraphMarkFormatting = ReadSupportedParagraphMarkRunFormatting(paragraphProperties);
            int paragraphStart = text.Length;

            OpenXmlElement[] children = paragraph.ChildElements.ToArray();
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendSupportedComplexPageNumberField(children, ref index, text, runs, bookmarks, LegacyDocWritableFormatting.Plain);
                        } else {
                            AppendSupportedRunText(text, runs, run, footnotes, endnotes, pictures, mainPart);
                        }

                        break;
                    case InsertedRun insertedRun:
                        AppendSupportedRevisionText(text, runs, insertedRun, LegacyDocRevisionKind.Inserted, footnotes, endnotes, LegacyDocWritableFormatting.Plain, pictures, mainPart);
                        break;
                    case DeletedRun deletedRun:
                        AppendSupportedRevisionText(text, runs, deletedRun, LegacyDocRevisionKind.Deleted, footnotes, endnotes, LegacyDocWritableFormatting.Plain, pictures, mainPart);
                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedHyperlinkText(text, runs, bookmarks, hyperlink, mainPart, footnotes, endnotes);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedPageNumberFieldFromSimpleField(text, runs, bookmarks, simpleField, LegacyDocWritableFormatting.Plain);
                        break;
                    case DocumentFormat.OpenXml.Math.OfficeMath officeMath:
                        AppendMathEquationField(text, runs, officeMath, LegacyDocWritableFormatting.Plain);
                        break;
                    case DocumentFormat.OpenXml.Math.Paragraph mathParagraph:
                        AppendMathEquationField(text, runs, mathParagraph, LegacyDocWritableFormatting.Plain);
                        break;
                    case SdtRun sdtRun:
                        AppendSupportedInlineContentControlText(text, runs, bookmarks, sdtRun, mainPart, pictures, footnotes, endnotes, LegacyDocWritableFormatting.Plain, "body paragraph inline content control");
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, text.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, text.Length);
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

        private static bool IsIgnorableParagraphMarkup(OpenXmlElement element) {
            return element is ProofError
                or CommentRangeStart
                or CommentRangeEnd;
        }

        private static byte[] CreateWordDocumentStream(LegacyDocWritableBody body, bool isTemplate) {
            bool compressedText = CanWriteCompressedText(body.StoredText);
            int bytesPerCharacter = compressedText ? 1 : 2;
            byte[] textBytes = compressedText ? EncodeCompressedText(body.StoredText) : Encoding.Unicode.GetBytes(body.StoredText);
            byte[] fontTable = CreateFontTable(body.FontFamilies);
            IReadOnlyList<IReadOnlyList<LegacyDocWritableSegment>> chpxPages = body.ChpxPages;
            int chpxFkpOffset = body.HasCharacterFormatting
                ? AlignToSector(TextOffset + textBytes.Length)
                : 0;
            IReadOnlyList<IReadOnlyList<LegacyDocWritableParagraphSegment>> papxPages = body.PapxPages;
            int papxFkpOffset = body.HasParagraphFormatting
                ? AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + (chpxPages.Count * OleSectorSize) : TextOffset + textBytes.Length)
                : 0;
            int sectionDataOffset = GetSectionDataOffset(body, textBytes.Length, chpxFkpOffset, chpxPages.Count, papxFkpOffset, papxPages.Count);
            IReadOnlyList<LegacyDocWritableSectionRecord> sectionRecords = CreateSectionRecords(body, sectionDataOffset);
            int streamLength = body.HasParagraphFormatting
                ? papxFkpOffset + (papxPages.Count * OleSectorSize)
                : body.HasCharacterFormatting
                    ? chpxFkpOffset + (chpxPages.Count * OleSectorSize)
                    : TextOffset + textBytes.Length;
            if (body.HasSectionDescriptors) {
                streamLength = Math.Max(streamLength, sectionRecords.Count == 0 ? sectionDataOffset : sectionRecords.Max(record => record.EndOffset));
            }
            var stream = new byte[Math.Max(FibLength, streamLength)];
            WriteUInt16(stream, 0x00, WordDocumentMagic);
            WriteUInt16(stream, 0x02, Word97FibVersion);
            WriteUInt16(stream, 0x06, DefaultLanguageId);
            ushort fibFlags = DefaultFibFlags;
            if (body.HasPictures) fibFlags = unchecked((ushort)(fibFlags | HasPicturesFibFlag));
            if (isTemplate) fibFlags = unchecked((ushort)(fibFlags | TemplateFibFlag));
            WriteUInt16(stream, 0x0A, fibFlags);
            WriteUInt16(stream, 0x0C, Word97FibBackVersion);
            WriteInt32(stream, 0x18, TextOffset);
            WriteInt32(stream, 0x1C, TextOffset + textBytes.Length);
            WriteUInt16(stream, 0x20, FibRgW97WordCount);
            WriteUInt16(stream, 0x3C, DefaultLanguageId);
            WriteUInt16(stream, 0x3E, FibRgLw97DwordCount);
            WriteInt32(stream, 0x4C, body.Text.Length);
            WriteInt32(stream, 0x50, body.FootnoteText.Length);
            WriteInt32(stream, 0x54, body.HeaderFooterText.Length);
            WriteInt32(stream, 0x5C, body.CommentText.Length);
            WriteInt32(stream, 0x60, body.EndnoteText.Length);
            WriteUInt16(stream, 0x98, FibRgFcLcb97Size);
            WriteInt32(stream, FcStshfOffset, body.HasStyleSheet ? body.StyleSheetOffsetInTableStream : 0);
            WriteInt32(stream, LcbStshfOffset, body.StyleSheet.Bytes.Length);
            WriteInt32(stream, 0xFA, body.HasCharacterFormatting ? ClxLength : 0);
            WriteInt32(stream, 0xFE, body.HasCharacterFormatting ? body.ChpxPlcLength : 0);
            WriteInt32(stream, FcPlcfBtePapxOffset, body.HasParagraphFormatting ? body.PapxPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBtePapxOffset, body.HasParagraphFormatting ? body.PapxPlcLength : 0);
            WriteInt32(stream, FcPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcLength : 0);
            WriteInt32(stream, FcPlcffndRefOffset, body.HasFootnotes ? body.PlcffndRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcffndRefOffset, body.HasFootnotes ? body.PlcffndRef.Length : 0);
            WriteInt32(stream, FcPlcffndTxtOffset, body.HasFootnotes ? body.PlcffndTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcffndTxtOffset, body.HasFootnotes ? body.PlcffndTxt.Length : 0);
            WriteInt32(stream, FcPlcfandRefOffset, body.HasComments ? body.PlcfandRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfandRefOffset, body.HasComments ? body.PlcfandRef.Length : 0);
            WriteInt32(stream, FcPlcfandTxtOffset, body.HasComments ? body.PlcfandTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfandTxtOffset, body.HasComments ? body.PlcfandTxt.Length : 0);
            WriteInt32(stream, FcPlcfendRefOffset, body.HasEndnotes ? body.PlcfendRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfendRefOffset, body.HasEndnotes ? body.PlcfendRef.Length : 0);
            WriteInt32(stream, FcPlcfendTxtOffset, body.HasEndnotes ? body.PlcfendTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfendTxtOffset, body.HasEndnotes ? body.PlcfendTxt.Length : 0);
            WriteInt32(stream, FcPlcfHddOffset, body.HasHeaderFooterStories ? body.PlcfHddOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfHddOffset, body.HasHeaderFooterStories ? body.PlcfHdd.Length : 0);
            WriteInt32(stream, FcSttbfBkmkOffset, body.HasBookmarks ? body.SttbfBkmkOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfBkmkOffset, body.HasBookmarks ? body.SttbfBkmk.Length : 0);
            WriteInt32(stream, FcSttbfRMarkOffset, body.HasRevisions ? body.SttbfRMarkOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfRMarkOffset, body.SttbfRMark.Length);
            WriteInt32(stream, FcPlcfBkfOffset, body.HasBookmarks ? body.PlcfBkfOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBkfOffset, body.HasBookmarks ? body.PlcfBkf.Length : 0);
            WriteInt32(stream, FcPlcfBklOffset, body.HasBookmarks ? body.PlcfBklOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBklOffset, body.HasBookmarks ? body.PlcfBkl.Length : 0);
            WriteInt32(stream, FcSttbfFfnOffset, body.HasFontTable ? body.FontTableOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfFfnOffset, fontTable.Length);
            WriteInt32(stream, FcDopOffset, body.HasDocumentOptions ? body.DopOffsetInTableStream : 0);
            WriteInt32(stream, LcbDopOffset, body.HasDocumentOptions ? body.DopLength : 0);
            WriteInt32(stream, 0x1A2, 0);
            WriteInt32(stream, 0x1A6, ClxLength);
            Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);
            if (body.HasCharacterFormatting) {
                for (int pageIndex = 0; pageIndex < chpxPages.Count; pageIndex++) {
                    WriteChpxFkp(
                        stream,
                        chpxFkpOffset + (pageIndex * OleSectorSize),
                        chpxPages[pageIndex],
                        body.FontFamilyIndexes,
                        body.RevisionAuthorIndexes,
                        bytesPerCharacter);
                }
            }

            if (body.HasParagraphFormatting) {
                for (int pageIndex = 0; pageIndex < papxPages.Count; pageIndex++) {
                    LegacyDocParagraphFormattingWriter.WritePapxFkp(
                        stream,
                        papxFkpOffset + (pageIndex * OleSectorSize),
                        TextOffset,
                        OleSectorSize,
                        papxPages[pageIndex],
                        bytesPerCharacter);
                }
            }

            if (body.HasSectionDescriptors) {
                foreach (LegacyDocWritableSectionRecord record in sectionRecords) {
                    if (record.Sepx.Length == 0) {
                        continue;
                    }

                    Buffer.BlockCopy(record.Sepx, 0, stream, record.SepxOffset, record.Sepx.Length);
                }
            }

            return stream;
        }

        private static byte[] CreateTableStream(LegacyDocWritableBody body) {
            byte[] fontTable = CreateFontTable(body.FontFamilies);
            bool compressedText = CanWriteCompressedText(body.StoredText);
            int bytesPerCharacter = compressedText ? 1 : 2;
            int textByteLength = checked(body.StoredText.Length * bytesPerCharacter);
            var table = new byte[body.FontTableOffsetInTableStream + fontTable.Length];
            table[0] = 0x02;
            WriteInt32(table, 1, 16);
            WriteInt32(table, 5, 0);
            WriteInt32(table, 9, body.PieceTableCharacterCount);
            WriteUInt16(table, 13, body.HasFootnotes ? FootnotePcdFlags : DefaultPcdFlags);
            WriteUInt32(table, 15, compressedText ? CompressedTextFlag | (uint)(TextOffset * 2) : TextOffset);
            WriteUInt16(table, 19, 0);

            if (body.HasCharacterFormatting) {
                int chpxFkpOffset = AlignToSector(TextOffset + textByteLength);
                WriteChpxBtePlc(table, body, chpxFkpOffset, bytesPerCharacter);
            }

            if (body.HasParagraphFormatting) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + textByteLength)
                    : 0;
                int papxFkpOffset = AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + (body.ChpxPageCount * OleSectorSize) : TextOffset + textByteLength);
                WritePapxBtePlc(table, body, papxFkpOffset, bytesPerCharacter);
            }

            if (body.HasSectionDescriptors) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + textByteLength)
                    : 0;
                int papxFkpOffset = body.HasParagraphFormatting
                    ? AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + (body.ChpxPageCount * OleSectorSize) : TextOffset + textByteLength)
                    : 0;
                int sepxOffset = AlignToEven(body.HasParagraphFormatting
                    ? papxFkpOffset + (body.PapxPageCount * OleSectorSize)
                    : body.HasCharacterFormatting
                        ? chpxFkpOffset + (body.ChpxPageCount * OleSectorSize)
                        : TextOffset + textByteLength);
                WritePlcfSed(table, body.SedPlcOffsetInTableStream, CreateSectionRecords(body, sepxOffset));
            }

            if (body.HasFootnotes) {
                Buffer.BlockCopy(body.PlcffndRef, 0, table, body.PlcffndRefOffsetInTableStream, body.PlcffndRef.Length);
                Buffer.BlockCopy(body.PlcffndTxt, 0, table, body.PlcffndTxtOffsetInTableStream, body.PlcffndTxt.Length);
            }

            if (body.HasHeaderFooterStories) {
                Buffer.BlockCopy(body.PlcfHdd, 0, table, body.PlcfHddOffsetInTableStream, body.PlcfHdd.Length);
            }

            if (body.HasComments) {
                Buffer.BlockCopy(body.PlcfandRef, 0, table, body.PlcfandRefOffsetInTableStream, body.PlcfandRef.Length);
                Buffer.BlockCopy(body.PlcfandTxt, 0, table, body.PlcfandTxtOffsetInTableStream, body.PlcfandTxt.Length);
            }

            if (body.HasEndnotes) {
                Buffer.BlockCopy(body.PlcfendRef, 0, table, body.PlcfendRefOffsetInTableStream, body.PlcfendRef.Length);
                Buffer.BlockCopy(body.PlcfendTxt, 0, table, body.PlcfendTxtOffsetInTableStream, body.PlcfendTxt.Length);
            }

            if (body.HasDocumentOptions) {
                byte[] dop = CreateDopBase(body);
                Buffer.BlockCopy(dop, 0, table, body.DopOffsetInTableStream, dop.Length);
            }

            if (body.HasBookmarks) {
                Buffer.BlockCopy(body.SttbfBkmk, 0, table, body.SttbfBkmkOffsetInTableStream, body.SttbfBkmk.Length);
                Buffer.BlockCopy(body.PlcfBkf, 0, table, body.PlcfBkfOffsetInTableStream, body.PlcfBkf.Length);
                Buffer.BlockCopy(body.PlcfBkl, 0, table, body.PlcfBklOffsetInTableStream, body.PlcfBkl.Length);
            }

            if (body.HasRevisions) {
                Buffer.BlockCopy(body.SttbfRMark, 0, table, body.SttbfRMarkOffsetInTableStream, body.SttbfRMark.Length);
            }

            if (body.HasStyleSheet) {
                Buffer.BlockCopy(body.StyleSheet.Bytes, 0, table, body.StyleSheetOffsetInTableStream, body.StyleSheet.Bytes.Length);
            }

            if (fontTable.Length > 0) {
                Buffer.BlockCopy(fontTable, 0, table, body.FontTableOffsetInTableStream, fontTable.Length);
            }

            return table;
        }

        private static byte[] CreateDopBase(LegacyDocWritableBody body) {
            var dop = new byte[body.DopLength];
            if (body.FacingPages) {
                WriteUInt16(dop, 0, FacingPagesDopFlag);
            }

            uint revisionFlags = 0;
            if (body.TrackRevisions) {
                revisionFlags |= 0x00008000;
            }
            if (body.LockRevisionTracking) {
                revisionFlags |= 0x40000000;
            }
            if (revisionFlags != 0) {
                WriteUInt32(dop, 4, revisionFlags);
            }

            if (body.EndnotePosition != null) {
                uint placement = (uint)GetEndnotePositionOperand(body.EndnotePosition.Value)!.Value;
                WriteUInt32(dop, DopBaseEndnotePlacementOffset, placement << DopBaseEndnotePlacementShift);
            }

            return dop;
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
                byte[] sepx = section.Format.HasFormatting ? CreateSepx(section.Format) : Array.Empty<byte>();
                int recordSepxOffset = 0;
                if (sepx.Length > 0) {
                    recordSepxOffset = sepxOffset;
                    sepxOffset = AlignToEven(sepxOffset + sepx.Length);
                }

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
