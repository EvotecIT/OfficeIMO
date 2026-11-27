using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int HeaderFooterSeparatorStoryCount = 6;
        private const int HeaderFooterStoriesPerSection = 6;
        private const string HeaderFooterAutoNumberSeparatorStory = "\u0003\r\r";
        private const string HeaderFooterContinuationSeparatorStory = "\u0004\r\r";
        private const string HeaderFooterSeparatorTerminator = "\r";

        private static LegacyDocWritableHeaderFooterStories BuildHeaderFooterStories(LegacyDocSourcePackage source, LegacyDocWritablePictures pictures, bool includeDefaultSeparators, IReadOnlyDictionary<string, ushort> styleIndexes) {
            if (!includeDefaultSeparators && source.HeaderParts.Count == 0 && source.FooterParts.Count == 0) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            IReadOnlyList<LegacyDocElement> sections = LegacyDocWordHelpers.GetDocumentSectionProperties(source);
            if (!includeDefaultSeparators && sections.Count == 0) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            int sectionCount = Math.Max(sections.Count, 1);
            LegacyDocWritableHeaderFooterStory[] stories = new LegacyDocWritableHeaderFooterStory[HeaderFooterSeparatorStoryCount + (sectionCount * HeaderFooterStoriesPerSection)];
            for (int storyIndex = 0; storyIndex < stories.Length; storyIndex++) {
                stories[storyIndex] = LegacyDocWritableHeaderFooterStory.Empty;
            }

            if (includeDefaultSeparators) {
                stories[0] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterAutoNumberSeparatorStory);
                stories[1] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterContinuationSeparatorStory);
                stories[3] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterAutoNumberSeparatorStory);
                stories[4] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterContinuationSeparatorStory);
            }

            LegacyDocWritableHeaderFooterStory? currentEvenHeader = null;
            LegacyDocWritableHeaderFooterStory? currentDefaultHeader = null;
            LegacyDocWritableHeaderFooterStory? currentFirstHeader = null;
            LegacyDocWritableHeaderFooterStory? currentEvenFooter = null;
            LegacyDocWritableHeaderFooterStory? currentDefaultFooter = null;
            LegacyDocWritableHeaderFooterStory? currentFirstFooter = null;

            for (int sectionIndex = 0; sectionIndex < sections.Count; sectionIndex++) {
                LegacyDocElement sectionProperties = sections[sectionIndex];
                int sectionStoryOffset = HeaderFooterSeparatorStoryCount + (sectionIndex * HeaderFooterStoriesPerSection);
                currentEvenHeader = ReadHeaderStory(source, sectionProperties, "even", pictures, styleIndexes) ?? currentEvenHeader;
                currentDefaultHeader = ReadHeaderStory(source, sectionProperties, "default", pictures, styleIndexes) ?? currentDefaultHeader;
                currentEvenFooter = ReadFooterStory(source, sectionProperties, "even", pictures, styleIndexes) ?? currentEvenFooter;
                currentDefaultFooter = ReadFooterStory(source, sectionProperties, "default", pictures, styleIndexes) ?? currentDefaultFooter;
                currentFirstHeader = ReadHeaderStory(source, sectionProperties, "first", pictures, styleIndexes) ?? currentFirstHeader;
                currentFirstFooter = ReadFooterStory(source, sectionProperties, "first", pictures, styleIndexes) ?? currentFirstFooter;

                stories[sectionStoryOffset] = currentEvenHeader ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 1] = currentDefaultHeader ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 2] = currentEvenFooter ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 3] = currentDefaultFooter ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 4] = currentFirstHeader ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 5] = currentFirstFooter ?? LegacyDocWritableHeaderFooterStory.Empty;
            }

            ThrowIfUnreferencedHeaderFooterContent(source, sections);
            if (!includeDefaultSeparators && stories.All(story => story.Text.Length == 0)) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            var text = new StringBuilder();
            var characterPositions = new List<int>(stories.Length + 2);
            var markerPositions = new List<int>();
            var formattedRuns = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            foreach (LegacyDocWritableHeaderFooterStory story in stories) {
                characterPositions.Add(text.Length);
                AddHeaderFooterSpecialCharacterPositions(text.Length, story.Text, markerPositions);
                foreach (LegacyDocWritableRun run in story.FormattedRuns) {
                    formattedRuns.Add(new LegacyDocWritableRun(text.Length + run.StartCharacter, run.Length, run.Formatting, run.PictureDataOffset, run.FieldLocked));
                }

                foreach (LegacyDocWritableParagraph paragraph in story.FormattedParagraphs) {
                    formattedParagraphs.Add(new LegacyDocWritableParagraph(text.Length + paragraph.StartCharacter, paragraph.Length, paragraph.Formatting));
                }

                bookmarks.AddRange(story.Bookmarks, text.Length);
                text.Append(story.Text);
            }

            text.Append(HeaderFooterSeparatorTerminator);
            characterPositions.Add(text.Length - HeaderFooterSeparatorTerminator.Length);
            characterPositions.Add(text.Length + 1);
            byte[] plcfHdd = new byte[characterPositions.Count * 4];
            for (int index = 0; index < characterPositions.Count; index++) {
                WriteInt32(plcfHdd, index * 4, characterPositions[index]);
            }

            return new LegacyDocWritableHeaderFooterStories(text.ToString(), plcfHdd, markerPositions, formattedRuns, formattedParagraphs, bookmarks.Create());
        }

        private static void AddHeaderFooterSpecialCharacterPositions(int storyStart, string story, List<int> markerPositions) {
            for (int index = 0; index < story.Length; index++) {
                char character = story[index];
                if (character == '\u0003' || character == '\u0004') {
                    markerPositions.Add(storyStart + index);
                }
            }
        }

        private static LegacyDocWritableHeaderFooterStory? ReadHeaderStory(LegacyDocSourcePackage source, LegacyDocElement sectionProperties, string type, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes) {
            LegacyDocElement[] references = sectionProperties.Elements("headerReference").ToArray();
            if (references.Length == 0) {
                return null;
            }

            string kind = GetHeaderFooterStoryDescription(type, "header");
            LegacyDocElement reference = GetSingleReference(references, type, "header");
            if (reference.IsNone) {
                return null;
            }

            LegacyDocPart headerPart = GetReferencedHeaderOrFooterPart(source, reference.Attr(OfficeXmlNamespaces.Relationships, "id"), kind, header: true);
            return ReadSimpleHeaderFooterStory(headerPart.Root, headerPart, pictures, kind, styleIndexes, source);
        }

        private static LegacyDocWritableHeaderFooterStory? ReadFooterStory(LegacyDocSourcePackage source, LegacyDocElement sectionProperties, string type, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes) {
            LegacyDocElement[] references = sectionProperties.Elements("footerReference").ToArray();
            if (references.Length == 0) {
                return null;
            }

            string kind = GetHeaderFooterStoryDescription(type, "footer");
            LegacyDocElement reference = GetSingleReference(references, type, "footer");
            if (reference.IsNone) {
                return null;
            }

            LegacyDocPart footerPart = GetReferencedHeaderOrFooterPart(source, reference.Attr(OfficeXmlNamespaces.Relationships, "id"), kind, header: false);
            return ReadSimpleHeaderFooterStory(footerPart.Root, footerPart, pictures, kind, styleIndexes, source);
        }

        /// <summary>
        /// The header (or footer) part a main-part relationship id points to. Throws the same errors the SDK lookup
        /// did: no id, an id without an internal part, and a part of another type.
        /// </summary>
        private static LegacyDocPart GetReferencedHeaderOrFooterPart(LegacyDocSourcePackage source, string? relationshipId, string kind, bool header) {
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                throw new NotSupportedException($"Native DOC saving cannot write a {kind} reference without a relationship id.");
            }

            string? partName = source.GetPartById(source.MainPartName, relationshipId);
            LegacyDocRelationship? relationship = partName == null ? null : source.FindRelationship(source.MainPartName, relationshipId);
            if (partName == null || relationship == null) {
                throw new NotSupportedException($"Native DOC saving cannot write a {kind} reference that points to a missing part.");
            }

            string expectedType = header ? LegacyDocSourcePackage.HeaderRelationshipType : LegacyDocSourcePackage.FooterRelationshipType;
            if (string.Equals(relationship.Type, expectedType, StringComparison.Ordinal)) {
                return source.Part(partName);
            }

            throw new NotSupportedException($"Native DOC saving cannot write a {kind} reference that points to an unexpected part type.");
        }

        private static LegacyDocElement GetSingleReference(IReadOnlyList<LegacyDocElement> references, string requestedType, string kind) {
            LegacyDocElement requestedReference = LegacyDocElement.None;
            foreach (LegacyDocElement reference in references) {
                string type = reference.Attr("type") ?? "default";
                if (!IsSupportedHeaderFooterType(type)) {
                    throw new NotSupportedException($"Native DOC saving supports only default, first-page, and even-page {kind}s.");
                }

                if (type != requestedType) {
                    continue;
                }

                if (requestedReference.Exists) {
                    throw new NotSupportedException($"Native DOC saving cannot write multiple {GetHeaderFooterDescription(requestedType, kind)} references in one section.");
                }

                requestedReference = reference;
            }

            return requestedReference;
        }

        private static bool IsSupportedHeaderFooterType(string type) {
            return type == "default"
                || type == "first"
                || type == "even";
        }

        private static string GetHeaderFooterDescription(string type, string kind) {
            if (type == "first") {
                return $"first-page {kind}";
            }

            if (type == "even") {
                return $"even-page {kind}";
            }

            return $"default {kind}";
        }

        private static string GetHeaderFooterStoryDescription(string type, string kind) {
            if (type == "default") {
                return kind;
            }

            return GetHeaderFooterDescription(type, kind);
        }

        private static LegacyDocWritableHeaderFooterStory? ReadSimpleHeaderFooterStory(LegacyDocElement container, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, string kind, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocSourcePackage source) {
            if (container.IsNone || !container.HasElements) {
                return LegacyDocWritableHeaderFooterStory.Empty;
            }

            var paragraphs = new List<string>();
            var storyText = new StringBuilder();
            var formattedRuns = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            foreach (LegacyDocElement child in container.Elements()) {
                AppendSimpleHeaderFooterStoryChild(storyText, formattedRuns, formattedParagraphs, bookmarks, paragraphs, child, relationshipOwner, pictures, kind, styleIndexes, source);
            }

            bool hasVisibleText = paragraphs.Any(paragraph => paragraph.Length > 0);
            if (!hasVisibleText && !container.HasDescendant("tbl") && !bookmarks.HasBookmarkMarkers) {
                return LegacyDocWritableHeaderFooterStory.Empty;
            }

            storyText.Append('\r');
            return new LegacyDocWritableHeaderFooterStory(storyText.ToString(), formattedRuns, formattedParagraphs, bookmarks.Create());
        }

        private static void AppendSimpleHeaderFooterStoryChild(
            StringBuilder storyText,
            List<LegacyDocWritableRun> formattedRuns,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            List<string> paragraphs,
            LegacyDocElement child,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            string kind,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            LegacyDocSourcePackage source) {
            switch (child.WName) {
                case "tbl":
                    AppendTable(storyText, formattedRuns, formattedParagraphs, bookmarks, child,
                        source, pictures, styleIndexes, ReadTableStyleDefinitions(source),
                        LegacyDocWritableFootnotes.Empty, LegacyDocWritableEndnotes.Empty,
                        relationshipOwner: relationshipOwner);
                    break;
                case "p":
                    AppendSimpleHeaderFooterParagraph(storyText, formattedRuns, formattedParagraphs, bookmarks, paragraphs, child, relationshipOwner, pictures, kind, styleIndexes);
                    break;
                case "sdt":
                    AppendSimpleHeaderFooterContentControl(storyText, formattedRuns, formattedParagraphs, bookmarks, paragraphs, child, relationshipOwner, pictures, kind, styleIndexes, source);
                    break;
                case "bookmarkStart":
                    bookmarks.AddStart(child, storyText.Length);
                    break;
                case "bookmarkEnd":
                    bookmarks.AddEnd(child, storyText.Length);
                    break;
                default:
                    throw new NotSupportedException($"Native DOC saving currently supports text paragraphs, tables, content controls, and bookmarks in {kind}s. Unsupported {kind} element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleHeaderFooterContentControl(
            StringBuilder storyText,
            List<LegacyDocWritableRun> formattedRuns,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            List<string> paragraphs,
            LegacyDocElement sdtBlock,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            string kind,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            LegacyDocSourcePackage source) {
            LegacyDocElement contentBlock = sdtBlock.Child("sdtContent");
            if (contentBlock.IsNone) {
                throw new NotSupportedException($"Native DOC saving supports {kind} content controls only when they contain simple paragraphs.");
            }

            foreach (LegacyDocElement child in contentBlock.Elements()) {
                AppendSimpleHeaderFooterStoryChild(storyText, formattedRuns, formattedParagraphs, bookmarks, paragraphs, child, relationshipOwner, pictures, kind, styleIndexes, source);
            }
        }

        private static void AppendSimpleHeaderFooterParagraph(
            StringBuilder storyText,
            List<LegacyDocWritableRun> formattedRuns,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            List<string> paragraphs,
            LegacyDocElement paragraph,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            string kind,
            IReadOnlyDictionary<string, ushort> styleIndexes) {
            int paragraphStart = storyText.Length;
            LegacyDocWritableFormatting paragraphMarkFormatting = ReadSupportedParagraphMarkRunFormatting(paragraph.Child("pPr"));
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSimpleHeaderFooterParagraph(storyText, formattedRuns, bookmarks, paragraph, relationshipOwner, pictures, kind, styleIndexes, out string paragraphText);
            paragraphs.Add(paragraphText);
            storyText.Append('\r');
            AddParagraphMarkRunFormatting(formattedRuns, storyText.Length - 1, paragraphMarkFormatting);
            if (paragraphFormatting.HasFormatting) {
                formattedParagraphs.Add(new LegacyDocWritableParagraph(paragraphStart, storyText.Length - paragraphStart, paragraphFormatting));
            }
        }

        private static LegacyDocWritableParagraphFormatting ReadSimpleHeaderFooterParagraph(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, LegacyDocWritableBookmarksBuilder bookmarks, LegacyDocElement paragraph, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, string kind, IReadOnlyDictionary<string, ushort> styleIndexes, out string paragraphText) {
            var text = new StringBuilder();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.Child("pPr"), styleIndexes);
            LegacyDocElement[] children = paragraph.Elements().ToArray();
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendFormattedHeaderFooterMathEquation(storyText, formattedRuns, text, child);
                    continue;
                }

                switch (child.WName) {
                    case "pPr":
                        break;
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendFormattedHeaderFooterComplexPageNumberField(storyText, formattedRuns, text, bookmarks, children, ref index, kind);
                        } else {
                            AppendFormattedHeaderFooterRun(storyText, formattedRuns, text, child, relationshipOwner, pictures, kind);
                        }

                        break;
                    case "ins":
                        AppendFormattedHeaderFooterRevision(storyText, formattedRuns, text, child, LegacyDocRevisionKind.Inserted, relationshipOwner, pictures, kind);
                        break;
                    case "del":
                        AppendFormattedHeaderFooterRevision(storyText, formattedRuns, text, child, LegacyDocRevisionKind.Deleted, relationshipOwner, pictures, kind);
                        break;
                    case "hyperlink":
                        AppendFormattedHeaderFooterHyperlink(storyText, formattedRuns, text, bookmarks, child, relationshipOwner, kind);
                        break;
                    case "fldSimple":
                        AppendFormattedHeaderFooterPageNumberField(storyText, formattedRuns, text, bookmarks, child, kind);
                        break;
                    case "sdt":
                        AppendFormattedHeaderFooterInlineContentControl(storyText, formattedRuns, text, bookmarks, child, relationshipOwner, pictures, kind);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, storyText.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, storyText.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving currently supports only text runs, {SupportedFieldNames} simple fields, bookmarks, inline content controls, and simple hyperlinks with supported direct formatting in {kind}s. Unsupported {kind} paragraph element: {child.LocalName}.");
                }
            }

            paragraphText = text.ToString();
            return paragraphFormatting;
        }

        private static void AppendFormattedHeaderFooterHyperlink(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, LegacyDocWritableBookmarksBuilder bookmarks, LegacyDocElement hyperlink, LegacyDocPart relationshipOwner, string kind) {
            int before = storyText.Length;
            try {
                AppendSupportedHyperlinkText(storyText, formattedRuns, bookmarks, hyperlink, relationshipOwner, LegacyDocWritableFootnotes.Empty, LegacyDocWritableEndnotes.Empty);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports simple {kind} hyperlinks only when they are external plain-text hyperlinks. {exception.Message}", exception);
            }

            if (storyText.Length > before) {
                paragraphText.Append(storyText.ToString(before, storyText.Length - before));
            }
        }

        private static void AppendFormattedHeaderFooterPageNumberField(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, LegacyDocWritableBookmarksBuilder bookmarks, LegacyDocElement simpleField, string kind) {
            int before = storyText.Length;
            try {
                AppendSupportedPageNumberFieldFromSimpleField(storyText, formattedRuns, bookmarks, simpleField, LegacyDocWritableFormatting.Plain);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports simple {kind} {SupportedFieldNames} fields only. {exception.Message}", exception);
            }

            if (storyText.Length > before) {
                paragraphText.Append(storyText.ToString(before, storyText.Length - before));
            }
        }

        private static void AppendFormattedHeaderFooterMathEquation(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, LegacyDocElement mathElement) {
            int before = storyText.Length;
            AppendMathEquationField(storyText, formattedRuns, mathElement, LegacyDocWritableFormatting.Plain);
            paragraphText.Append(storyText.ToString(before, storyText.Length - before));
        }

        private static void AppendFormattedHeaderFooterComplexPageNumberField(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, LegacyDocWritableBookmarksBuilder bookmarks, IReadOnlyList<LegacyDocElement> paragraphChildren, ref int childIndex, string kind) {
            int before = storyText.Length;
            try {
                AppendSupportedComplexPageNumberField(paragraphChildren, ref childIndex, storyText, formattedRuns, bookmarks, LegacyDocWritableFormatting.Plain);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports complex {kind} {SupportedFieldNames} fields only. {exception.Message}", exception);
            }

            if (storyText.Length > before) {
                paragraphText.Append(storyText.ToString(before, storyText.Length - before));
            }
        }

        private static void AppendFormattedHeaderFooterRevision(
            StringBuilder storyText,
            List<LegacyDocWritableRun> formattedRuns,
            StringBuilder paragraphText,
            LegacyDocElement revisionElement,
            LegacyDocRevisionKind revisionKind,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            string kind) {
            LegacyDocRevision revision = ReadSupportedRevision(revisionElement, revisionKind);
            foreach (LegacyDocElement child in revisionElement.Elements()) {
                if (child.Is("r")) {
                    AppendFormattedHeaderFooterRun(storyText, formattedRuns, paragraphText, child, relationshipOwner, pictures, kind, revision);
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports tracked insertions and deletions in {kind}s only when they contain text runs. Unsupported revision element: {child.LocalName}.");
            }
        }

        private static void AppendFormattedHeaderFooterRun(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, LegacyDocElement run, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, string kind, LegacyDocRevision revision = default) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.Child("rPr")).WithRevision(revision);
            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                        break;
                    case "lastRenderedPageBreak":
                        break;
                    case "pgNum":
                        AppendFormattedHeaderFooterPageNumberRun(storyText, formattedRuns, paragraphText, formatting);
                        break;
                    case "t":
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, child.Text, formatting);
                        break;
                    case "delText":
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, child.Text, formatting);
                        break;
                    case "tab":
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, "\t", formatting);
                        break;
                    case "cr":
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                        break;
                    case "noBreakHyphen":
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting);
                        break;
                    case "softHyphen":
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting);
                        break;
                    case "br":
                        AppendFormattedHeaderFooterBreak(storyText, formattedRuns, paragraphText, child, kind, formatting);
                        break;
                    case "drawing":
                        int picturePosition = storyText.Length;
                        int pictureDataOffset = pictures.AddInlinePicture(child, relationshipOwner);
                        storyText.Append('\u0001');
                        paragraphText.Append('\u0001');
                        formattedRuns.Add(new LegacyDocWritableRun(
                            picturePosition,
                            1,
                            LegacyDocWritableFormatting.SpecialCharacter.WithRevision(revision),
                            pictureDataOffset));
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only text, embedded inline pictures, PAGE fields, tabs, carriage returns, soft/no-break hyphens, and text-wrapping/page/column breaks in {kind}s. Unsupported {kind} run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendFormattedHeaderFooterPageNumberRun(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, LegacyDocWritableFormatting formatting) {
            int before = storyText.Length;
            AppendSupportedPageNumberField(storyText, formattedRuns, formatting);
            paragraphText.Append(storyText.ToString(before, storyText.Length - before));
        }

        private static void AppendFormattedHeaderFooterText(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, string? value, LegacyDocWritableFormatting formatting) {
            int before = storyText.Length;
            AppendFormattedNoteText(storyText, formattedRuns, value, formatting, storyStart: 0);
            if (storyText.Length > before) {
                paragraphText.Append(storyText.ToString(before, storyText.Length - before));
            }
        }

        private static void AppendFormattedHeaderFooterBreak(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, LegacyDocElement breakNode, string kind, LegacyDocWritableFormatting formatting) {
            string? breakType = breakNode.Attr("type");
            if (breakType == null || breakType == "textWrapping") {
                AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                return;
            }

            if (breakType == "page") {
                AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting);
                return;
            }

            if (breakType == "column") {
                AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.ColumnBreak.ToString(), formatting);
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple {kind}s only with text-wrapping, page, and column breaks.");
        }

        private static void ThrowIfUnreferencedHeaderFooterContent(LegacyDocSourcePackage source, IEnumerable<LegacyDocElement> sectionPropertiesCollection) {
            var referencedIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (LegacyDocElement sectionProperties in sectionPropertiesCollection) {
                foreach (LegacyDocElement reference in sectionProperties.Elements("headerReference")) {
                    string? id = reference.Attr(OfficeXmlNamespaces.Relationships, "id");
                    if (!string.IsNullOrWhiteSpace(id)) {
                        referencedIds.Add(id!);
                    }
                }

                foreach (LegacyDocElement reference in sectionProperties.Elements("footerReference")) {
                    string? id = reference.Attr(OfficeXmlNamespaces.Relationships, "id");
                    if (!string.IsNullOrWhiteSpace(id)) {
                        referencedIds.Add(id!);
                    }
                }
            }

            foreach (LegacyDocRelatedPart part in source.HeaderParts) {
                string id = part.Id;
                if (!referencedIds.Contains(id) && HasVisibleHeaderFooterContent(source.Root(part.PartName))) {
                    throw new NotSupportedException("Native DOC saving cannot preserve unreferenced header content.");
                }
            }

            foreach (LegacyDocRelatedPart part in source.FooterParts) {
                string id = part.Id;
                if (!referencedIds.Contains(id) && HasVisibleHeaderFooterContent(source.Root(part.PartName))) {
                    throw new NotSupportedException("Native DOC saving cannot preserve unreferenced footer content.");
                }
            }
        }

        private static bool HasVisibleHeaderFooterContent(LegacyDocElement container) {
            if (container.IsNone) return false;
            foreach (LegacyDocElement text in container.Descendants("t")) {
                if (!string.IsNullOrEmpty(text.Text)) return true;
            }

            return false;
        }

        private readonly struct LegacyDocWritableHeaderFooterStories {
            internal static LegacyDocWritableHeaderFooterStories Empty { get; } = new LegacyDocWritableHeaderFooterStories(string.Empty, Array.Empty<byte>(), Array.Empty<int>(), Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);

            internal LegacyDocWritableHeaderFooterStories(string text, byte[] plcfHdd, IReadOnlyList<int> markerPositions, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs, LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                PlcfHdd = plcfHdd;
                MarkerPositions = markerPositions;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal byte[] PlcfHdd { get; }

            internal IReadOnlyList<int> MarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }
        }

        private readonly struct LegacyDocWritableHeaderFooterStory {
            internal static LegacyDocWritableHeaderFooterStory Empty { get; } = new LegacyDocWritableHeaderFooterStory(string.Empty, Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);

            internal LegacyDocWritableHeaderFooterStory(string text, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs, LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }

            internal static LegacyDocWritableHeaderFooterStory Plain(string text) {
                return new LegacyDocWritableHeaderFooterStory(text, Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);
            }
        }
    }
}
