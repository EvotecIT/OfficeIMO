using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Cuts a Word package into pieces. Every piece starts as a copy of the whole package, so settings, theme, fonts,
    /// properties and identifiers stay as they were; the body outside the piece is removed and every story, style,
    /// list and part the remaining content no longer uses is pruned.
    /// </summary>
    internal static class WordDocumentSplit {
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static readonly HashSet<string> StyleReferenceElements = new(StringComparer.Ordinal) {
            "pStyle", "rStyle", "tblStyle", "styleLink", "numStyleLink"
        };

        internal static IReadOnlyList<WordDocument> Split(WordDocument source, int? headingLevel) {
            byte[] package = source.CreatePackageSnapshot();
            List<(int Start, int End)> ranges;
            using (var stream = new MemoryStream(package, writable: false))
            using (WordprocessingDocument document = WordprocessingDocument.Open(stream, false)) {
                MainDocumentPart main = document.MainDocumentPart ?? throw new InvalidOperationException("The document has no main document part.");
                Body body = main.Document?.Body ?? throw new InvalidOperationException("The document has no body.");
                List<OpenXmlElement> content = GetContent(body);
                ranges = headingLevel == null
                    ? GetSectionRanges(content)
                    : GetHeadingRanges(content, main, headingLevel.Value);
            }

            var pieces = new List<WordDocument>(ranges.Count);
            try {
                foreach ((int start, int end) in ranges) {
                    pieces.Add(CreatePiece(package, start, end));
                }
            } catch {
                foreach (WordDocument piece in pieces) piece.Dispose();
                throw;
            }

            return pieces;
        }

        private static List<OpenXmlElement> GetContent(Body body) =>
            body.ChildElements.Where(element => element is not SectionProperties).ToList();

        private static SectionProperties? GetSectionBreak(OpenXmlElement element) =>
            (element as Paragraph)?.ParagraphProperties?.SectionProperties;

        private static List<(int Start, int End)> GetSectionRanges(List<OpenXmlElement> content) {
            var ranges = new List<(int, int)>();
            int start = 0;
            for (int i = 0; i < content.Count; i++) {
                if (GetSectionBreak(content[i]) != null) {
                    ranges.Add((start, i));
                    start = i + 1;
                }
            }

            if (start < content.Count) ranges.Add((start, content.Count - 1));
            return ranges;
        }

        private static List<(int Start, int End)> GetHeadingRanges(List<OpenXmlElement> content, MainDocumentPart main, int level) {
            var styles = new Dictionary<string, Style>(StringComparer.Ordinal);
            foreach (Style style in main.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>()) {
                if (style.StyleId?.Value is { Length: > 0 } id && !styles.ContainsKey(id)) styles.Add(id, style);
            }

            var starts = new List<int>();
            for (int i = 0; i < content.Count; i++) {
                if (content[i] is Paragraph paragraph) {
                    int paragraphLevel = WordDocument.GetHeadingLevel(paragraph, styles);
                    if (paragraphLevel >= 1 && paragraphLevel <= level) starts.Add(i);
                }
            }

            var ranges = new List<(int, int)>();
            int firstStart = starts.Count > 0 ? starts[0] : content.Count;
            if (firstStart > 0 && content.Take(firstStart).Any(HasContent)) ranges.Add((0, firstStart - 1));
            for (int i = 0; i < starts.Count; i++) {
                int end = i + 1 < starts.Count ? starts[i + 1] - 1 : content.Count - 1;
                ranges.Add((starts[i], end));
            }

            return ranges;
        }

        private static bool HasContent(OpenXmlElement element) =>
            element is Table ||
            element is AltChunk ||
            element.Descendants<AltChunk>().Any() ||
            !string.IsNullOrWhiteSpace(element.InnerText) ||
            element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() ||
            element.Descendants<Picture>().Any() ||
            element.Descendants<EmbeddedObject>().Any();

        private static WordDocument CreatePiece(byte[] package, int start, int end) {
            using var stream = new MemoryStream();
            stream.Write(package, 0, package.Length);
            stream.Position = 0;
            using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = false })) {
                MainDocumentPart main = document.MainDocumentPart!;
                Body body = main.Document!.Body!;
                CutBody(body, start, end);
                RemoveOrphanedMarkers(body);
                BalanceFields(body);
                main.Document.Save();
                Prune(main);
                document.Save();
            }

            stream.Position = 0;
            return WordDocument.Load(stream, new WordLoadOptions {
                AccessMode = DocumentAccessMode.ReadWrite,
                PersistenceMode = DocumentPersistenceMode.Explicit
            });
        }

        /// <summary>
        /// Keeps the elements from <paramref name="start"/> to <paramref name="end"/>. The section that holds the last
        /// element closes the body, and the first section inherits the headers and footers it had in the source.
        /// </summary>
        private static void CutBody(Body body, int start, int end) {
            List<OpenXmlElement> content = GetContent(body);
            SectionProperties? finalSection = body.Elements<SectionProperties>().LastOrDefault();

            // Sections in source order: index of the closing element (or content.Count for the final one) and its properties.
            var sections = new List<(int End, SectionProperties? Properties)>();
            for (int i = 0; i < content.Count; i++) {
                if (GetSectionBreak(content[i]) is SectionProperties properties) sections.Add((i, properties));
            }
            sections.Add((content.Count, finalSection));

            int SectionOf(int index) => sections.FindIndex(section => section.End >= index);
            int firstSection = SectionOf(start);
            int lastSection = SectionOf(end);

            SectionProperties? closing;
            if (sections[lastSection].End == end && sections[lastSection].Properties is SectionProperties breakProperties && breakProperties.Parent is ParagraphProperties) {
                breakProperties.Remove();
                closing = breakProperties;
            } else {
                closing = (SectionProperties?)sections[lastSection].Properties?.CloneNode(true);
            }

            SectionProperties? first = firstSection == lastSection ? closing : sections[firstSection].Properties;
            if (first != null) InheritHeadersAndFooters(first, sections.Take(firstSection).Select(section => section.Properties));

            for (int i = 0; i < content.Count; i++) {
                if (i < start || i > end) content[i].Remove();
            }

            finalSection?.Remove();
            if (closing != null) body.Append(closing);
        }

        /// <summary>
        /// A section without its own header or footer reference shows the one from the previous section.
        /// The piece has no previous section, so the inherited references are written explicitly.
        /// </summary>
        private static void InheritHeadersAndFooters(SectionProperties target, IEnumerable<SectionProperties?> earlier) {
            var inherited = new Dictionary<(Type, HeaderFooterValues), OpenXmlElement>();
            foreach (SectionProperties? section in earlier) {
                if (section == null) continue;
                foreach (HeaderReference reference in section.Elements<HeaderReference>()) inherited[(typeof(HeaderReference), reference.Type?.Value ?? HeaderFooterValues.Default)] = reference;
                foreach (FooterReference reference in section.Elements<FooterReference>()) inherited[(typeof(FooterReference), reference.Type?.Value ?? HeaderFooterValues.Default)] = reference;
            }

            OpenXmlElement? anchor = target.ChildElements.FirstOrDefault(child => child is not HeaderReference && child is not FooterReference);
            foreach (KeyValuePair<(Type Kind, HeaderFooterValues Type), OpenXmlElement> pair in inherited) {
                bool present = pair.Key.Kind == typeof(HeaderReference)
                    ? target.Elements<HeaderReference>().Any(reference => (reference.Type?.Value ?? HeaderFooterValues.Default) == pair.Key.Type)
                    : target.Elements<FooterReference>().Any(reference => (reference.Type?.Value ?? HeaderFooterValues.Default) == pair.Key.Type);
                if (present) continue;
                OpenXmlElement clone = pair.Value.CloneNode(true);
                if (anchor != null) target.InsertBefore(clone, anchor); else target.Append(clone);
            }
        }

        /// <summary>Removes bookmark and comment range markers whose partner fell outside the piece.</summary>
        private static void RemoveOrphanedMarkers(Body body) {
            var bookmarkStarts = body.Descendants<BookmarkStart>().ToList();
            var bookmarkEnds = body.Descendants<BookmarkEnd>().ToList();
            var startIds = new HashSet<string>(bookmarkStarts.Select(item => item.Id?.Value).OfType<string>(), StringComparer.Ordinal);
            var endIds = new HashSet<string>(bookmarkEnds.Select(item => item.Id?.Value).OfType<string>(), StringComparer.Ordinal);
            foreach (BookmarkStart item in bookmarkStarts) if (item.Id?.Value is not string id || !endIds.Contains(id)) item.Remove();
            foreach (BookmarkEnd item in bookmarkEnds) if (item.Id?.Value is not string id || !startIds.Contains(id)) item.Remove();

            var rangeStarts = body.Descendants<CommentRangeStart>().ToList();
            var rangeEnds = body.Descendants<CommentRangeEnd>().ToList();
            var commentStartIds = new HashSet<string>(rangeStarts.Select(item => item.Id?.Value).OfType<string>(), StringComparer.Ordinal);
            var commentEndIds = new HashSet<string>(rangeEnds.Select(item => item.Id?.Value).OfType<string>(), StringComparer.Ordinal);
            var referenced = new HashSet<string>(body.Descendants<CommentReference>().Select(item => item.Id?.Value).OfType<string>(), StringComparer.Ordinal);
            foreach (CommentRangeStart item in rangeStarts) {
                if (item.Id?.Value is not string id || !commentEndIds.Contains(id) || !referenced.Contains(id)) item.Remove();
            }
            foreach (CommentRangeEnd item in rangeEnds) {
                if (item.Id?.Value is not string id || !commentStartIds.Contains(id) || !referenced.Contains(id)) item.Remove();
            }
        }

        /// <summary>
        /// A complex field that started before the piece loses its separator and end marks; one that runs past the
        /// piece is closed at the end of the last paragraph.
        /// </summary>
        private static void BalanceFields(Body body) {
            int depth = 0;
            foreach (FieldChar mark in body.Descendants<FieldChar>().ToList()) {
                FieldCharValues? type = mark.FieldCharType?.Value;
                if (type == FieldCharValues.Begin) {
                    depth++;
                } else if (depth == 0) {
                    mark.Remove();
                } else if (type == FieldCharValues.End) {
                    depth--;
                }
            }

            if (depth == 0) return;
            Paragraph? last = body.Descendants<Paragraph>().LastOrDefault();
            if (last == null) {
                last = new Paragraph();
                SectionProperties? closing = body.Elements<SectionProperties>().LastOrDefault();
                if (closing != null) body.InsertBefore(last, closing); else body.Append(last);
            }

            for (; depth > 0; depth--) {
                last.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            }
        }

        private static void Prune(MainDocumentPart main) {
            Document document = main.Document!;
            PruneRelationships(main, document);
            PruneNotes(main.FootnotesPart, main.FootnotesPart?.Footnotes, document.Descendants<FootnoteReference>().Select(item => item.Id?.Value));
            PruneNotes(main.EndnotesPart, main.EndnotesPart?.Endnotes, document.Descendants<EndnoteReference>().Select(item => item.Id?.Value));
            PruneComments(main, document);
            UnwrapBrokenLinks(main, document);
            PrunePeople(main, document);
            PruneStylesAndNumbering(main, document);
        }

        /// <summary>
        /// A link to a bookmark that is now in another piece keeps its text but stops being a link. Bookmarks are looked
        /// up in every story the piece kept, since notes and comments can hold link targets too.
        /// </summary>
        private static void UnwrapBrokenLinks(MainDocumentPart main, Document document) {
            var stories = new List<OpenXmlElement> { document };
            foreach (OpenXmlPart story in StoryParts(main)) if (story.RootElement != null) stories.Add(story.RootElement);
            if (main.WordprocessingCommentsPart?.Comments is Comments comments) stories.Add(comments);

            var bookmarkNames = new HashSet<string>(
                stories.SelectMany(story => story.Descendants<BookmarkStart>()).Select(item => item.Name?.Value).OfType<string>(),
                StringComparer.Ordinal);
            foreach (Hyperlink link in stories.SelectMany(story => story.Descendants<Hyperlink>()).ToList()) {
                if (link.Id == null && link.Anchor?.Value is string anchor && !bookmarkNames.Contains(anchor)) {
                    foreach (OpenXmlElement child in link.ChildElements.ToList()) {
                        child.Remove();
                        link.InsertBeforeSelf(child);
                    }
                    link.Remove();
                }
            }
        }

        /// <summary>
        /// Removes the people records (people.xml) of authors whose comments and tracked changes are no longer in the
        /// piece, so a piece does not name reviewers of the parts it left out.
        /// </summary>
        private static void PrunePeople(MainDocumentPart main, Document document) {
            WordprocessingPeoplePart? part = main.WordprocessingPeoplePart;
            OpenXmlElement? people = part?.RootElement;
            if (part == null || people == null) return;

            var stories = new List<OpenXmlElement> { document };
            foreach (OpenXmlPart story in StoryParts(main)) if (story.RootElement != null) stories.Add(story.RootElement);
            if (main.WordprocessingCommentsPart?.Comments is Comments comments) stories.Add(comments);
            var authors = new HashSet<string>(StringComparer.Ordinal);
            foreach (OpenXmlElement story in stories) {
                foreach (OpenXmlElement element in story.Descendants()) {
                    if (GetAttribute(element, "author") is string author) authors.Add(author);
                }
            }

            foreach (OpenXmlElement person in people.ChildElements.ToList()) {
                if (GetAttribute(person, "author") is not string author || !authors.Contains(author)) person.Remove();
            }
            if (!people.HasChildren) main.DeletePart(part);
        }

        /// <summary>
        /// Deletes the content parts and hyperlink targets that a part's markup no longer refers to. Parts reached
        /// through implicit relationships (styles, settings, notes, comments, ...) are kept.
        /// </summary>
        private static void PruneRelationships(OpenXmlPart part, OpenXmlElement? root) {
            if (root == null) return;
            var referenced = new HashSet<string>(StringComparer.Ordinal);
            foreach (OpenXmlElement element in root.Descendants()) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (attribute.Value != null) referenced.Add(attribute.Value);
                }
            }

            foreach (IdPartPair pair in part.Parts.ToList()) {
                if (IsContentPart(pair.OpenXmlPart) && !referenced.Contains(pair.RelationshipId)) part.DeletePart(pair.RelationshipId);
            }

            foreach (HyperlinkRelationship hyperlink in part.HyperlinkRelationships.ToList()) {
                if (!referenced.Contains(hyperlink.Id)) part.DeleteReferenceRelationship(hyperlink);
            }

            foreach (ExternalRelationship external in part.ExternalRelationships.ToList()) {
                if (!referenced.Contains(external.Id)) part.DeleteExternalRelationship(external);
            }
        }

        private static bool IsContentPart(OpenXmlPart part) => part is HeaderPart or FooterPart or ImagePart or ChartPart
            or ExtendedChartPart or EmbeddedPackagePart or EmbeddedObjectPart or EmbeddedControlPersistencePart
            or DiagramDataPart or DiagramLayoutDefinitionPart or DiagramStylePart or DiagramColorsPart or DiagramPersistLayoutPart
            or AlternativeFormatImportPart;

        private static void PruneNotes(OpenXmlPart? part, OpenXmlElement? notes, IEnumerable<long?> references) {
            if (part == null || notes == null) return;
            var used = new HashSet<long>(references.Where(id => id != null).Select(id => id!.Value));
            foreach (FootnoteEndnoteType note in notes.Elements<FootnoteEndnoteType>().ToList()) {
                // Separator and continuation notes have a type and are referenced by the settings, not the body.
                if (note.Type == null && (note.Id?.Value is not long id || !used.Contains(id))) note.Remove();
            }

            PruneRelationships(part, notes);
        }

        private static void PruneComments(MainDocumentPart main, Document document) {
            WordprocessingCommentsPart? part = main.WordprocessingCommentsPart;
            Comments? comments = part?.Comments;
            if (part == null || comments == null) return;

            var used = new HashSet<string>(document.Descendants<CommentReference>().Select(item => item.Id?.Value).OfType<string>(), StringComparer.Ordinal);
            foreach (OpenXmlPart story in StoryParts(main)) {
                if (story.RootElement != null) used.UnionWith(story.RootElement.Descendants<CommentReference>().Select(item => item.Id?.Value).OfType<string>());
            }

            // Replies are threaded to their parent through commentsExtended and may have no anchor of their own.
            var parents = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            OpenXmlElement? extended = main.WordprocessingCommentsExPart?.RootElement;
            foreach (OpenXmlElement entry in extended?.ChildElements ?? Enumerable.Empty<OpenXmlElement>()) {
                if (GetAttribute(entry, "paraId") is string paragraph && GetAttribute(entry, "paraIdParent") is string parent) parents[paragraph] = parent;
            }

            var keptParagraphs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<Comment> all = comments.Elements<Comment>().ToList();
            var kept = new HashSet<Comment>();
            bool added;
            do {
                added = false;
                foreach (Comment comment in all) {
                    if (kept.Contains(comment)) continue;
                    List<string> paragraphIds = comment.Elements<Paragraph>().Select(item => item.ParagraphId?.Value).OfType<string>().ToList();
                    bool isUsed = comment.Id?.Value is string id && used.Contains(id);
                    bool isReply = paragraphIds.Any(paragraph => parents.TryGetValue(paragraph, out string? parent) && keptParagraphs.Contains(parent));
                    if (isUsed || isReply) {
                        kept.Add(comment);
                        keptParagraphs.UnionWith(paragraphIds);
                        added = true;
                    }
                }
            } while (added);

            foreach (Comment comment in all) if (!kept.Contains(comment)) comment.Remove();
            PruneRelationships(part, comments);

            if (extended != null) RemoveEntries(extended, entry => GetAttribute(entry, "paraId") is string id && keptParagraphs.Contains(id));
            OpenXmlElement? ids = main.WordprocessingCommentsIdsPart?.RootElement;
            var durableIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (ids != null) {
                RemoveEntries(ids, entry => GetAttribute(entry, "paraId") is string id && keptParagraphs.Contains(id));
                durableIds.UnionWith(ids.ChildElements.Select(entry => GetAttribute(entry, "durableId")).OfType<string>());
            }
            OpenXmlElement? extensible = main.WordCommentsExtensiblePart?.RootElement;
            if (extensible != null) RemoveEntries(extensible, entry => GetAttribute(entry, "durableId") is string id && durableIds.Contains(id));
        }

        private static void RemoveEntries(OpenXmlElement root, Func<OpenXmlElement, bool> keep) {
            foreach (OpenXmlElement entry in root.ChildElements.ToList()) {
                if (GetAttribute(entry, "paraId") != null || GetAttribute(entry, "durableId") != null) {
                    if (!keep(entry)) entry.Remove();
                }
            }
        }

        private static string? GetAttribute(OpenXmlElement element, string localName) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)) return attribute.Value;
            }
            return null;
        }

        private static IEnumerable<OpenXmlPart> StoryParts(MainDocumentPart main) {
            foreach (HeaderPart part in main.HeaderParts) yield return part;
            foreach (FooterPart part in main.FooterParts) yield return part;
            if (main.FootnotesPart != null) yield return main.FootnotesPart;
            if (main.EndnotesPart != null) yield return main.EndnotesPart;
        }

        /// <summary>
        /// Keeps the styles and lists the remaining stories use, the styles they depend on, and the default styles.
        /// Styles and lists refer to each other, so both are resolved together until nothing new is found.
        /// </summary>
        private static void PruneStylesAndNumbering(MainDocumentPart main, Document document) {
            Styles? styles = main.StyleDefinitionsPart?.Styles;
            Numbering? numbering = main.NumberingDefinitionsPart?.Numbering;

            var stories = new List<OpenXmlElement> { document };
            foreach (OpenXmlPart part in StoryParts(main)) if (part.RootElement != null) stories.Add(part.RootElement);
            if (main.WordprocessingCommentsPart?.Comments is Comments comments) stories.Add(comments);

            var stylesById = new Dictionary<string, Style>(StringComparer.Ordinal);
            foreach (Style style in styles?.Elements<Style>() ?? Enumerable.Empty<Style>()) {
                if (style.StyleId?.Value is { Length: > 0 } id && !stylesById.ContainsKey(id)) stylesById.Add(id, style);
            }
            var instances = new Dictionary<int, NumberingInstance>();
            foreach (NumberingInstance instance in numbering?.Elements<NumberingInstance>() ?? Enumerable.Empty<NumberingInstance>()) {
                if (instance.NumberID?.Value is int id && !instances.ContainsKey(id)) instances.Add(id, instance);
            }
            var abstracts = new Dictionary<int, AbstractNum>();
            foreach (AbstractNum definition in numbering?.Elements<AbstractNum>() ?? Enumerable.Empty<AbstractNum>()) {
                if (definition.AbstractNumberId?.Value is int id && !abstracts.ContainsKey(id)) abstracts.Add(id, definition);
            }

            var usedStyles = new HashSet<string>(StringComparer.Ordinal);
            var usedNumbers = new HashSet<int>();
            var usedAbstracts = new HashSet<int>();
            var pendingStyles = new Queue<string>();
            var pendingElements = new Queue<OpenXmlElement>(stories);
            foreach (Style style in stylesById.Values) {
                if (style.Default?.Value == true) pendingStyles.Enqueue(style.StyleId!.Value!);
            }

            while (pendingStyles.Count > 0 || pendingElements.Count > 0) {
                while (pendingElements.Count > 0) {
                    OpenXmlElement root = pendingElements.Dequeue();
                    foreach (OpenXmlElement element in root.Descendants()) {
                        if (!string.Equals(element.NamespaceUri, WordNamespace, StringComparison.Ordinal)) continue;
                        if (StyleReferenceElements.Contains(element.LocalName) && GetAttribute(element, "val") is { Length: > 0 } styleId) {
                            pendingStyles.Enqueue(styleId);
                        } else if (element is NumberingId numberingId && numberingId.Val?.Value is int numberId && numberId > 0 && usedNumbers.Add(numberId) &&
                                   instances.TryGetValue(numberId, out NumberingInstance? instance)) {
                            // Level overrides on the instance can name styles too.
                            pendingElements.Enqueue(instance);
                            if (instance.AbstractNumId?.Val?.Value is int abstractId && usedAbstracts.Add(abstractId) &&
                                abstracts.TryGetValue(abstractId, out AbstractNum? definition)) {
                                pendingElements.Enqueue(definition);
                            }
                        }
                    }
                }

                while (pendingStyles.Count > 0) {
                    string id = pendingStyles.Dequeue();
                    if (!usedStyles.Add(id) || !stylesById.TryGetValue(id, out Style? style)) continue;
                    foreach (string? related in new[] { style.BasedOn?.Val?.Value, style.LinkedStyle?.Val?.Value, style.NextParagraphStyle?.Val?.Value }) {
                        if (!string.IsNullOrEmpty(related)) pendingStyles.Enqueue(related!);
                    }
                    pendingElements.Enqueue(style);
                }
            }

            foreach (Style style in stylesById.Values) {
                if (!usedStyles.Contains(style.StyleId!.Value!)) style.Remove();
            }

            if (numbering == null) return;
            foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>().ToList()) {
                if (instance.NumberID?.Value is not int id || !usedNumbers.Contains(id)) instance.Remove();
            }
            foreach (AbstractNum definition in numbering.Elements<AbstractNum>().ToList()) {
                if (definition.AbstractNumberId?.Value is not int id || !usedAbstracts.Contains(id)) definition.Remove();
            }

            var usedPictureBullets = new HashSet<int>(numbering.Descendants<LevelPictureBulletId>().Select(item => item.Val?.Value).OfType<int>());
            foreach (NumberingPictureBullet bullet in numbering.Elements<NumberingPictureBullet>().ToList()) {
                if (bullet.NumberingPictureBulletId?.Value is not int id || !usedPictureBullets.Contains(id)) bullet.Remove();
            }
            PruneRelationships(main.NumberingDefinitionsPart!, numbering);
        }
    }
}
