using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// A span of visible text in one story of a Word document: the body, a table, a header or footer, a footnote or an endnote.
    /// A range can start and end inside runs and can cross runs and paragraphs.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Get a range from <see cref="WordSearchResult.Ranges"/>, <see cref="WordBookmark.GetRange"/> or
    /// <see cref="WordParagraph.GetRange(int, int)"/>. Offsets count the paragraph's visible text: deleted text and field
    /// instructions are not counted, a tab is one character and a line break is <c>\n</c>.
    /// </para>
    /// <para>
    /// Replacing text through a range moves the other ranges of the same document that follow it, so you can replace
    /// every search hit in document order. A range whose text was replaced by another range, or edited outside the range
    /// API, throws <see cref="InvalidOperationException"/> when used.
    /// </para>
    /// <para>
    /// Edits that cannot be made safely throw <see cref="InvalidOperationException"/> and leave the text unchanged:
    /// replacing or linking field results, tracked insertions or deletions, equations, drawings, footnote references or
    /// part of a content control, and joining paragraphs across a table or a section break.
    /// </para>
    /// </remarks>
    public sealed class WordRange {
        private const string ParagraphSeparator = "\n";

        private readonly WordDocument _document;
        private readonly BookmarkStart? _anchor;
        private Paragraph _startParagraph;
        private int _start;
        private Paragraph _endParagraph;
        private int _end;
        private int _version;

        internal WordRange(WordDocument document, Paragraph startParagraph, int start, Paragraph endParagraph, int end, BookmarkStart? anchor = null) {
            _document = document;
            _startParagraph = startParagraph;
            _start = start;
            _endParagraph = endParagraph;
            _end = end;
            _anchor = anchor;
            _version = document.RangeEdits.Count;
        }

        /// <summary>Gets the visible text of the range. Paragraph breaks inside the range are returned as <c>\n</c>.</summary>
        public string Text {
            get {
                Sync();
                if (ReferenceEquals(_startParagraph, _endParagraph)) {
                    return WordTextMap.Build(_startParagraph).Text.Substring(_start, _end - _start);
                }

                List<Paragraph> paragraphs = GetParagraphs();
                var parts = new List<string>(paragraphs.Count);
                for (int index = 0; index < paragraphs.Count; index++) {
                    string text = WordTextMap.Build(paragraphs[index]).Text;
                    if (index == 0) text = text.Substring(_start);
                    else if (index == paragraphs.Count - 1) text = text.Substring(0, _end);
                    parts.Add(text);
                }
                return string.Join(ParagraphSeparator, parts);
            }
        }

        /// <summary>
        /// Returns a range from the start of this range to the end of <paramref name="end"/>, such as a span across several paragraphs.
        /// </summary>
        /// <param name="end">A range in the same story that does not start before this range.</param>
        /// <returns>The combined range.</returns>
        public WordRange ExtendTo(WordRange end) {
            ArgumentNullException.ThrowIfNull(end);
            if (!ReferenceEquals(end._document, _document)) throw new ArgumentException("The range belongs to another document.", nameof(end));
            Sync();
            end.Sync();
            if (Compare(_startParagraph, _start, end._endParagraph, end._end) > 0) {
                throw new ArgumentException("The end range must not start before this range.", nameof(end));
            }
            return new WordRange(_document, _startParagraph, _start, end._endParagraph, end._end);
        }

        /// <summary>
        /// Replaces the text of the range. The new text takes the formatting of the first replaced character, or of the
        /// character before an empty range. <c>\t</c> becomes a tab and <c>\n</c> a line break.
        /// </summary>
        /// <param name="text">Replacement text. An empty string deletes the range text.</param>
        /// <returns>This range, now covering the new text.</returns>
        public WordRange Replace(string text) {
            ArgumentNullException.ThrowIfNull(text);
            Sync();
            EditPlan plan = PlanEdit();
            Run? source = plan.FormatSource;
            RunProperties? properties = source?.RunProperties?.CloneNode(true) as RunProperties;
            List<OpenXmlElement> content = CreateTextContent(text, properties);
            ApplyEdit(content, stripHyperlinkStyle: source != null && source.Ancestors<Hyperlink>().Any());
            return this;
        }

        /// <summary>
        /// Replaces the text of the range with content added through a paragraph: formatted text, hyperlinks, images or fields.
        /// </summary>
        /// <param name="content">
        /// Adds the replacement content, for example <c>p =&gt; p.AddText("Total").SetBold()</c>. Only inline content is used;
        /// paragraph formatting set on <c>p</c> is ignored.
        /// </param>
        /// <returns>This range, now covering the new content.</returns>
        public WordRange Replace(Action<WordParagraph> content) {
            ArgumentNullException.ThrowIfNull(content);
            Sync();
            EditPlan plan = PlanEdit();
            List<OpenXmlElement> elements = BuildContent(content);
            if (plan.InsideHyperlink && elements.Any(element => element is Hyperlink || element.Descendants<Hyperlink>().Any())) {
                throw new InvalidOperationException("A hyperlink cannot be inserted inside another hyperlink.");
            }
            ApplyEdit(elements, stripHyperlinkStyle: false);
            return this;
        }

        /// <summary>
        /// Formats every run of the range. Runs are split at the range edges first, so text outside the range keeps its formatting.
        /// </summary>
        /// <param name="format">Formats one run, for example <c>r =&gt; r.SetBold().SetHighlight(...)</c>.</param>
        /// <returns>This range.</returns>
        public WordRange Format(Action<WordParagraph> format) {
            ArgumentNullException.ThrowIfNull(format);
            Sync();
            List<Paragraph> paragraphs = GetParagraphs();
            foreach (Paragraph paragraph in paragraphs) {
                (int from, int to) = Bounds(paragraph, paragraphs);
                foreach (WordTextAtom atom in WordTextMap.Build(paragraph).Atoms) {
                    if (atom.Length > 0 && atom.Start < to && atom.End > from &&
                        (atom.Hazard & (WordTextHazard.Field | WordTextHazard.TrackedChange)) != 0) {
                        throw Unsafe("it contains a field result or a tracked change");
                    }
                }
            }

            SplitBoundaries();
            foreach (Paragraph paragraph in paragraphs) {
                (int from, int to) = Bounds(paragraph, paragraphs);
                var runs = new List<Run>();
                foreach (WordTextAtom atom in WordTextMap.Build(paragraph).Atoms) {
                    if (atom.Run == null || atom.Length == 0 || atom.Start >= to || atom.End <= from) continue;
                    if ((atom.Hazard & WordTextHazard.Opaque) != 0) continue;
                    if (!runs.Any(run => ReferenceEquals(run, atom.Run))) runs.Add(atom.Run);
                }
                foreach (Run run in runs) format(new WordParagraph(_document, paragraph, run));
            }
            return this;
        }

        /// <summary>Adds a bookmark that spans the range.</summary>
        /// <param name="name">Bookmark name.</param>
        /// <returns>The created bookmark.</returns>
        public WordBookmark AddBookmark(string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Bookmark name cannot be empty.", nameof(name));
            Sync();
            string id = _document.GetNextBookmarkId().ToString(System.Globalization.CultureInfo.InvariantCulture);
            var start = new BookmarkStart { Name = name, Id = id };
            InsertMarkers(start, new BookmarkEnd { Id = id }, null);
            return new WordBookmark(_document, start.Ancestors<Paragraph>().FirstOrDefault() ?? _startParagraph, start);
        }

        /// <summary>Adds a comment on the range.</summary>
        /// <param name="author">Comment author.</param>
        /// <param name="initials">Author initials.</param>
        /// <param name="text">Comment text.</param>
        /// <returns>The created comment.</returns>
        public WordComment AddComment(string author, string initials, string text) {
            ArgumentNullException.ThrowIfNull(text);
            return AddComment(author, initials, paragraph => paragraph.AddText(text));
        }

        /// <summary>Adds a comment on the range with a formatted body.</summary>
        /// <param name="author">Comment author.</param>
        /// <param name="initials">Author initials.</param>
        /// <param name="content">
        /// Writes the comment body into its first paragraph, for example <c>p =&gt; p.AddText("Check ").AddText("this").SetBold()</c>.
        /// </param>
        /// <returns>The created comment.</returns>
        public WordComment AddComment(string author, string initials, Action<WordParagraph> content) {
            ArgumentNullException.ThrowIfNull(content);
            Sync();
            OpenXmlElement root = GetStoryRoot(_startParagraph);
            if (root is Header or Footer or Comment) {
                throw new InvalidOperationException("Word does not support comments in headers, footers or comments.");
            }
            ValidateMarkerBoundaries();
            WordComment comment = WordComment.Create(_document, author, initials, content);
            string id = comment.Id ?? throw new InvalidOperationException("The comment has no identifier.");
            InsertMarkers(new CommentRangeStart { Id = id }, new CommentRangeEnd { Id = id }, new Run(new CommentReference { Id = id }));
            return comment;
        }

        /// <summary>Turns the range text into a hyperlink to an external address. The text keeps its formatting and gets the Hyperlink character style.</summary>
        /// <param name="uri">Link target.</param>
        /// <returns>The created hyperlink.</returns>
        public WordHyperLink AddHyperLink(Uri uri) {
            ArgumentNullException.ThrowIfNull(uri);
            return AddHyperLink(paragraph => {
                OpenXmlPart part = WordHyperLink.ResolveRelationshipPart(_document, paragraph)
                    ?? throw new InvalidOperationException("Unable to resolve hyperlink relationship owner.");
                return new Hyperlink { Id = part.AddHyperlinkRelationship(uri, true).Id, History = true };
            });
        }

        /// <summary>Turns the range text into a link to a bookmark in the document. The text keeps its formatting and gets the Hyperlink character style.</summary>
        /// <param name="anchor">Name of the target bookmark.</param>
        /// <returns>The created hyperlink.</returns>
        public WordHyperLink AddHyperLink(string anchor) {
            if (string.IsNullOrWhiteSpace(anchor)) throw new ArgumentException("Anchor cannot be empty.", nameof(anchor));
            return AddHyperLink(_ => new Hyperlink { Anchor = anchor, History = true });
        }

        private WordHyperLink AddHyperLink(Func<Paragraph, Hyperlink> create) {
            Sync();
            if (!ReferenceEquals(_startParagraph, _endParagraph) || _start == _end) {
                throw Unsafe("a hyperlink needs non-empty text within one paragraph");
            }
            EditPlan plan = PlanEdit();
            if (plan.InsideHyperlink || plan.Removed.Any(element => element.Ancestors<Hyperlink>().Any())) {
                throw Unsafe("it is already part of a hyperlink");
            }

            SplitBoundaries();
            plan = PlanEdit();
            OpenXmlElement? parent = plan.StartNode?.Parent;
            List<OpenXmlElement> top = plan.TopLevel;
            if (parent == null || top.Any(element => !ReferenceEquals(element.Parent, parent))) {
                throw Unsafe("its text is split across different containers");
            }

            Hyperlink hyperlink = create(_startParagraph);
            top[0].InsertBeforeSelf(hyperlink);
            foreach (OpenXmlElement element in top) {
                element.Remove();
                hyperlink.Append(element);
                if (element is Run run) WordHyperLink.ApplyHyperlinkStyle(run);
            }
            return new WordHyperLink(_document, _startParagraph, hyperlink);
        }

        internal void Sync() {
            List<WordRangeEdit> edits = _document.RangeEdits;
            for (; _version < edits.Count; _version++) {
                WordRangeEdit edit = edits[_version];
                if (!edit.TryMap(ref _startParagraph, ref _start) || !edit.TryMap(ref _endParagraph, ref _end)) {
                    _version = edits.Count;
                    throw new InvalidOperationException("The range no longer exists: another edit replaced its text.");
                }
            }
            if (_startParagraph.Parent == null || _endParagraph.Parent == null ||
                _end > WordTextMap.Build(_endParagraph).Text.Length || _start > WordTextMap.Build(_startParagraph).Text.Length ||
                (ReferenceEquals(_startParagraph, _endParagraph) && _start > _end)) {
                throw new InvalidOperationException("The range no longer matches the document: its text was edited outside the range.");
            }
        }

        internal bool IsEmpty => ReferenceEquals(_startParagraph, _endParagraph) && _start == _end;

        /// <summary>Throws when <see cref="Replace(string)"/> cannot edit the range safely.</summary>
        internal void ValidateReplace() {
            Sync();
            PlanEdit();
        }

        private static InvalidOperationException Unsafe(string reason) =>
            new InvalidOperationException($"The range cannot be edited safely: {reason}.");

        private (int From, int To) Bounds(Paragraph paragraph, List<Paragraph> paragraphs) {
            int from = ReferenceEquals(paragraph, _startParagraph) ? _start : 0;
            int to = ReferenceEquals(paragraph, _endParagraph) ? _end : WordTextMap.Build(paragraph).Text.Length;
            return (from, to);
        }

        private List<Paragraph> GetParagraphs() {
            if (ReferenceEquals(_startParagraph, _endParagraph)) return new List<Paragraph> { _startParagraph };
            return GetParagraphsBetween(_startParagraph, _endParagraph)
                ?? throw new InvalidOperationException("The range no longer matches the document: its end comes before its start.");
        }

        internal static OpenXmlElement GetStoryRoot(Paragraph paragraph) {
            foreach (OpenXmlElement ancestor in paragraph.Ancestors()) {
                if (ancestor is Body or Header or Footer or Footnote or Endnote or Comment or TextBoxContent) return ancestor;
            }
            throw new InvalidOperationException("The paragraph is not part of a document story.");
        }

        /// <summary>Paragraphs from <paramref name="start"/> to <paramref name="end"/> in document order, or <c>null</c> when <paramref name="end"/> comes first.</summary>
        internal static List<Paragraph>? GetParagraphsBetween(Paragraph start, Paragraph end) {
            OpenXmlElement root = GetStoryRoot(start);
            if (!ReferenceEquals(GetStoryRoot(end), root)) {
                throw new ArgumentException("A range cannot span different stories, such as the body and a header.");
            }
            List<Paragraph>? result = null;
            foreach (Paragraph paragraph in root.Descendants<Paragraph>()) {
                if (result == null) {
                    if (ReferenceEquals(paragraph, end) && !ReferenceEquals(paragraph, start)) return null;
                    if (!ReferenceEquals(paragraph, start)) continue;
                    result = new List<Paragraph>();
                }
                if (!ReferenceEquals(GetStoryRoot(paragraph), root)) continue;
                result.Add(paragraph);
                if (ReferenceEquals(paragraph, end)) return result;
            }
            return null;
        }

        private static int Compare(Paragraph first, int firstOffset, Paragraph second, int secondOffset) {
            if (ReferenceEquals(first, second)) return firstOffset.CompareTo(secondOffset);
            return GetParagraphsBetween(first, second) != null ? -1 : 1;
        }

        // ---- Splitting ----

        private void SplitBoundaries() {
            SplitAt(_endParagraph, _end);
            SplitAt(_startParagraph, _start);
        }

        /// <summary>Makes the character at <paramref name="offset"/> start its own run. Text is unchanged.</summary>
        internal static void SplitAt(Paragraph paragraph, int offset) {
            WordTextMap map = WordTextMap.Build(paragraph);
            WordTextAtom? atom = map.AtomAt(offset);
            if (atom?.Run == null) return;
            OpenXmlElement element = atom.Element;
            if (atom.Start < offset) {
                if (element is not Text text) throw Unsafe("it starts or ends inside a non-text element");
                element = SplitText(text, offset - atom.Start);
            }
            SplitRunBefore(atom.Run, element);
        }

        private static Text SplitText(Text text, int index) {
            string visible = WordParagraph.ReadVisibleText(text);
            text.Text = visible.Substring(0, index);
            text.Space = SpaceProcessingModeValues.Preserve;
            var right = new Text(visible.Substring(index)) { Space = SpaceProcessingModeValues.Preserve };
            text.InsertAfterSelf(right);
            return right;
        }

        private static void SplitRunBefore(Run run, OpenXmlElement element) {
            bool hasContentBefore = false;
            for (OpenXmlElement? previous = element.PreviousSibling(); previous != null; previous = previous.PreviousSibling()) {
                if (previous is not RunProperties) {
                    hasContentBefore = true;
                    break;
                }
            }
            if (!hasContentBefore) return;

            var second = new Run();
            second.SetAttributes(run.GetAttributes());
            if (run.RunProperties != null) second.Append(run.RunProperties.CloneNode(true));
            var moved = new List<OpenXmlElement>();
            for (OpenXmlElement? current = element; current != null; current = current.NextSibling()) moved.Add(current);
            foreach (OpenXmlElement child in moved) {
                child.Remove();
                second.Append(child);
            }
            run.InsertAfterSelf(second);
        }

        // ---- Markers (bookmarks and comments) ----

        private void ValidateMarkerBoundaries() {
            if (IsEmpty) {
                WordTextMap map = WordTextMap.Build(_startParagraph);
                WordTextAtom? inside = map.AtomAt(_start);
                if (inside != null && inside.Start < _start && inside.Hazard != WordTextHazard.None) {
                    throw Unsafe("it falls inside a field result, tracked change or embedded object");
                }
                return;
            }
            WordTextAtom? first = FirstAtom();
            WordTextAtom? last = LastAtom();
            if ((first != null && (first.Hazard & (WordTextHazard.Field | WordTextHazard.TrackedChange)) != 0) ||
                (last != null && (last.Hazard & (WordTextHazard.Field | WordTextHazard.TrackedChange)) != 0)) {
                throw Unsafe("it starts or ends inside a field result or tracked change");
            }
        }

        private WordTextAtom? FirstAtom() => _start < WordTextMap.Build(_startParagraph).Text.Length
            ? WordTextMap.Build(_startParagraph).AtomAt(_start)
            : null;

        private WordTextAtom? LastAtom() => _end > 0 ? WordTextMap.Build(_endParagraph).AtomAt(_end - 1) : null;

        private void InsertMarkers(OpenXmlElement startMarker, OpenXmlElement endMarker, OpenXmlElement? afterEnd) {
            ValidateMarkerBoundaries();
            if (IsEmpty) {
                if (_start < WordTextMap.Build(_startParagraph).Text.Length) SplitAt(_startParagraph, _start);
                InsertAtPoint(startMarker);
                startMarker.InsertAfterSelf(endMarker);
            } else {
                SplitBoundaries();
                WordTextAtom? first = FirstAtom();
                if (first != null) first.InlineNode.InsertBeforeSelf(startMarker);
                else _startParagraph.Append(startMarker);

                WordTextAtom? last = LastAtom();
                if (last != null) last.InlineNode.InsertAfterSelf(endMarker);
                else InsertAtParagraphStart(_endParagraph, endMarker);
            }
            if (afterEnd != null) endMarker.InsertAfterSelf(afterEnd);
        }

        /// <summary>Inserts an element at the empty range position: after a bookmark anchor, after the previous character or at the paragraph start.</summary>
        private void InsertAtPoint(OpenXmlElement element) {
            if (_anchor?.Parent != null && ReferenceEquals(_anchor.Ancestors<Paragraph>().FirstOrDefault(), _startParagraph)) {
                _anchor.InsertAfterSelf(element);
                return;
            }
            WordTextMap map = WordTextMap.Build(_startParagraph);
            if (_start > 0) {
                map.AtomAt(_start - 1)!.InlineNode.InsertAfterSelf(element);
            } else if (map.Atoms.FirstOrDefault(atom => atom.Length > 0) is { } next) {
                next.InlineNode.InsertBeforeSelf(element);
            } else {
                _startParagraph.Append(element);
            }
        }

        private static void InsertAtParagraphStart(Paragraph paragraph, OpenXmlElement element) {
            if (paragraph.ParagraphProperties != null) paragraph.ParagraphProperties.InsertAfterSelf(element);
            else paragraph.PrependChild(element);
        }

        // ---- Replacement ----

        private sealed class EditPlan {
            internal readonly List<OpenXmlElement> Removed = new();
            internal readonly List<OpenXmlElement> Preserved = new();
            internal readonly List<OpenXmlElement> TopLevel = new();
            internal readonly List<Paragraph> RemovedParagraphs = new();
            internal readonly List<OpenXmlElement> BlockMarkers = new();
            internal OpenXmlElement? StartNode;
            internal Run? FormatSource;
            internal bool InsideHyperlink;
        }

        /// <summary>Classifies the content of the range and throws when removing it could damage the document.</summary>
        private EditPlan PlanEdit() {
            var plan = new EditPlan();
            List<Paragraph> paragraphs = GetParagraphs();
            Paragraph first = paragraphs[0];
            Paragraph last = paragraphs[paragraphs.Count - 1];
            bool multiple = paragraphs.Count > 1;

            if (multiple) {
                for (OpenXmlElement? sibling = first.NextSibling(); !ReferenceEquals(sibling, last); sibling = sibling.NextSibling()) {
                    if (sibling == null) throw Unsafe("it crosses a table or a different container");
                    if (sibling is Paragraph paragraph) {
                        plan.RemovedParagraphs.Add(paragraph);
                    } else if (IsMarker(sibling)) {
                        plan.BlockMarkers.Add(sibling);
                    } else {
                        throw Unsafe($"it crosses a {sibling.LocalName} element");
                    }
                }
                plan.RemovedParagraphs.Add(last);
                foreach (Paragraph removed in plan.RemovedParagraphs) {
                    if (removed.ParagraphProperties?.SectionProperties != null) throw Unsafe("it crosses a section break");
                }
            }

            WordTextMap firstMap = WordTextMap.Build(first);
            WordTextAtom? startAtom = _start < firstMap.Text.Length ? firstMap.AtomAt(_start) : null;
            WordTextAtom? previousAtom = _start > 0 ? firstMap.AtomAt(_start - 1) : null;
            // An empty range continues the formatting of the text before it, like typing does.
            WordTextAtom? formatAtom = IsEmpty ? previousAtom ?? startAtom : startAtom ?? previousAtom;
            plan.FormatSource = formatAtom?.Run;

            OpenXmlElement? startSdt = null;
            if (!IsEmpty || startAtom != null) {
                OpenXmlElement? reference = startAtom?.InlineNode ?? formatAtom?.InlineNode;
                startSdt = reference?.Ancestors<SdtRun>().FirstOrDefault();
                plan.InsideHyperlink = reference?.Ancestors<Hyperlink>().Any() == true;
            } else if (_anchor != null) {
                plan.InsideHyperlink = _anchor.Ancestors<Hyperlink>().Any();
            }
            if (multiple && startSdt != null) throw Unsafe("it crosses a content control");

            if (IsEmpty) {
                WordTextAtom? neighbour = _anchor != null ? null : previousAtom ?? startAtom;
                if (neighbour != null && (neighbour.Hazard & (WordTextHazard.Field | WordTextHazard.TrackedChange)) != 0) {
                    throw Unsafe("its position is inside a field result or tracked change");
                }
                return plan;
            }

            if (!multiple) {
                WordTextAtom lastAtom = firstMap.AtomAt(_end - 1)!;
                CollectZone(plan, first, firstMap, startAtom!.InlineNode, lastAtom.InlineNode, startSdt);
            } else {
                if (startAtom != null) CollectZone(plan, first, firstMap, startAtom.InlineNode, null, startSdt);
                for (int index = 1; index < paragraphs.Count - 1; index++) {
                    Paragraph middle = paragraphs[index];
                    CollectZone(plan, middle, WordTextMap.Build(middle), null, null, null);
                }
                WordTextMap lastMap = WordTextMap.Build(last);
                OpenXmlElement? lastNode = _end > 0 ? lastMap.AtomAt(_end - 1)!.InlineNode : null;
                if (lastNode != null) CollectZone(plan, last, lastMap, null, lastNode, null);
                else CollectLeadingMarkers(plan, last);
            }
            // A container covered by the range is removed, so new content lands outside it.
            if (plan.StartNode != null) plan.InsideHyperlink = plan.StartNode.Ancestors<Hyperlink>().Any();
            return plan;
        }

        /// <summary>
        /// Walks the content of <paramref name="paragraph"/> from <paramref name="startNode"/> (or the paragraph start)
        /// through <paramref name="endNode"/> (or the paragraph end).
        /// </summary>
        private static void CollectZone(EditPlan plan, Paragraph paragraph, WordTextMap map, OpenXmlElement? startNode, OpenXmlElement? endNode, OpenXmlElement? startSdt) {
            var hazardous = new HashSet<OpenXmlElement>();
            foreach (WordTextAtom atom in map.Atoms) {
                if (atom.Hazard != WordTextHazard.None) hazardous.Add(atom.InlineNode);
            }

            List<OpenXmlElement> elements = paragraph.Descendants().ToList();
            var positions = new Dictionary<OpenXmlElement, int>(elements.Count);
            for (int index = 0; index < elements.Count; index++) positions[elements[index]] = index;

            OpenXmlElement? first = startNode;
            if (first != null) {
                // Include containers that start with the range and end inside it.
                while (first.Parent is Hyperlink or CustomXmlRun && IsFirstContent(first) &&
                       (endNode == null || SubtreeEnd(first.Parent, positions) <= SubtreeEnd(endNode, positions))) {
                    first = first.Parent;
                }
                plan.StartNode ??= first;
            }

            int from = first != null ? positions[first] : 0;
            int to = endNode != null ? SubtreeEnd(endNode, positions) : elements.Count;
            int index2 = from;
            while (index2 < to) {
                OpenXmlElement element = elements[index2];
                int end = SubtreeEnd(element, positions);
                if (end > to) {
                    // A container that continues past the range: only its covered children are removed.
                    index2++;
                    continue;
                }

                bool topLevel = element.Parent == null || !positions.TryGetValue(element.Parent, out int parentPosition) || parentPosition < from;
                switch (element) {
                    case ParagraphProperties:
                        index2 = end;
                        continue;
                    case Run run:
                        ClassifyRun(plan, run, hazardous, startSdt, topLevel);
                        index2 = end;
                        continue;
                    case Hyperlink:
                    case CustomXmlRun:
                        if (topLevel) plan.TopLevel.Add(element);
                        index2++;
                        continue;
                    case CustomXmlProperties:
                        index2 = end;
                        continue;
                    case ProofError:
                        plan.Removed.Add(element);
                        if (topLevel) plan.TopLevel.Add(element);
                        index2 = end;
                        continue;
                }

                if (IsMarker(element)) {
                    plan.Preserved.Add(element);
                    if (topLevel) plan.TopLevel.Add(element);
                    index2 = end;
                    continue;
                }

                throw element switch {
                    SdtRun or SdtContentRun or SdtProperties => Unsafe("it covers a content control boundary"),
                    InsertedRun or DeletedRun or MoveToRun or MoveFromRun => Unsafe("it contains a tracked change"),
                    SimpleField => Unsafe("it contains a field"),
                    _ => Unsafe($"it contains a {element.LocalName} element")
                };
            }
        }

        private static void CollectLeadingMarkers(EditPlan plan, Paragraph paragraph) {
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                if (child is ParagraphProperties) continue;
                if (IsMarker(child)) plan.Preserved.Add(child);
            }
        }

        private static void ClassifyRun(EditPlan plan, Run run, HashSet<OpenXmlElement> hazardous, OpenXmlElement? startSdt, bool topLevel) {
            if (hazardous.Contains(run)) throw Unsafe("it contains a field result, tracked change, drawing or embedded object");
            foreach (OpenXmlElement ancestor in run.Ancestors()) {
                if (ancestor is Paragraph) break;
                if (ancestor is InsertedRun or DeletedRun or MoveToRun or MoveFromRun) throw Unsafe("it contains a tracked change");
                if (ancestor is SimpleField) throw Unsafe("it contains a field");
            }
            if (!ReferenceEquals(run.Ancestors<SdtRun>().FirstOrDefault(), startSdt)) throw Unsafe("it covers part of a content control");

            bool hasCommentReference = false;
            bool hasText = false;
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        continue;
                    case CommentReference:
                        hasCommentReference = true;
                        continue;
                    case DocumentFormat.OpenXml.Wordprocessing.Text:
                        hasText = true;
                        continue;
                }
                if (WordTextMap.IsPlainRunContent(child)) continue;
                throw child switch {
                    FieldChar or FieldCode => Unsafe("it contains a field"),
                    FootnoteReference or EndnoteReference => Unsafe("it contains a footnote or endnote reference"),
                    _ => Unsafe($"it contains a {child.LocalName} element")
                };
            }

            if (hasCommentReference && !hasText) plan.Preserved.Add(run);
            else plan.Removed.Add(run);
            if (topLevel) plan.TopLevel.Add(run);
        }

        private static bool IsMarker(OpenXmlElement element) =>
            element is BookmarkStart or BookmarkEnd or CommentRangeStart or CommentRangeEnd or PermStart or PermEnd;

        private static bool IsFirstContent(OpenXmlElement element) {
            for (OpenXmlElement? previous = element.PreviousSibling(); previous != null; previous = previous.PreviousSibling()) {
                if (previous is not CustomXmlProperties) return false;
            }
            return true;
        }

        private static int SubtreeEnd(OpenXmlElement element, Dictionary<OpenXmlElement, int> positions) {
            OpenXmlElement? last = element;
            while (last.LastChild != null) last = last.LastChild;
            return positions[last] + 1;
        }

        private List<OpenXmlElement> BuildContent(Action<WordParagraph> content) {
            var scratch = new Paragraph();
            _startParagraph.InsertAfterSelf(scratch);
            try {
                content(new WordParagraph(_document, scratch));
            } finally {
                scratch.Remove();
            }
            var elements = scratch.ChildElements.Where(child => child is not ParagraphProperties).ToList();
            foreach (OpenXmlElement element in elements) element.Remove();
            return elements;
        }

        private static List<OpenXmlElement> CreateTextContent(string text, RunProperties? properties) {
            var elements = new List<OpenXmlElement>();
            if (text.Length == 0) return elements;
            var run = new Run();
            if (properties != null) run.Append(properties);
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            int segmentStart = 0;
            for (int index = 0; index <= normalized.Length; index++) {
                if (index < normalized.Length && normalized[index] != '\t' && normalized[index] != '\n') continue;
                if (index > segmentStart) {
                    run.Append(new Text(normalized.Substring(segmentStart, index - segmentStart)) { Space = SpaceProcessingModeValues.Preserve });
                }
                if (index < normalized.Length) run.Append(normalized[index] == '\t' ? new TabChar() : new Break());
                segmentStart = index + 1;
            }
            elements.Add(run);
            return elements;
        }

        private void ApplyEdit(List<OpenXmlElement> content, bool stripHyperlinkStyle) {
            Paragraph first = _startParagraph;
            Paragraph last = _endParagraph;
            int start = _start;
            int end = _end;
            int suffixLength = WordTextMap.Build(last).Text.Length - end;
            List<Paragraph> removedParagraphs;

            var placeholder = new Run();
            if (IsEmpty) {
                if (_start < WordTextMap.Build(first).Text.Length) SplitAt(first, _start);
                InsertAtPoint(placeholder);
                removedParagraphs = new List<Paragraph>();
            } else {
                SplitBoundaries();
                EditPlan plan = PlanEdit();
                removedParagraphs = plan.RemovedParagraphs;
                if (plan.StartNode != null) plan.StartNode.InsertBeforeSelf(placeholder);
                else first.Append(placeholder);

                var touched = new HashSet<OpenXmlElement>();
                foreach (OpenXmlElement marker in plan.BlockMarkers.Concat(plan.Preserved)) {
                    if (marker.Parent != null) touched.Add(marker.Parent);
                    marker.Remove();
                    placeholder.InsertBeforeSelf(marker);
                }
                foreach (OpenXmlElement element in plan.Removed) {
                    if (element.Parent == null) continue;
                    touched.Add(element.Parent);
                    element.Remove();
                }

                if (removedParagraphs.Count > 0) {
                    var tail = last.ChildElements.Where(child => child is not ParagraphProperties).ToList();
                    foreach (OpenXmlElement element in tail) {
                        element.Remove();
                        first.Append(element);
                    }
                    foreach (Paragraph paragraph in removedParagraphs) paragraph.Remove();
                }

                foreach (OpenXmlElement container in touched) RemoveIfEmpty(container);
            }

            bool placeholderInHyperlink = placeholder.Ancestors<Hyperlink>().Any();
            foreach (OpenXmlElement element in content) {
                if (stripHyperlinkStyle && !placeholderInHyperlink && element is Run run &&
                    run.RunProperties?.RunStyle?.Val?.Value == "Hyperlink") {
                    run.RunProperties.RunStyle.Remove();
                }
                placeholder.InsertBeforeSelf(element);
            }
            placeholder.Remove();

            int insertedLength = WordTextMap.Build(first).Text.Length - start - suffixLength;
            _document.RangeEdits.Add(new WordRangeEdit(first, start, last, end, insertedLength, removedParagraphs));
            _startParagraph = first;
            _endParagraph = first;
            _end = start + insertedLength;
            _version = _document.RangeEdits.Count;
        }

        private static void RemoveIfEmpty(OpenXmlElement container) {
            while (container is Hyperlink or CustomXmlRun) {
                if (container.ChildElements.Any(child => child is not CustomXmlProperties)) return;
                OpenXmlElement? parent = container.Parent;
                container.Remove();
                if (parent == null) return;
                container = parent;
            }
        }
    }

    /// <summary>A text replacement made through a <see cref="WordRange"/>, used to move other ranges of the document.</summary>
    internal sealed class WordRangeEdit {
        private readonly Paragraph _startParagraph;
        private readonly int _start;
        private readonly Paragraph _endParagraph;
        private readonly int _end;
        private readonly int _insertedLength;
        private readonly List<Paragraph> _removedParagraphs;

        internal WordRangeEdit(Paragraph startParagraph, int start, Paragraph endParagraph, int end, int insertedLength, List<Paragraph> removedParagraphs) {
            _startParagraph = startParagraph;
            _start = start;
            _endParagraph = endParagraph;
            _end = end;
            _insertedLength = insertedLength;
            _removedParagraphs = removedParagraphs;
        }

        /// <summary>Moves a position recorded before this edit. Returns <c>false</c> when the edit removed it.</summary>
        internal bool TryMap(ref Paragraph paragraph, ref int offset) {
            if (ReferenceEquals(paragraph, _startParagraph)) {
                if (offset <= _start) return true;
                if (!ReferenceEquals(_startParagraph, _endParagraph) || offset < _end) return false;
                offset = offset - _end + _start + _insertedLength;
                return true;
            }
            if (ReferenceEquals(paragraph, _endParagraph)) {
                if (offset < _end) return false;
                paragraph = _startParagraph;
                offset = offset - _end + _start + _insertedLength;
                return true;
            }
            foreach (Paragraph removed in _removedParagraphs) {
                if (ReferenceEquals(paragraph, removed)) return false;
            }
            return true;
        }
    }
}
