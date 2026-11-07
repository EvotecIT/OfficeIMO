using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Replaces each occurrence of text in consecutive body paragraphs with a block-level
        /// AltChunk at the match location, splitting a paragraph when necessary.
        /// </summary>
        /// <param name="textToFind">Text to search for.</param>
        /// <param name="htmlContent">HTML fragment to insert.</param>
        /// <param name="type">Optional format type of the fragment.</param>
        /// <param name="stringComparison">String comparison option for the search.</param>
        /// <returns>The number of replacements performed.</returns>
        /// <exception cref="NotSupportedException">
        /// A match intersects content that cannot be split or removed as body text, such as
        /// a field, content control, revision, or explicit page or column break. The operation
        /// also rejects a match that would move a paragraph's page-break-before boundary.
        /// </exception>
        public int ReplaceTextWithHtmlFragment(string textToFind, string htmlContent,
            WordAlternativeFormatImportPartType type = WordAlternativeFormatImportPartType.Html,
            StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrEmpty(textToFind)) {
                throw new ArgumentNullException(nameof(textToFind));
            }

            MainDocumentPart mainPart = _document.MainDocumentPart ?? throw new InvalidOperationException("The document does not contain a main document part.");
            Document? document = mainPart.Document;
            if (document?.Body is not Body body) {
                return 0;
            }

            // Plan all edits first so an unsupported match leaves both document XML and parts intact.
            var matches = new List<HtmlFragmentMatch>();
            var paragraphs = new List<WordParagraph>();
            foreach (OpenXmlElement block in body.ChildElements) {
                if (block is Paragraph paragraph) {
                    paragraphs.AddRange(WordSection.ConvertParagraphToWordParagraphs(this, paragraph));
                } else {
                    CollectHtmlFragmentMatches(paragraphs, textToFind, stringComparison, matches);
                    paragraphs.Clear();
                }
            }
            CollectHtmlFragmentMatches(paragraphs, textToFind, stringComparison, matches);

            if (matches.Count == 0) {
                return 0;
            }

            EnsureMatchesOutsideComplexFields(body, matches);
            Styles? styles = mainPart.StyleDefinitionsPart?.Styles ?? mainPart.StylesWithEffectsPart?.Styles;
            EnsureHtmlFragmentsPreservePageBreaks(matches, styles);

            var fragments = new List<(AltChunk Anchor, AlternativeFormatImportPart Part)>(matches.Count);
            var linkSplits = new HyperlinkRelationship?[matches.Count];
            var changedLinkIds = new HashSet<string>(StringComparer.Ordinal);
            var authoredLinks = mainPart.HyperlinkRelationships.ToDictionary(link => link.Id, StringComparer.Ordinal);
            try {
                for (int index = matches.Count - 1; index >= 0; index--) {
                    fragments.Add(CreateEmbeddedFragment(mainPart, htmlContent, type));
                    if (matches[index].Slices[0].Node.Parent?.Parent is Hyperlink hyperlink
                        && hyperlink.Id?.Value is string linkId) {
                        if (!authoredLinks.TryGetValue(linkId, out HyperlinkRelationship? authoredLink)) {
                            throw UnsupportedHtmlFragmentMatch();
                        }
                        HyperlinkRelationship copy = mainPart.AddHyperlinkRelationship(authoredLink.Uri, authoredLink.IsExternal);
                        linkSplits[index] = copy;
                        changedLinkIds.Add(copy.Id);
                    }
                }

                // Keeping the original prefix attached preserves the offsets of earlier matches.
                for (int index = 0; index < fragments.Count; index++) {
                    int matchIndex = matches.Count - index - 1;
                    ReplaceHtmlFragmentMatch(matches[matchIndex], fragments[index].Anchor, linkSplits[matchIndex], changedLinkIds);
                }
            } catch {
                foreach (var fragment in fragments) {
                    if (fragment.Anchor.Parent == null) {
                        mainPart.DeletePart(fragment.Part);
                    }
                }
                throw;
            } finally {
                if (changedLinkIds.Count > 0) {
                    changedLinkIds.ExceptWith(RelationshipIds(document));
                    foreach (HyperlinkRelationship relationship in mainPart.HyperlinkRelationships.ToArray()) {
                        if (changedLinkIds.Contains(relationship.Id)) {
                            mainPart.DeleteReferenceRelationship(relationship);
                        }
                    }
                }
            }

            document.Save();
            return matches.Count;
        }

        /// <summary>
        /// Inserts an HTML fragment after the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph after which the fragment should be inserted.</param>
        /// <param name="htmlContent">HTML content to insert.</param>
        /// <param name="type">Optional format type of the fragment.</param>
        /// <returns>The created <see cref="WordEmbeddedDocument"/>.</returns>
        public WordEmbeddedDocument AddEmbeddedFragmentAfter(WordParagraph paragraph,
            string htmlContent, WordAlternativeFormatImportPartType type = WordAlternativeFormatImportPartType.Html) {
            ArgumentNullException.ThrowIfNull(paragraph);

            MainDocumentPart mainPart = _document.MainDocumentPart ?? throw new InvalidOperationException("The document does not contain a main document part.");
            var fragment = CreateEmbeddedFragment(mainPart, htmlContent, type);
            try {
                paragraph._paragraph.InsertAfterSelf(fragment.Anchor);
                return new WordEmbeddedDocument(this, fragment.Anchor);
            } catch {
                if (fragment.Anchor.Parent != null) {
                    fragment.Anchor.Remove();
                }
                mainPart.DeletePart(fragment.Part);
                throw;
            }
        }

        private static (AltChunk Anchor, AlternativeFormatImportPart Part) CreateEmbeddedFragment(MainDocumentPart mainPart,
            string htmlContent, WordAlternativeFormatImportPartType type) {
            PartTypeInfo partTypeInfo = type switch {
                WordAlternativeFormatImportPartType.Rtf => AlternativeFormatImportPartType.Rtf,
                WordAlternativeFormatImportPartType.Html => AlternativeFormatImportPartType.Html,
                WordAlternativeFormatImportPartType.TextPlain => AlternativeFormatImportPartType.TextPlain,
                _ => throw new InvalidOperationException("Unsupported format type")
            };

            htmlContent = EmbeddedFragmentContentNormalizer.Normalize(mainPart, htmlContent, type);
            AlternativeFormatImportPart part = mainPart.AddAlternativeFormatImportPart(partTypeInfo);
            try {
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent));
                part.FeedData(stream);
                return (new AltChunk { Id = mainPart.GetIdOfPart(part) }, part);
            } catch {
                mainPart.DeletePart(part);
                throw;
            }
        }

        private static void CollectHtmlFragmentMatches(List<WordParagraph> paragraphs, string searched, StringComparison comparison,
            List<HtmlFragmentMatch> matches) {
            if (paragraphs.Count == 0) {
                return;
            }

            string[] texts = paragraphs.Select(paragraph => paragraph.Text).ToArray();
            foreach (WordTextSegment segment in SearchText(texts, searched, comparison)) {
                WordParagraph first = paragraphs[segment.BeginIndex];
                Paragraph last = paragraphs[segment.EndIndex]._paragraph;
                var slices = new List<HtmlFragmentTextSlice>();
                for (int index = segment.BeginIndex; index <= segment.EndIndex; index++) {
                    int begin = index == segment.BeginIndex ? segment.BeginChar : 0;
                    int end = index == segment.EndIndex ? segment.EndChar + 1 : texts[index].Length;
                    if (begin < end) {
                        CollectHtmlFragmentSlices(paragraphs[index], begin, end, slices);
                    }
                }

                var followingParagraphs = new List<Paragraph>();
                Paragraph current = first._paragraph;
                while (true) {
                    if (current.ParagraphProperties?.Descendants().Any(element =>
                        element is NumberingChange
                        || (element.NamespaceUri == WordprocessingNamespace && TryGetRevisionType(element, out _))) == true) {
                        throw UnsupportedHtmlFragmentMatch();
                    }
                    if (ReferenceEquals(current, last)) {
                        break;
                    }
                    current = (Paragraph)current.NextSibling()!;
                    followingParagraphs.Add(current);
                }
                matches.Add(new HtmlFragmentMatch(first, followingParagraphs, slices));
            }
        }

        private static void CollectHtmlFragmentSlices(WordParagraph paragraph, int begin, int end,
            List<HtmlFragmentTextSlice> slices) {
            int offset = 0;
            if (paragraph._run is Run run && ReferenceEquals(run.Parent, paragraph._paragraph)) {
                CollectRun(run);
            } else if (paragraph._hyperlink is Hyperlink hyperlink && ReferenceEquals(hyperlink.Parent, paragraph._paragraph)
                && !hyperlink.Descendants<FieldChar>().Any()) {
                foreach (OpenXmlElement child in hyperlink.ChildElements) {
                    if (child is Run linkRun) {
                        CollectRun(linkRun);
                    } else {
                        CollectLeaf(child);
                    }
                }
            } else {
                throw UnsupportedHtmlFragmentMatch();
            }

            if (offset < end) {
                throw UnsupportedHtmlFragmentMatch();
            }

            void CollectRun(Run source) {
                // Read the run before reading its children because hidden comment runs contribute no text.
                if (WordParagraph.ReadVisibleText(source).Length == 0) {
                    return;
                }
                foreach (OpenXmlElement child in source.ChildElements) {
                    CollectLeaf(child);
                }
            }

            void CollectLeaf(OpenXmlElement child) {
                int nextOffset = offset + WordParagraph.ReadVisibleText(child).Length;
                if (nextOffset > offset && nextOffset > begin && offset < end) {
                    if (child.Parent is not Run owningRun || owningRun.RunProperties?.GetFirstChild<RunPropertiesChange>() != null
                        || child is not (Text or TabChar or PositionalTab or CarriageReturn or NoBreakHyphen or SoftHyphen)
                        && !(child is Break lineBreak && (lineBreak.Type == null || lineBreak.Type.Value == BreakValues.TextWrapping))) {
                        throw UnsupportedHtmlFragmentMatch();
                    }
                    slices.Add(new HtmlFragmentTextSlice(child, Math.Max(begin - offset, 0), Math.Min(end, nextOffset) - offset));
                }
                offset = nextOffset;
            }
        }

        private static void EnsureMatchesOutsideComplexFields(Body body, List<HtmlFragmentMatch> matches) {
            if (!body.Descendants<FieldChar>().Any()) {
                return;
            }

            var matchingRuns = new HashSet<Run>(matches.SelectMany(match => match.Slices).Select(slice => (Run)slice.Node.Parent!));
            int depth = 0;
            foreach (Run run in body.Descendants<Run>()) {
                if (run.Ancestors<TextBoxContent>().Any() || run.Ancestors<DeletedRun>().Any() || run.Ancestors<MoveFromRun>().Any()) {
                    continue;
                }

                bool isFieldRun = depth > 0;
                foreach (FieldChar marker in run.Elements<FieldChar>()) {
                    isFieldRun = true;
                    if (marker.FieldCharType?.Value == FieldCharValues.Begin) {
                        depth++;
                    } else if (marker.FieldCharType?.Value == FieldCharValues.End && depth > 0) {
                        depth--;
                    }
                }
                if (isFieldRun && matchingRuns.Contains(run)) {
                    throw UnsupportedHtmlFragmentMatch();
                }
            }
        }

        private static void EnsureHtmlFragmentsPreservePageBreaks(List<HtmlFragmentMatch> matches, Styles? styles) {
            foreach (HtmlFragmentMatch match in matches) {
                match.HasPageBreakBefore = HasEffectiveHtmlFragmentPageBreakBefore(match.FirstParagraph._paragraph, styles);
                if ((match.HasPageBreakBefore && !HasHtmlFragmentPrefix(match.Slices[0]))
                    || match.FollowingParagraphs.Any(paragraph => HasEffectiveHtmlFragmentPageBreakBefore(paragraph, styles))) {
                    throw UnsupportedHtmlFragmentMatch();
                }
            }
        }

        private static bool HasEffectiveHtmlFragmentPageBreakBefore(Paragraph paragraph, Styles? styles) {
            ParagraphProperties? properties = paragraph.ParagraphProperties;
            PageBreakBefore? pageBreak = properties?.PageBreakBefore;
            if (pageBreak == null) {
                foreach (Style style in WordRunFormatting.EnumerateStyleChain(styles, properties?.ParagraphStyleId?.Val?.Value, StyleValues.Paragraph)) {
                    pageBreak = style.StyleParagraphProperties?.PageBreakBefore;
                    if (pageBreak != null) {
                        break;
                    }
                }
            }
            pageBreak ??= styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.GetFirstChild<PageBreakBefore>();
            return pageBreak != null && pageBreak.Val?.Value != false;
        }

        // The original paragraph stays before the fragment when there is surviving content before its first matched node.
        private static bool HasHtmlFragmentPrefix(HtmlFragmentTextSlice first) {
            if (first.Begin > 0) {
                return true;
            }
            for (OpenXmlElement node = first.Node; node is not Paragraph; node = node.Parent!) {
                if (node.PreviousSibling() is OpenXmlElement preceding && preceding is not (ParagraphProperties or RunProperties)) {
                    return true;
                }
            }
            return false;
        }

        private static NotSupportedException UnsupportedHtmlFragmentMatch() =>
            new NotSupportedException("The text match intersects content that cannot be split for a block-level fragment.");

        private static void ReplaceHtmlFragmentMatch(HtmlFragmentMatch match, AltChunk anchor,
            HyperlinkRelationship? linkSplit, HashSet<string> changedLinkIds) {
            HtmlFragmentTextSlice first = match.Slices[0];
            Run firstRun = (Run)first.Node.Parent!;
            for (int index = 1; index < match.Slices.Count; index++) {
                HtmlFragmentTextSlice slice = match.Slices[index];
                Run run = (Run)slice.Node.Parent!;
                string text = WordParagraph.ReadVisibleText(slice.Node);
                if (slice.Begin == 0 && slice.End == text.Length) {
                    slice.Node.Remove();
                    OpenXmlElement? parent = run.Parent;
                    RemoveEmptyRun(run);
                    if (parent is Hyperlink link && !link.HasChildren) {
                        RemoveFragmentHyperlink(link, changedLinkIds);
                    }
                } else {
                    SetFragmentText((Text)slice.Node, text.Remove(slice.Begin, slice.End - slice.Begin));
                }
            }

            Paragraph paragraph = match.FirstParagraph._paragraph;
            OpenXmlElement? suffixStart = SplitFragmentStartRun(firstRun, first, linkSplit, changedLinkIds);
            if (ReferenceEquals(suffixStart, paragraph.ParagraphProperties?.NextSibling() ?? paragraph.FirstChild)
                || !HasParagraphContent(paragraph)) {
                paragraph.InsertBeforeSelf(anchor);
                RemoveEmptyParagraph(paragraph);
            } else {
                paragraph.InsertAfterSelf(anchor);
                if (suffixStart != null || paragraph.ParagraphProperties?.SectionProperties != null) {
                    Paragraph suffix = CreateFragmentContinuation(match.FirstParagraph, match.HasPageBreakBefore);
                    MoveFragmentSiblings(suffixStart, suffix);
                    anchor.InsertAfterSelf(suffix);
                }
            }

            foreach (Paragraph following in match.FollowingParagraphs) {
                RemoveEmptyParagraph(following);
            }
        }

        private static OpenXmlElement? SplitFragmentStartRun(Run run, HtmlFragmentTextSlice first,
            HyperlinkRelationship? linkSplit, HashSet<string> changedLinkIds) {
            OpenXmlElement? followingChild = first.Node.NextSibling();
            string text = WordParagraph.ReadVisibleText(first.Node);
            string suffix = text.Substring(first.End);
            if (suffix.Length > 0 || followingChild != null) {
                Run continuation = (Run)run.CloneNode(false);
                continuation.RunProperties = (RunProperties?)run.RunProperties?.CloneNode(true);
                if (suffix.Length > 0) {
                    var suffixText = (Text)first.Node.CloneNode(false);
                    SetFragmentText(suffixText, suffix);
                    continuation.Append(suffixText);
                }
                MoveFragmentSiblings(followingChild, continuation);
                run.InsertAfterSelf(continuation);
            }

            if (first.Begin > 0) {
                SetFragmentText((Text)first.Node, text.Substring(0, first.Begin));
            } else {
                first.Node.Remove();
            }

            OpenXmlElement? suffixStart = run.NextSibling();
            OpenXmlElement? parent = run.Parent;
            RemoveEmptyRun(run);
            if (parent is not Hyperlink hyperlink) {
                return suffixStart;
            }

            // Each surviving external link needs an independent relationship for later edits or removal.
            OpenXmlElement? afterLink = hyperlink.NextSibling();
            if (suffixStart != null) {
                var rightLink = (Hyperlink)hyperlink.CloneNode(false);
                MoveFragmentSiblings(suffixStart, rightLink);
                if (hyperlink.HasChildren && linkSplit != null) {
                    rightLink.Id = linkSplit.Id;
                }
                hyperlink.InsertAfterSelf(rightLink);
                afterLink = rightLink;
            }
            if (!hyperlink.HasChildren) {
                RemoveFragmentHyperlink(hyperlink, changedLinkIds);
            }
            return afterLink;
        }

        private static void RemoveFragmentHyperlink(Hyperlink hyperlink, HashSet<string> changedLinkIds) {
            if (hyperlink.Id?.Value is string relationshipId) {
                changedLinkIds.Add(relationshipId);
            }
            hyperlink.Remove();
        }

        private static Paragraph CreateFragmentContinuation(WordParagraph original, bool pageBreakBefore) {
            var result = new Paragraph();
            ParagraphProperties? properties = original._paragraph.ParagraphProperties;
            bool hasListMarker = WordListNumberingResolver.TryResolve(original, out var numbering);
            if (properties != null || hasListMarker || pageBreakBefore) {
                ParagraphProperties continuation = (ParagraphProperties?)properties?.CloneNode(true) ?? new ParagraphProperties();
                if (properties?.SectionProperties is SectionProperties section) {
                    section.Remove();
                    continuation.SectionProperties = section;
                }
                if (pageBreakBefore) {
                    continuation.PageBreakBefore = new PageBreakBefore { Val = false };
                }
                if (hasListMarker) {
                    continuation.NumberingProperties = new NumberingProperties(new NumberingId { Val = 0 });
                    var definitions = WordListNumberingResolver.GetCachedStyleCatalog(original._document).ListDefinitions;
                    if (definitions.TryGetValue(numbering.NumberId, out var definition) &&
                        definition.Levels.TryGetValue(numbering.Level, out var level) &&
                        level.SourceIndentation is Indentation listIndentation) {
                        continuation.Indentation = MergeFragmentIndentation(listIndentation, continuation.Indentation);
                    }
                }
                result.ParagraphProperties = continuation;
            }
            return result;
        }

        private static Indentation MergeFragmentIndentation(Indentation listIndentation, Indentation? directIndentation) {
            directIndentation ??= new Indentation();

            // Keep the list's text inset, but leave its marker's first-line position on the original paragraph.
            var directEdges = new HashSet<string>(directIndentation.GetAttributes().Select(attribute => Edge(attribute.LocalName)), StringComparer.Ordinal);
            foreach (OpenXmlAttribute attribute in listIndentation.GetAttributes()) {
                string edge = Edge(attribute.LocalName);
                if (edge != "firstLine" && !directEdges.Contains(edge)) {
                    directIndentation.SetAttribute(attribute);
                }
            }
            return directIndentation;

            static string Edge(string name) => name switch {
                "left" or "leftChars" or "start" or "startChars" => "start",
                "right" or "rightChars" or "end" or "endChars" => "end",
                "firstLine" or "firstLineChars" or "hanging" or "hangingChars" => "firstLine",
                _ => name
            };
        }

        private static void SetFragmentText(Text node, string value) {
            node.Text = value;
            node.Space = SpaceProcessingModeValues.Preserve;
        }

        private static void MoveFragmentSiblings(OpenXmlElement? node, OpenXmlCompositeElement destination) {
            while (node != null) {
                OpenXmlElement? next = node.NextSibling();
                node.Remove();
                destination.Append(node);
                node = next;
            }
        }

        private static void RemoveEmptyRun(Run run) {
            if (!run.ChildElements.Any(child => child is not RunProperties)) {
                run.Remove();
            }
        }

        private static bool HasParagraphContent(Paragraph paragraph) =>
            paragraph.ChildElements.Any(child => child is not ParagraphProperties);

        private static void RemoveEmptyParagraph(Paragraph paragraph) {
            if (!HasParagraphContent(paragraph) && paragraph.ParagraphProperties?.SectionProperties == null) {
                paragraph.Remove();
            }
        }

        private sealed class HtmlFragmentMatch {
            internal HtmlFragmentMatch(WordParagraph firstParagraph, List<Paragraph> followingParagraphs, List<HtmlFragmentTextSlice> slices) {
                FirstParagraph = firstParagraph;
                FollowingParagraphs = followingParagraphs;
                Slices = slices;
            }

            internal bool HasPageBreakBefore { get; set; }
            internal WordParagraph FirstParagraph { get; }
            internal List<Paragraph> FollowingParagraphs { get; }
            internal List<HtmlFragmentTextSlice> Slices { get; }
        }

        private readonly struct HtmlFragmentTextSlice {
            internal HtmlFragmentTextSlice(OpenXmlElement node, int begin, int end) {
                Node = node;
                Begin = begin;
                End = end;
            }

            internal OpenXmlElement Node { get; }
            internal int Begin { get; }
            internal int End { get; }
        }
    }
}
