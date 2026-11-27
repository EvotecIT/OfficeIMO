
namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Returns paragraph-level aliases without admitting bookmarks from nested stories.</summary>
    private static IEnumerable<string> GetNativeParagraphBookmarkNames(WordParagraph paragraph) =>
        paragraph.Node().WElements("bookmarkStart").Select(bookmark => bookmark.WAttr("name"))
            .OfType<string>().Where(name => !string.IsNullOrWhiteSpace(name)).Distinct(StringComparer.Ordinal);

    private static void RenderNativeParagraphBookmarks(INativePdfFlow flow, WordParagraph paragraph) {
        foreach (string name in GetNativeParagraphBookmarkNames(paragraph)) flow.Bookmark(name);
    }

    // Source existence and actual paragraph projection are distinct. A cached
    // Word link can target a table or text-box bookmark whose text is retained
    // without a corresponding PDF destination. Keep that text and its styling;
    // do not invent a destination or pass a broken link to the strict PDF owner.
    private sealed class NativeBookmarkTargets {
        private readonly HashSet<string> sourceNames;
        private readonly HashSet<string> projectedNames;
        private readonly HashSet<string> reported = new(StringComparer.Ordinal);
        private readonly WordPdfExportOptions? options;

        internal NativeBookmarkTargets(HashSet<string> sourceNames, HashSet<string> projectedNames,
            WordPdfExportOptions? options) {
            this.sourceNames = sourceNames;
            this.projectedNames = projectedNames;
            this.options = options;
        }

        internal string? Resolve(string? name) => string.IsNullOrWhiteSpace(name) || !projectedNames.Contains(name!)
            ? null : name;

        internal void Report(WordHyperLink link) {
            if (link.Uri?.IsAbsoluteUri == true || string.IsNullOrWhiteSpace(link.Anchor) ||
                projectedNames.Contains(link.Anchor!) || options == null || !reported.Add(link.Anchor!)) return;
            bool exists = sourceNames.Contains(link.Anchor!);
            AddNativeExportWarning(options,
                exists ? "native-bookmark-target-unprojected" : "native-bookmark-target-missing", "hyperlink",
                exists
                    ? "Internal link target '" + link.Anchor + "' exists in the Word document but its source story or bookmark position is not projected as a PDF destination. The visible link text is retained without navigation."
                    : "Internal link target '" + link.Anchor + "' was not found in the source stories inspected for conversion. The visible link text is retained without navigation.");
        }
    }

    private static NativeBookmarkTargets? CreateNativeBookmarkTargets(WordDocument document,
        WordPdfExportOptions? options, NativeDocumentDefaults defaults, NativeFontMap fonts,
        NativeNoteNumbering notes) {
        var roots = new List<OfficeXmlNode>();
        if (document.DocumentNode().Exists) roots.Add(document.DocumentNode());
        roots.AddRange(document.HeaderFooterRootNodes());
        if (document.FootnotesNode().Exists) roots.Add(document.FootnotesNode());
        if (document.EndnotesNode().Exists) roots.Add(document.EndnotesNode());
        if (!roots.Any(root => root.ModelDescendants("hyperlink").Any(link => !string.IsNullOrWhiteSpace(link.WAttr("anchor")))))
            return null;
        var known = new HashSet<string>(roots.SelectMany(root => root.Descendants(WordXml.W, "bookmarkStart"))
            .Select(bookmark => bookmark.WAttr("name")).OfType<string>().Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.Ordinal);
        var projected = new HashSet<string>(StringComparer.Ordinal);
        var paragraphs = new List<WordParagraph>();
        var standaloneLinks = new List<WordHyperLink>();
        var referencedNotes = new HashSet<(OfficeXmlNode Source, WordSection Section)>();
        IReadOnlyList<WordSection> sections = document.Sections;
        foreach (WordSection section in sections) {
            IReadOnlyList<WordElement> elements = CollapseNativeParagraphElements(section.Elements);
            foreach (OfficeXmlNode bookmark in GetNativeEmptySectionMarkSource(section, elements).WElements("bookmarkStart"))
                if (bookmark.WAttr("name") is { Length: > 0 } name) projected.Add(name);
            Collect(elements, admitBookmarks: true);
            foreach (WordHeaderFooter story in EnumerateNativeActiveRunningStories(section))
                Collect(story.Elements, admitBookmarks: false);

            // Reuse the converter's selected, visible reference owner. Stored
            // unused notes cannot supply a destination to the rendered document.
            IEnumerable<WordParagraph> sourceParagraphs = EnumerateNativeNoteParagraphs(section.Elements)
                .Concat(EnumerateNativeActiveRunningStories(section).SelectMany(story =>
                    EnumerateNativeNoteParagraphs(story.Elements)));
            foreach (OfficeXmlNode reference in sourceParagraphs.SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph))) {
                OfficeXmlNode source = reference.IsW("footnoteReference") ? document.FindFootnoteNode(WordXml.ParseInt64(reference.WAttr("id"))) :
                    reference.IsW("endnoteReference") ? document.FindEndnoteNode(WordXml.ParseInt64(reference.WAttr("id"))) : default;
                WordSection targetSection = source.IsW("endnote") && notes.EndnotesAtDocumentEnd ? sections[sections.Count - 1] : section;
                if (source.Exists && (source.HasDescendant(WordXml.W, "bookmarkStart") ||
                    source.ModelDescendants("hyperlink").Any(link => !string.IsNullOrWhiteSpace(link.WAttr("anchor")))) &&
                    referencedNotes.Add((source, targetSection)) &&
                    CanProjectNativePageNote(source, targetSection, defaults, fonts, options, recordDiagnostics: false))
                    Collect(GetNativeNoteElements(document, source), admitBookmarks: true);
            }
        }
        var result = new NativeBookmarkTargets(known, projected, options);
        // Reports are captured before deferred running-story callbacks. Inspect
        // the same visible run wrappers now and deduplicate repeated link runs.
        foreach (WordParagraph paragraph in paragraphs) {
            foreach (WordParagraph run in GetNativeRuns(paragraph, includeInlineContentControls: true))
                if (!IsNativeHiddenTextRun(run, paragraph) && run.IsHyperLink && run.Hyperlink is { } link &&
                    !string.IsNullOrEmpty(link.Text)) result.Report(link);
        }
        foreach (WordHyperLink link in standaloneLinks)
            if (!string.IsNullOrEmpty(link.Text)) result.Report(link);
        return result;

        void Collect(IEnumerable<WordElement> source, bool admitBookmarks, int structuredDepth = 0, int tableDepth = 0, int textBoxDepth = 0) {
            foreach (WordElement element in CollapseNativeParagraphElements(source)) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (element is WordParagraph paragraph) {
                    paragraphs.Add(paragraph);
                    if (admitBookmarks) AddProjected(paragraph);
                    foreach (WordTextBox box in GetNativeVisibleTextBoxes(paragraph, GetNativeRuns(paragraph))) {
                        if (textBoxDepth >= MaximumNativeHeaderFooterTextBoxNestingDepth)
                            throw new InvalidDataException($"Text-box nesting exceeds the supported limit of {MaximumNativeHeaderFooterTextBoxNestingDepth} levels.");
                        Collect(GetNativeTextBoxParagraphs(box), admitBookmarks: false, structuredDepth, tableDepth, textBoxDepth + 1);
                    }
                } else if (element is WordHyperLink link) {
                    standaloneLinks.Add(link);
                } else if (element is WordTableOfContent stored && stored.ShowsStoredEntries) {
                    EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                    Collect(GetNativeStructuredBlockElements(document, stored.SdtBlockNode()), admitBookmarks, structuredDepth + 1, tableDepth, textBoxDepth);
                } else if (element is WordTableOfContent toc && toc.TitleParagraph is { } title) {
                    Collect(new[] { title }, admitBookmarks, structuredDepth, tableDepth, textBoxDepth);
                } else if (element is WordTable table) {
                    EnsureNativeTableDepth(tableDepth);
                    foreach (WordTableRow row in table.Rows)
                        foreach (WordTableCell cell in row.Cells)
                            if (!IsNativeHorizontalMergeContinuation(cell) && !IsNativeVerticalMergeContinuation(cell))
                                Collect(EnumerateNativeTableCellElements(cell), admitBookmarks: false, structuredDepth, tableDepth + 1, textBoxDepth);
                } else if (element is WordStructuredDocumentTag structured) {
                    EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                    if (!TryGetNativeStructuredBlockPropertyValue(document, structured.SdtBlockNode(), out _))
                        Collect(GetNativeStructuredBlockElements(document, structured.SdtBlockNode()), admitBookmarks, structuredDepth + 1, tableDepth, textBoxDepth);
                } else if (element is WordCoverPage cover) {
                    EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                    Collect(GetNativeStructuredBlockElements(document, cover.SdtBlockNode()), admitBookmarks, structuredDepth + 1, tableDepth, textBoxDepth);

                }
            }
        }

        void AddProjected(WordParagraph paragraph) {
            // Flow-break fragments own only the aliases on their side of the break.
            if (TrySplitNativeParagraphAtVisibleFlowBreak(paragraph, out _, out WordParagraph? before, out WordParagraph? after)) {
                if (before != null) AddProjected(before);
                if (after != null) AddProjected(after);
            } else projected.UnionWith(GetNativeParagraphBookmarkNames(paragraph));
        }
    }
}
