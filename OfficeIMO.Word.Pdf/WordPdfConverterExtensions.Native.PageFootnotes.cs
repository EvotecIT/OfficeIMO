using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private sealed record NativePageFootnoteSource(NativeNoteKey Key, WordSection Section,
        IReadOnlyList<WordParagraph> Paragraphs, PdfCore.PdfFootnoteDefinition Definition);

    private static void RegisterNativePageFootnoteFonts(WordDocument document, PdfCore.PdfCreateOptions options,
        HashSet<string> families, HashSet<PdfCore.PdfStandardFont> slots, bool embedSystemFonts, NativeFontMap fonts) {
        var referenced = new HashSet<NativeNoteKey>();
        foreach (WordSection section in document.Sections) {
            foreach (OfficeXmlNode reference in EnumerateNativeNoteParagraphs(section.Elements)
                .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph)))
                if (GetNativeNoteReferenceKey(reference) is NativeNoteKey key) referenced.Add(key);
            foreach (WordHeaderFooter story in EnumerateNativeActiveRunningStories(section))
                foreach (OfficeXmlNode reference in EnumerateNativeRunningFootnoteReferences(story))
                    if (WordXml.ParseInt64(reference.WAttr("id")) is long id) referenced.Add(GetNativeFootnoteKey(id));
        }
        if (referenced.Count == 0) return;
        IEnumerable<OfficeXmlNode> ordinaryNotes = document.FootnotesNode().WElements("footnote")
            .Where(note => WordXml.ParseInt64(note.WAttr("id")) is long id && referenced.Contains(GetNativeFootnoteKey(id)) && WordConverterNodes.IsOrdinaryNote(note))
            .Concat(document.EndnotesNode().WElements("endnote")
                .Where(note => WordXml.ParseInt64(note.WAttr("id")) is long id && referenced.Contains(GetNativeEndnoteKey(id)) && WordConverterNodes.IsOrdinaryNote(note)));
        var selectedStories = new List<OfficeXmlNode>();
        if (referenced.Any(key => key.Kind == NativeNoteKind.Footnote)) {
            var stories = ResolveNativePageFootnoteStories(document);
            selectedStories.AddRange(new[] { stories.Separator, stories.ContinuationSeparator,
                stories.ContinuationNotice }.Where(story => story.Exists));
        }
        if (referenced.Any(key => key.Kind == NativeNoteKind.Endnote)) {
            var stories = ResolveNativeEndnoteStories(document);
            selectedStories.AddRange(new[] { stories.Separator, stories.ContinuationSeparator,
                stories.ContinuationNotice }.Where(story => story.Exists));
        }
        foreach (OfficeXmlNode note in ordinaryNotes.Concat(selectedStories).Distinct()) {
            IReadOnlyList<WordParagraph> paragraphs = note.Descendants(WordXml.W, "p")
                .Select(paragraph => document.ParagraphView(paragraph)).ToArray();
            using WordDocumentTraversal.ListMarkerRenderScope markerScope = WordDocumentTraversal
                .BuildNoteListMarkersForRendering(document, paragraphs);
            foreach (WordElement element in GetNativeNoteElements(document, note))
                RegisterNativeElementFonts(element, options, families, slots, embedSystemFonts, fonts);
        }
    }

    private static IEnumerable<OfficeXmlNode> EnumerateNativeRunningFootnoteReferences(WordHeaderFooter story) =>
        EnumerateNativeNoteParagraphs(CollapseNativeParagraphElements(story.Elements))
            .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph))
            .Where(reference => reference.IsW("footnoteReference")).Distinct();

    private static bool IsNativeNoteReference(OfficeXmlNode element) =>
        element.IsW("footnoteReference") || element.IsW("endnoteReference");

    private static IEnumerable<OfficeXmlNode> EnumerateNativeVisibleNoteReferences(WordParagraph paragraph, int textBoxDepth = 0) {
        // Note-free paragraphs need no recursive inventory. In particular, do
        // not impose note traversal limits on an otherwise supported cell box.
        OfficeXmlNode source = paragraph.Node();
        if (source.IsNone || !source.DescendantElements().Any(IsNativeNoteReference)) yield break;
        if (textBoxDepth > MaximumNativeHeaderFooterTextBoxNestingDepth)
            throw new InvalidDataException($"Text-box nesting exceeds the supported limit of {MaximumNativeHeaderFooterTextBoxNestingDepth} levels.");
        List<WordParagraph> runs = GetNativeRuns(paragraph, includeInlineContentControls: true);
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OfficeXmlNode child in run.EffectiveRunContentNodes())
                if (IsNativeNoteReference(child)) yield return child;
        }
        foreach (WordTextBox box in GetNativeVisibleTextBoxes(paragraph, runs))
            foreach (WordParagraph inner in GetNativeTextBoxParagraphs(box))
                foreach (OfficeXmlNode reference in EnumerateNativeVisibleNoteReferences(inner, textBoxDepth + 1))
                    yield return reference;
    }

    private static void CollectNativeRunningFootnoteFallback(WordSection section, List<PdfFootnote> fallback,
        NativeNoteNumbering notes, WordPdfExportOptions? options) {
        var seen = new HashSet<OfficeXmlNode>();
        foreach (WordHeaderFooter story in EnumerateNativeActiveRunningStories(section)) {
            foreach (OfficeXmlNode reference in EnumerateNativeRunningFootnoteReferences(story)) {
                if (!seen.Add(reference) || WordXml.ParseInt64(reference.WAttr("id")) is not long id) continue;
                OfficeXmlNode source = section._document.FindFootnoteNode(id);
                if (source.IsNone) continue;
                AddNativeFootnote(section._document, id, fallback, notes);
                ReportNativeFootnoteFallbackContent(source, options);
                notes.ReportPlacementFallback();
            }
        }
    }

    private static bool TryPrepareNativePageFootnotes(INativePdfFlow flow, WordSection section,
        IReadOnlyList<WordElement> elements, NativeNoteNumbering notes, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? bodyProfileGap, bool stableContinuationWidth) {
        List<OfficeXmlNode> references = EnumerateNativeNoteParagraphs(elements)
            .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph))
            .Where(reference => reference.IsW("footnoteReference") && WordXml.ParseInt64(reference.WAttr("id")) is long)
            .Distinct().ToList();
        if (references.Count == 0) return true;
        bool qualified = bodyProfileGap == null && stableContinuationWidth && (section.ColumnCount ?? 1) == 1;
        var sources = new List<(NativeNoteKey Key, OfficeXmlNode Source)>();
        foreach (OfficeXmlNode reference in references) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (WordXml.ParseInt64(reference.WAttr("id")) is not long id) continue;
            OfficeXmlNode source = section._document.FindFootnoteNode(id);
            if (source.IsNone) continue;
            if (!CanProjectNativePageNoteReference(reference) ||
                reference.Ancestors().Any(control => control.IsRunContentControl() && !CanProjectNativePageNoteControl(control)) ||
                reference.Parent.AncestorOrSelf(WordXml.W, "p").HasDescendant(WordXml.M, "oMath") ||
                !CanProjectNativePageNote(source, section, defaults, fonts, options)) {
                qualified = false;
            }
            NativeNoteKey key = GetNativeFootnoteKey(id);
            if (!sources.Any(item => item.Key == key)) sources.Add((key, source));
        }
        if (sources.Count == 0) return true;
        if (!qualified) {
            notes.ReportPlacementFallback(bodyProfileGap ?? (!stableContinuationWidth
                ? "A following section has a different usable page width; page-note continuation cannot yet reflow to that width."
                : null));
            foreach (var source in sources) ReportNativeFootnoteFallbackContent(source.Source, options);
            return false;
        }
        var pendingSources = sources.Where(item => !notes.TryGetPageFootnote(item.Key, out _) &&
            notes.TryGetValue(item.Key, out _)).ToList();
        if (pendingSources.Count == 0) return true;
        var stories = ResolveNativePageFootnoteStories(section._document);
        OfficeXmlNode separator = stories.Separator;
        OfficeXmlNode continuation = stories.ContinuationSeparator;
        OfficeXmlNode notice = stories.ContinuationNotice;
        const int maximumComposedNoteParagraphs = 100000;
        int composedParagraphs = 0;
        foreach (OfficeXmlNode story in pendingSources.Select(item => item.Source)
            .Concat(new[] { separator, continuation, notice }.Where(story => story.Exists))) {
            foreach (OfficeXmlNode _ in story.Descendants(WordXml.W, "p")) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (++composedParagraphs > maximumComposedNoteParagraphs)
                    throw new InvalidDataException($"Footnote composition exceeds {maximumComposedNoteParagraphs} paragraphs.");
            }
        }
        if (!CanProjectNativePageNote(separator, section, defaults, fonts, options) ||
            !CanProjectNativePageNote(continuation, section, defaults, fonts, options) ||
            notice.Exists && !CanProjectNativePageNote(notice, section, defaults, fonts, options)) {
            notes.ReportPlacementFallback();
            foreach (var source in sources) ReportNativeFootnoteFallbackContent(source.Source, options);
            ReportNativeFootnoteFallbackContent(separator, options);
            ReportNativeFootnoteFallbackContent(continuation, options);
            if (notice.Exists) ReportNativeFootnoteFallbackContent(notice, options);
            return false;
        }
        var noteOptions = new PdfCore.PdfFootnoteOptions {
            Placement = section.FootnoteSettings.Position == WordFootnotePosition.BeneathText
                ? PdfCore.PdfFootnotePlacement.BeneathText : PdfCore.PdfFootnotePlacement.BottomOfPage,
            ReserveContinuationNoticeSpaceOnFinalPage = notice.Exists && !UsesModernNativeWordLayout(section._document),
            Separator = content => RenderNativeNoteStory(content, separator, section,
                notes, options, defaults, fonts, label: null),
            ContinuationSeparator = content => RenderNativeNoteStory(content, continuation,
                section, notes, options, defaults, fonts, label: null)
        };
        if (notice.Exists) noteOptions.ContinuationNotice = content => RenderNativeNoteStory(content,
            notice, section, notes, options, defaults, fonts, label: null);
        foreach (var item in pendingSources) {
            if (!notes.TryGetValue(item.Key, out int token)) continue;
            string label = notes.GetLabel(token);
            IReadOnlyList<WordParagraph> paragraphs = item.Source.Descendants(WordXml.W, "p")
                .Select(paragraph => section._document.ParagraphView(paragraph)).ToList();
            PdfCore.PdfFootnoteDefinition definition = flow.DefineFootnote(content => {
                content.Bookmark(GetNativePageNoteDestination(item.Key));
                RenderNativeNoteStory(content, item.Source, section, notes, options, defaults, fonts, label);
            }, noteOptions);
            if (noteOptions.SharedDecorationOwner == null) noteOptions.ReuseDecorations(definition);
            notes.SetPageFootnote(new NativePageFootnoteSource(item.Key, section, paragraphs, definition));
        }
        return true;
    }

    private static IEnumerable<WordParagraph> EnumerateNativeNoteParagraphs(IReadOnlyList<WordElement> elements,
        int structuredDepth = 0, int tableDepth = 0) {
        foreach (WordElement element in elements) {
            if (element is WordParagraph paragraph) yield return paragraph;
            else if (element is WordTable table) {
                EnsureNativeTableDepth(tableDepth);
                foreach (WordTableRow row in table.Rows)
                    foreach (WordTableCell cell in row.Cells)
                        foreach (WordParagraph nested in EnumerateNativeNoteParagraphs(cell.Elements, structuredDepth, tableDepth + 1))
                            yield return nested;
            } else if (element is WordStructuredDocumentTag structured) {
                EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                if (TryGetNativeStructuredBlockPropertyValue(structured.Document, structured.SdtBlockNode(), out _)) continue;
                foreach (WordParagraph nested in EnumerateNativeNoteParagraphs(
                    GetNativeStructuredBlockElements(structured.Document, structured.SdtBlockNode()), structuredDepth + 1, tableDepth))
                    yield return nested;
            } else if (element is WordCoverPage cover) {
                EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                foreach (WordParagraph nested in EnumerateNativeNoteParagraphs(
                    GetNativeStructuredBlockElements(cover.Document, cover.SdtBlockNode()), structuredDepth + 1, tableDepth))
                    yield return nested;
            }
        }
    }

    private static bool CanProjectNativePageNoteReference(OfficeXmlNode reference) {
        if (reference.Parent.AncestorOrSelf(WordXml.W, "txbxContent").Exists) return false;
        if (reference.Ancestors().Any(row => row.IsW("tr") &&
            row.WChild("trPr").WChild("tblHeader") is { Exists: true } header && WordConverterNodes.ReadTableHeaderOnOff(header))) return false;
        if (reference.Ancestors().Any(table => table.IsW("tbl") &&
            table.WChild("tblPr").WChild("tblpPr").Exists)) return false;
        if (reference.Ancestors().Any(cell => {
            if (!cell.IsW("tc")) return false;
            string? direction = cell.WChild("tcPr").WChild("textDirection").WAttr("val");
            return direction != null && direction != "lrTb";
        })) return false;
        OfficeXmlNode properties = reference.Parent.AncestorOrSelf(WordXml.W, "p").WChild("pPr");
        return properties.WChild("framePr").IsNone &&
            properties.WChild("pBdr").IsNone && properties.WChild("shd").IsNone;
    }


    private static bool CanProjectNativePageNoteControl(OfficeXmlNode control) =>
        IsNativeSimpleTextContentControl(control) &&
        (control.WChild("sdtPr").IsNone || !IsNativePropertyBoundStructuredBlock(control.WChild("sdtPr")));

    private static bool CanProjectNativePageNote(OfficeXmlNode note, WordSection section,
        NativeDocumentDefaults defaults, NativeFontMap fonts, WordPdfExportOptions? options = null, bool recordDiagnostics = true) {
        // Ordinary paragraph and qualified inline-table flow share their existing
        // owners. Floating objects, recursive notes and active fields retain the
        // explicitly reported section-flow fallback.
        defaults = CreateNativeNoteStoryDefaults(defaults, section, label: null);
        if (note.Elements().Any(child => !child.IsW("p") && !child.IsW("tbl")) ||
            !WordConverterNodes.IsOrdinaryNote(note) && note.WElements("tbl").Any() ||
            note.WElements("tbl").Any(table => !CanProjectNativeNoteTable(table, section, options, defaults, fonts))) return false;
        foreach (OfficeXmlNode element in note.DescendantElements()) {
            if (element.NamespaceId == WordXml.W) {
                switch (element.LocalName) {
                    case "footnoteReference": case "endnoteReference": case "txbxContent": case "pict":
                    case "object": case "fldSimple": case "fldChar": case "altChunk": case "framePr":
                        return false;
                    case "pageBreakBefore" when element.OnOff() == true:
                        return false;
                    case "br" when element.WAttr("type") is { } type && type != "textWrapping":
                        return false;
                    case "sdt" when element.IsRunContentControl() && !CanProjectNativePageNoteControl(element):
                        return false;
                    case "drawing":
                        OfficeXmlNode data = element.Child(WordXml.Wp, "inline").Child(WordXml.A, "graphic").Child(WordXml.A, "graphicData");
                        if (element.Child(WordXml.Wp, "anchor").Exists ||
                            data.Attr("uri") != "http://schemas.openxmlformats.org/drawingml/2006/picture" ||
                            data.Elements(WordXml.Pic, "pic").Count() != 1) return false;
                        break;
                }
            } else if (element.Is(WordXml.Wp, "anchor") || element.Is(WordXml.M, "oMath")) return false;
        }
        foreach (OfficeXmlNode source in note.Descendants(WordXml.W, "p")) {
            // Match the style actually sent to the PDF owner, including the
            // paragraph's inherited style and explicit false overrides.
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(section._document.ParagraphView(source), defaults, fonts);
            if (style.AnchoredCanvas != null) return false;
        }
        // Definitions can be painted after the operation report is captured.
        // Record admitted note decoration now, after all profile checks pass.
        if (recordDiagnostics)
            foreach (OfficeXmlNode source in note.Descendants(WordXml.W, "p"))
                RecordNativeParagraphBarTabDiagnostics(section._document.ParagraphView(source), options, "note paragraph");
        return true;
    }

    private static void ReportNativeFootnoteFallbackContent(OfficeXmlNode source, WordPdfExportOptions? options) {
        if (options != null && source.DescendantElements().Any(element => element.IsW("drawing") || element.IsW("pict") ||
            element.IsW("object") || element.IsW("altChunk") || IsNativeNoteReference(element) || element.Is(WordXml.M, "oMath")))
            AddNativeExportWarning(options, "native-footnote-content-omitted", "footnote",
                "Non-text footnote content is omitted by the unsupported-profile plain-text fallback.");
    }

    private static void RenderNativeNoteStory(PdfCore.PdfContentBuilder content, OfficeXmlNode source,
        WordSection section, NativeNoteNumbering notes, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? label) {
        INativePdfFlow flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(content,
            GetNativePageSize(section, options)));
        RenderNativeNoteParagraphs(flow, source, section, notes, options, defaults, fonts, label);
    }

    private static NativeDocumentDefaults CreateNativeNoteStoryDefaults(NativeDocumentDefaults defaults,
        WordSection section, string? label) => defaults with {
            NoteContentSection = section, NoteBodyLabel = label,
            RunningContentContext = null, RunningContentSection = null
        };
}
