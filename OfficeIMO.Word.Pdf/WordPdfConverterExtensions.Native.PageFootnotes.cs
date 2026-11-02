using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private sealed record NativePageFootnoteSource(NativeNoteKey Key, WordSection Section,
        IReadOnlyList<WordParagraph> Paragraphs, PdfCore.PdfFootnoteDefinition Definition);

    private static void RegisterNativePageFootnoteFonts(WordDocument document, PdfCore.PdfCreateOptions options,
        HashSet<string> families, HashSet<PdfCore.PdfStandardFont> slots, bool embedSystemFonts, NativeFontMap fonts) {
        var referenced = new HashSet<NativeNoteKey>();
        foreach (WordSection section in document.Sections) {
            foreach (OpenXmlElement reference in EnumerateNativeNoteParagraphs(section.Elements)
                .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph)))
                if (GetNativeNoteReferenceKey(reference) is NativeNoteKey key) referenced.Add(key);
            foreach (WordHeaderFooter story in EnumerateNativeActiveRunningStories(section))
                foreach (W.FootnoteReference reference in EnumerateNativeRunningFootnoteReferences(story))
                    if (reference.Id?.Value is long id) referenced.Add(GetNativeFootnoteKey(id));
        }
        if (referenced.Count == 0) return;
        var main = document._wordprocessingDocument.MainDocumentPart;
        IEnumerable<W.FootnoteEndnoteType> ordinaryNotes = (main?.FootnotesPart?.Footnotes?
                .Elements<W.Footnote>() ?? Enumerable.Empty<W.Footnote>())
            .Where(note => note.Id?.Value is long id && referenced.Contains(GetNativeFootnoteKey(id)) && WordNoteResolver.IsOrdinary(note))
            .Cast<W.FootnoteEndnoteType>().Concat((main?.EndnotesPart?.Endnotes?
                .Elements<W.Endnote>() ?? Enumerable.Empty<W.Endnote>())
                .Where(note => note.Id?.Value is long id && referenced.Contains(GetNativeEndnoteKey(id)) && WordNoteResolver.IsOrdinary(note)));
        var selectedStories = new List<W.FootnoteEndnoteType>();
        if (referenced.Any(key => key.Kind == NativeNoteKind.Footnote)) {
            var stories = ResolveNativePageFootnoteStories(document);
            selectedStories.AddRange(new[] { stories.Separator, stories.ContinuationSeparator,
                stories.ContinuationNotice }.OfType<W.Footnote>());
        }
        if (referenced.Any(key => key.Kind == NativeNoteKind.Endnote)) {
            var stories = ResolveNativeEndnoteStories(document);
            selectedStories.AddRange(new[] { stories.Separator, stories.ContinuationSeparator,
                stories.ContinuationNotice }.OfType<W.Endnote>());
        }
        foreach (W.FootnoteEndnoteType note in ordinaryNotes.Concat(selectedStories).Distinct()) {
            IReadOnlyList<WordParagraph> paragraphs = note.Descendants<W.Paragraph>()
                .Select(paragraph => new WordParagraph(document, paragraph)).ToArray();
            using WordDocumentTraversal.ListMarkerRenderScope markerScope = WordDocumentTraversal
                .BuildNoteListMarkersForRendering(document, paragraphs);
            foreach (WordElement element in GetNativeNoteElements(document, note))
                RegisterNativeElementFonts(element, options, families, slots, embedSystemFonts, fonts);
        }
    }

    private static IEnumerable<W.FootnoteReference> EnumerateNativeRunningFootnoteReferences(WordHeaderFooter story) =>
        EnumerateNativeNoteParagraphs(CollapseNativeParagraphElements(story.Elements))
            .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph))
            .OfType<W.FootnoteReference>().Distinct();

    private static IEnumerable<OpenXmlElement> EnumerateNativeVisibleNoteReferences(WordParagraph paragraph, int textBoxDepth = 0) {
        // Note-free paragraphs need no recursive inventory. In particular, do
        // not impose note traversal limits on an otherwise supported cell box.
        if (paragraph._paragraph == null || !paragraph._paragraph.Descendants()
            .Any(element => element is W.FootnoteReference or W.EndnoteReference)) yield break;
        if (textBoxDepth > MaximumNativeHeaderFooterTextBoxNestingDepth)
            throw new InvalidDataException($"Text-box nesting exceeds the supported limit of {MaximumNativeHeaderFooterTextBoxNestingDepth} levels.");
        List<WordParagraph> runs = GetNativeRuns(paragraph, includeInlineContentControls: true);
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent())
                if (child is W.FootnoteReference or W.EndnoteReference) yield return child;
        }
        foreach (WordTextBox box in GetNativeVisibleTextBoxes(paragraph, runs))
            foreach (WordParagraph inner in GetNativeTextBoxParagraphs(box))
                foreach (OpenXmlElement reference in EnumerateNativeVisibleNoteReferences(inner, textBoxDepth + 1))
                    yield return reference;
    }

    private static void CollectNativeRunningFootnoteFallback(WordSection section, List<PdfFootnote> fallback,
        NativeNoteNumbering notes, WordPdfExportOptions? options) {
        var seen = new HashSet<OpenXmlElement>();
        foreach (WordHeaderFooter story in EnumerateNativeActiveRunningStories(section)) {
            foreach (W.FootnoteReference reference in EnumerateNativeRunningFootnoteReferences(story)) {
                if (!seen.Add(reference) || reference.Id?.Value is not long id) continue;
                W.Footnote? source = WordNoteResolver.FindFootnote(
                    section._document._wordprocessingDocument.MainDocumentPart, id);
                if (source == null) continue;
                AddNativeFootnote(section._document, id, fallback, notes);
                ReportNativeFootnoteFallbackContent(source, options);
                notes.ReportPlacementFallback();
            }
        }
    }

    private static bool TryPrepareNativePageFootnotes(INativePdfFlow flow, WordSection section,
        IReadOnlyList<WordElement> elements, NativeNoteNumbering notes, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? bodyProfileGap, bool stableContinuationWidth) {
        List<W.FootnoteReference> references = EnumerateNativeNoteParagraphs(elements)
            .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph))
            .OfType<W.FootnoteReference>().Where(reference => reference.Id?.Value is long)
            .Distinct().ToList();
        if (references.Count == 0) return true;
        bool qualified = bodyProfileGap == null && stableContinuationWidth && (section.ColumnCount ?? 1) == 1;
        var sources = new List<(NativeNoteKey Key, W.Footnote Source)>();
        W.Footnotes? part = section._document._wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes;
        foreach (W.FootnoteReference reference in references) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (reference.Id?.Value is not long id) continue;
            W.Footnote? source = WordNoteResolver.FindFootnote(
                section._document._wordprocessingDocument.MainDocumentPart, id);
            if (source == null) continue;
            if (!CanProjectNativePageNoteReference(reference) ||
                reference.Ancestors<W.SdtRun>().Any(control => !CanProjectNativePageNoteControl(control)) ||
                reference.Ancestors<W.Paragraph>().FirstOrDefault()?.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() == true ||
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
        W.Footnote separator = stories.Separator;
        W.Footnote continuation = stories.ContinuationSeparator;
        W.Footnote? notice = stories.ContinuationNotice;
        const int maximumComposedNoteParagraphs = 100000;
        int composedParagraphs = 0;
        foreach (W.FootnoteEndnoteType story in pendingSources.Select(item => (W.FootnoteEndnoteType)item.Source)
            .Concat(new W.FootnoteEndnoteType?[] { separator, continuation, notice }.OfType<W.FootnoteEndnoteType>())) {
            foreach (W.Paragraph _ in story.Descendants<W.Paragraph>()) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (++composedParagraphs > maximumComposedNoteParagraphs)
                    throw new InvalidDataException($"Footnote composition exceeds {maximumComposedNoteParagraphs} paragraphs.");
            }
        }
        if (!CanProjectNativePageNote(separator, section, defaults, fonts, options) ||
            !CanProjectNativePageNote(continuation, section, defaults, fonts, options) ||
            notice != null && !CanProjectNativePageNote(notice, section, defaults, fonts, options)) {
            notes.ReportPlacementFallback();
            foreach (var source in sources) ReportNativeFootnoteFallbackContent(source.Source, options);
            ReportNativeFootnoteFallbackContent(separator, options);
            ReportNativeFootnoteFallbackContent(continuation, options);
            if (notice != null) ReportNativeFootnoteFallbackContent(notice, options);
            return false;
        }
        var noteOptions = new PdfCore.PdfFootnoteOptions {
            Placement = section.FootnoteSettings.Position == WordFootnotePosition.BeneathText
                ? PdfCore.PdfFootnotePlacement.BeneathText : PdfCore.PdfFootnotePlacement.BottomOfPage,
            ReserveContinuationNoticeSpaceOnFinalPage = notice != null && !UsesModernNativeWordLayout(section._document),
            Separator = content => RenderNativeNoteStory(content, separator, section,
                notes, options, defaults, fonts, label: null),
            ContinuationSeparator = content => RenderNativeNoteStory(content, continuation,
                section, notes, options, defaults, fonts, label: null)
        };
        if (notice != null) noteOptions.ContinuationNotice = content => RenderNativeNoteStory(content,
            notice, section, notes, options, defaults, fonts, label: null);
        foreach (var item in pendingSources) {
            if (!notes.TryGetValue(item.Key, out int token)) continue;
            string label = notes.GetLabel(token);
            IReadOnlyList<WordParagraph> paragraphs = item.Source.Descendants<W.Paragraph>()
                .Select(paragraph => new WordParagraph(section._document, paragraph)).ToList();
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
                if (TryGetNativeStructuredBlockPropertyValue(structured.Document, structured.SdtBlock, out _)) continue;
                foreach (WordParagraph nested in EnumerateNativeNoteParagraphs(
                    GetNativeStructuredBlockElements(structured.Document, structured.SdtBlock), structuredDepth + 1, tableDepth))
                    yield return nested;
            } else if (element is WordCoverPage cover) {
                EnsureNativeStructuredDocumentTagDepth(structuredDepth);
                foreach (WordParagraph nested in EnumerateNativeNoteParagraphs(
                    GetNativeStructuredBlockElements(cover.Document, cover.SdtBlock), structuredDepth + 1, tableDepth))
                    yield return nested;
            }
        }
    }

    private static bool CanProjectNativePageNoteReference(W.FootnoteReference reference) {
        if (reference.Ancestors<W.TextBoxContent>().Any()) return false;
        if (reference.Ancestors<W.TableRow>().Any(row =>
            row.TableRowProperties?.GetFirstChild<W.TableHeader>() is { } header && WordTableRow.ReadOnOff(header))) return false;
        if (reference.Ancestors<W.Table>().Any(table =>
            table.GetFirstChild<W.TableProperties>()?.GetFirstChild<W.TablePositionProperties>() != null)) return false;
        if (reference.Ancestors<W.TableCell>().Any(cell => {
            string? direction = cell.GetFirstChild<W.TableCellProperties>()?.GetFirstChild<W.TextDirection>()?.Val?.InnerText;
            return direction != null && direction != "lrTb";
        })) return false;
        W.ParagraphProperties? properties = reference.Ancestors<W.Paragraph>().FirstOrDefault()?.ParagraphProperties;
        return properties?.GetFirstChild<W.FrameProperties>() == null &&
            properties?.GetFirstChild<W.ParagraphBorders>() == null && properties?.GetFirstChild<W.Shading>() == null;
    }

    private static bool CanProjectNativePageNoteControl(W.SdtRun control) =>
        IsNativeSimpleTextContentControl(control) &&
        (control.SdtProperties == null || !IsNativePropertyBoundStructuredBlock(control.SdtProperties));

    private static bool CanProjectNativePageNote(W.FootnoteEndnoteType note, WordSection section,
        NativeDocumentDefaults defaults, NativeFontMap fonts, WordPdfExportOptions? options = null, bool recordDiagnostics = true) {
        // Ordinary paragraph and qualified inline-table flow share their existing
        // owners. Floating objects, recursive notes and active fields retain the
        // explicitly reported section-flow fallback.
        defaults = CreateNativeNoteStoryDefaults(defaults, section, label: null);
        if (note.ChildElements.Any(child => child is not W.Paragraph and not W.Table) ||
            !WordNoteResolver.IsOrdinary(note) && note.Elements<W.Table>().Any() ||
            note.Elements<W.Table>().Any(table => !CanProjectNativeNoteTable(table, section, options, defaults, fonts))) return false;
        foreach (OpenXmlElement element in note.Descendants()) {
            if (element is W.FootnoteReference or W.EndnoteReference or W.TextBoxContent or W.Picture or
                W.EmbeddedObject or W.SimpleField or W.FieldChar or W.AltChunk or
                W.FrameProperties ||
                element is W.PageBreakBefore pageBreak && (pageBreak.Val == null || pageBreak.Val.Value) ||
                element is W.Break lineBreak && lineBreak.Type?.Value is { } type && type != W.BreakValues.TextWrapping ||
                element is DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor ||
                element is DocumentFormat.OpenXml.Math.OfficeMath ||
                element is W.SdtRun control && !CanProjectNativePageNoteControl(control)) return false;
            if (element is W.Drawing drawing) {
                var data = drawing.Inline?.Graphic?.GraphicData;
                if (drawing.Anchor != null || data?.Uri?.Value != "http://schemas.openxmlformats.org/drawingml/2006/picture" ||
                    data.Elements<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().Count() != 1) return false;
            }
        }
        foreach (W.Paragraph source in note.Descendants<W.Paragraph>()) {
            // Match the style actually sent to the PDF owner, including the
            // paragraph's inherited style and explicit false overrides.
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(new WordParagraph(section._document, source), defaults, fonts);
            if (style.AnchoredCanvas != null) return false;
        }
        // Definitions can be painted after the operation report is captured.
        // Record admitted note decoration now, after all profile checks pass.
        if (recordDiagnostics)
            foreach (W.Paragraph source in note.Descendants<W.Paragraph>())
                RecordNativeParagraphBarTabDiagnostics(new WordParagraph(section._document, source), options, "note paragraph");
        return true;
    }

    private static void ReportNativeFootnoteFallbackContent(W.FootnoteEndnoteType source, WordPdfExportOptions? options) {
        if (options != null && source.Descendants().Any(element => element is W.Drawing or W.Picture or
            W.EmbeddedObject or W.AltChunk or W.FootnoteReference or W.EndnoteReference or DocumentFormat.OpenXml.Math.OfficeMath))
            AddNativeExportWarning(options, "native-footnote-content-omitted", "footnote",
                "Non-text footnote content is omitted by the unsupported-profile plain-text fallback.");
    }

    private static void RenderNativeNoteStory(PdfCore.PdfContentBuilder content, W.FootnoteEndnoteType source,
        WordSection section, NativeNoteNumbering notes, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? label) {
        INativePdfFlow flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(content,
            GetNativePageSize(section, options)));
        RenderNativeNoteParagraphs(flow, source, section, notes, options, defaults, fonts, label);
    }

    private static NativeDocumentDefaults CreateNativeNoteStoryDefaults(NativeDocumentDefaults defaults,
        WordSection section, string? label) => defaults with {
            NoteContentSection = section, NoteBodyLabel = label,
            BodyTextInkTopOffset = null, RunningContentContext = null, RunningContentSection = null
        };
}
