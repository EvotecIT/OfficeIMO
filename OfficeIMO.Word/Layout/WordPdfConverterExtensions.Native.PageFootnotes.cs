using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private sealed record NativePageFootnoteSource(NativeNoteKey Key, WordSection Section,
        IReadOnlyList<WordParagraph> Paragraphs, PdfCore.PdfFootnoteDefinition Definition);

    private static void RegisterNativePageFootnoteFonts(WordDocument document, PdfCore.PdfOptions options,
        HashSet<string> families, HashSet<PdfCore.PdfStandardFont> slots, bool embedSystemFonts, NativeFontMap fonts) {
        var referenced = new HashSet<long>();
        foreach (WordSection section in document.Sections) {
            foreach (W.FootnoteReference reference in EnumerateNativeNoteParagraphs(section.Elements)
                .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph)).OfType<W.FootnoteReference>())
                if (reference.Id?.Value is long id && id != 0) referenced.Add(id);
            foreach (WordHeaderFooter story in EnumerateNativeActiveRunningStories(section))
                foreach (W.FootnoteReference reference in EnumerateNativeRunningFootnoteReferences(story))
                    if (reference.Id?.Value is long id && id != 0) referenced.Add(id);
        }
        if (referenced.Count == 0) return;
        W.Footnotes? source = document._wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes;
        IEnumerable<W.Footnote> ordinaryNotes = (source?.Elements<W.Footnote>() ?? Enumerable.Empty<W.Footnote>())
            .Where(note => referenced.Contains(note.Id?.Value ?? 0) && note.Type?.Value != W.FootnoteEndnoteValues.Separator &&
                note.Type?.Value != W.FootnoteEndnoteValues.ContinuationSeparator);
        foreach (W.Footnote note in ordinaryNotes.Concat(new[] {
            ResolveNativePageFootnoteSeparatorStory(source, continuation: false),
            ResolveNativePageFootnoteSeparatorStory(source, continuation: true)
        })) {
            foreach (W.Paragraph paragraph in note.Descendants<W.Paragraph>())
                RegisterNativeParagraphContentFonts(new WordParagraph(document, paragraph), NativeTableRunStyleDefaults.Empty,
                    options, families, slots, embedSystemFonts, fonts);
        }
    }

    private static W.Footnote ResolveNativePageFootnoteSeparatorStory(W.Footnotes? source, bool continuation) {
        W.FootnoteEndnoteValues type = continuation
            ? W.FootnoteEndnoteValues.ContinuationSeparator : W.FootnoteEndnoteValues.Separator;
        W.Footnote? stored = source?.Elements<W.Footnote>().FirstOrDefault(note => note.Type?.Value == type);
        if (stored != null) return stored;
        // Word supplies this story when the part omits it. Keep the projection
        // detached and let the normal paragraph/font owners resolve its line.
        OpenXmlElement mark = continuation ? new W.ContinuationSeparatorMark() : new W.SeparatorMark();
        return new W.Footnote(new W.Paragraph(
            new W.ParagraphProperties(new W.SpacingBetweenLines { Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
            new W.Run(mark))) { Type = type, Id = continuation ? 0 : -1 };
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
        NativeNoteNumbering notes, WordToPdfOptions? options) {
        var seen = new HashSet<OpenXmlElement>();
        foreach (WordHeaderFooter story in EnumerateNativeActiveRunningStories(section)) {
            foreach (W.FootnoteReference reference in EnumerateNativeRunningFootnoteReferences(story)) {
                if (!seen.Add(reference) || reference.Id?.Value is not long id || id == 0) continue;
                AddNativeFootnote(section._document, id, fallback, notes);
                W.Footnote? source = section._document._wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes?
                    .Elements<W.Footnote>().FirstOrDefault(note => note.Id?.Value == id);
                if (source != null) ReportNativeFootnoteFallbackContent(source, options);
                notes.ReportPlacementFallback();
            }
        }
    }

    private static bool TryPrepareNativePageFootnotes(INativePdfFlow flow, WordSection section,
        IReadOnlyList<WordElement> elements, NativeNoteNumbering notes, WordToPdfOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? bodyProfileGap, bool stableContinuationWidth) {
        List<W.FootnoteReference> references = EnumerateNativeNoteParagraphs(elements)
            .SelectMany(paragraph => EnumerateNativeVisibleNoteReferences(paragraph))
            .OfType<W.FootnoteReference>().Where(reference => reference.Id?.Value is long id && id != 0)
            .Distinct().ToList();
        if (references.Count == 0) return true;
        bool qualified = bodyProfileGap == null && stableContinuationWidth && (section.ColumnCount ?? 1) == 1;
        var sources = new List<(NativeNoteKey Key, W.Footnote Source)>();
        W.Footnotes? part = section._document._wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes;
        foreach (W.FootnoteReference reference in references) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (reference.Id?.Value is not long id || id == 0) continue;
            W.Footnote? source = part?.Elements<W.Footnote>().FirstOrDefault(note => note.Id?.Value == id);
            if (source == null) throw new InvalidDataException("The footnote reference has no matching note body.");
            if (!CanProjectNativePageNoteReference(reference) ||
                reference.Ancestors<W.SdtRun>().Any(control => !CanProjectNativePageNoteControl(control)) ||
                reference.Ancestors<W.Paragraph>().FirstOrDefault()?.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() == true ||
                !CanProjectNativePageNote(source, section, defaults, fonts)) {
                qualified = false;
            }
            NativeNoteKey key = GetNativeFootnoteKey(id);
            if (!sources.Any(item => item.Key == key)) sources.Add((key, source));
        }
        if (!qualified) {
            notes.ReportPlacementFallback(bodyProfileGap ?? (!stableContinuationWidth
                ? "A following section has a different usable page width; page-note continuation cannot yet reflow to that width."
                : null));
            foreach (var source in sources) ReportNativeFootnoteFallbackContent(source.Source, options);
            return false;
        }
        W.Footnote separator = ResolveNativePageFootnoteSeparatorStory(part, continuation: false);
        W.Footnote continuation = ResolveNativePageFootnoteSeparatorStory(part, continuation: true);
        if (!CanProjectNativePageNote(separator, section, defaults, fonts) ||
            !CanProjectNativePageNote(continuation, section, defaults, fonts)) {
            notes.ReportPlacementFallback();
            foreach (var source in sources) ReportNativeFootnoteFallbackContent(source.Source, options);
            ReportNativeFootnoteFallbackContent(separator, options);
            ReportNativeFootnoteFallbackContent(continuation, options);
            return false;
        }
        var noteOptions = new PdfCore.PdfFootnoteOptions {
            Placement = section.FootnoteSettings.Position == WordFootnotePosition.BeneathText
                ? PdfCore.PdfFootnotePlacement.BeneathText : PdfCore.PdfFootnotePlacement.BottomOfPage,
            Separator = content => RenderNativeNoteStory(content, separator, section,
                notes, options, defaults, fonts, label: null),
            ContinuationSeparator = content => RenderNativeNoteStory(content, continuation,
                section, notes, options, defaults, fonts, label: null)
        };
        foreach (var item in sources) {
            if (notes.TryGetPageFootnote(item.Key, out _)) continue;
            if (!notes.TryGetValue(item.Key, out int token)) continue;
            string label = notes.GetLabel(token);
            IReadOnlyList<WordParagraph> paragraphs = item.Source.Elements<W.Paragraph>()
                .Select(paragraph => new WordParagraph(section._document, paragraph)).ToList();
            PdfCore.PdfFootnoteDefinition definition = flow.DefineFootnote(content => {
                content.Bookmark(GetNativePageNoteDestination(item.Key));
                RenderNativeNoteStory(content, item.Source, section, notes, options, defaults, fonts, label);
            }, noteOptions);
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

    private static bool CanProjectNativePageNote(W.Footnote note, WordSection section,
        NativeDocumentDefaults defaults, NativeFontMap fonts) {
        // This profile is ordinary paragraph flow with inline images and links.
        // Native tables, floating objects, recursive notes and active fields retain
        // the explicitly reported section-flow fallback until their owner qualifies them.
        if (note.ChildElements.Any(child => child is not W.Paragraph)) return false;
        foreach (OpenXmlElement element in note.Descendants()) {
            if (element is W.FootnoteReference or W.EndnoteReference or W.TextBoxContent or W.Picture or
                W.EmbeddedObject or W.SimpleField or W.FieldChar or W.AltChunk or W.NumberingProperties or
                W.PageBreakBefore or W.FrameProperties or W.ParagraphBorders ||
                element is W.Shading && element.Parent is W.ParagraphProperties ||
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
        defaults = CreateNativeNoteStoryDefaults(defaults, section, label: null);
        foreach (W.Paragraph source in note.Elements<W.Paragraph>()) {
            // Match the style actually sent to the PDF owner, including the
            // paragraph's inherited style and explicit false overrides.
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(new WordParagraph(section._document, source), defaults, fonts);
            if (style.KeepWithNext || style.AnchoredCanvas != null) return false;
        }
        return true;
    }

    private static void ReportNativeFootnoteFallbackContent(W.Footnote source, WordToPdfOptions? options) {
        if (options != null && source.Descendants().Any(element => element is W.Drawing or W.Picture or
            W.EmbeddedObject or W.AltChunk or W.FootnoteReference or W.EndnoteReference or DocumentFormat.OpenXml.Math.OfficeMath))
            AddNativeExportWarning(options, "NativeFootnoteContentOmitted", "footnote",
                "Non-text footnote content is omitted by the unsupported-profile plain-text fallback.");
    }

    private static void RenderNativeNoteStory(PdfCore.PdfContentBuilder content, W.Footnote source,
        WordSection section, NativeNoteNumbering notes, WordToPdfOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, string? label) {
        defaults = CreateNativeNoteStoryDefaults(defaults, section, label);
        IReadOnlyList<W.Paragraph> paragraphs = source.Elements<W.Paragraph>().ToList();
        double previousAfter = 0D;
        for (int index = 0; index < paragraphs.Count; index++) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            var paragraph = new WordParagraph(section._document, paragraphs[index]);
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (runs.Where(run => !IsNativeHiddenTextRun(run, paragraph))
                .SelectMany(run => run.EnumerateEffectiveRunContent())
                .Any(child => child is W.SeparatorMark or W.ContinuationSeparatorMark))
                notes.ReportSeparatorMarkOmission();
            var images = new Dictionary<OpenXmlElement, PdfCore.PdfTextRun>();
            int imageCount = 0;
            foreach (WordParagraph run in runs) {
                if (IsNativeHiddenTextRun(run, paragraph)) continue;
                foreach (WordImage image in run.EnumerateImages()) {
                    if (++imageCount > (options?.MaxImagesPerParagraph ?? 1_000))
                        throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                    if (image._Image != null && TryCreateNativeCellInlineImage(image, out PdfCore.PdfTextRun? imageRun, options))
                        images[image._Image] = imageRun!;
                }
            }
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph, defaults, fonts);
            style.SpacingBefore = Math.Max(0D, style.SpacingBefore - previousAfter);
            previousAfter = style.SpacingAfter ?? 0D;
            if (ShouldSuppressNativeContextualSpacingBefore(paragraph)) style.SpacingBefore = 0D;
            if (ShouldSuppressNativeContextualSpacingAfter(paragraph,
                index + 1 < paragraphs.Count ? new WordParagraph(section._document, paragraphs[index + 1]) : null))
                style.SpacingAfter = 0D;
            List<PdfCore.PdfTextRun> projected = CreateNativeCellParagraphRuns(paragraph, notes,
                NativeTableStyleDefaults.Empty, defaults, fonts, images);
            content.Paragraph(builder => builder.Runs(projected), ResolveNativeParagraphAlign(paragraph, allowJustify: true), style: style);
        }
    }

    private static NativeDocumentDefaults CreateNativeNoteStoryDefaults(NativeDocumentDefaults defaults,
        WordSection section, string? label) => defaults with {
            NoteContentSection = section, NoteBodyLabel = label,
            BodyTextInkTopOffset = null, RunningContentContext = null, RunningContentSection = null
        };
}
