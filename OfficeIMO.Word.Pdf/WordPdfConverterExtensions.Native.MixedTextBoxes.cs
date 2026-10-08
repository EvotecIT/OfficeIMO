using DocumentFormat.OpenXml;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Text and a text box can share the same source run. Keep their original
    // order and formatting identities instead of treating the whole paragraph
    // as an object-only paragraph and losing its surrounding visible content.
    private static void RenderNativeMixedTextBoxes(INativePdfFlow pdf, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, (int Level, string Marker)? marker,
        Func<WordParagraph, (int Level, string Marker)?> getMarker, NativeNoteNumbering notes,
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap, int depth,
        WordParagraph? nextParagraph = null) {
        RecordNativeMixedTextBoxDiagnostic(runs, options,
            defaults.RunningContentContext == null ? "body paragraph" : "header/footer paragraph");
        int imageLimit = options?.MaxImagesPerParagraph ?? 1_000;
        if (imageLimit <= 0) throw new ArgumentOutOfRangeException(nameof(WordToPdfOptions.MaxImagesPerParagraph));
        int imageCount = 0;
        var emittedNoteKeys = new HashSet<long>();
        var pending = new List<WordParagraph>();
        var paragraphStyle = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
        if (paragraphStyle.SpacingBefore is > 0D)
            pdf.ParagraphSpacingBefore(paragraphStyle.SpacingBefore);
        bool first = true;
        foreach (WordParagraph run in GetNativeMixedTextBoxContentRuns(paragraph, runs)) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
                if (child is W.RunProperties) continue;
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                List<WordTextBox> boxes = view.GetTextBoxes().ToList();
                if (boxes.Count == 0) {
                    if (view.IsImage && view.Image != null) {
                        Flush();
                        if (++imageCount > imageLimit)
                            throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                        RenderNativeImage(pdf, view.Image, options: options, source: "mixed text-box paragraph image");
                        continue;
                    }
                    if (HasNativeParagraphShapeGroups(new[] { view })) {
                        Flush();
                        var groupStyle = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
                        groupStyle.SpacingBefore = 0D;
                        groupStyle.SpacingAfter = 0D;
                        RenderNativeParagraphShapeGroups(pdf, paragraph, new[] { view },
                            ResolveNativeParagraphAlign(paragraph, allowJustify: false), options, groupStyle, ref imageCount, null);
                        if (groupStyle.AnchoredCanvas != null)
                            pdf.Paragraph(builder => builder.FontSize(ResolveNativeParagraphFontSize(paragraph,
                                defaults, GetNativeParagraphStyleDefaults(paragraph))).LineBreak(), style: groupStyle);
                        continue;
                    }
                    if (view.Shape is WordShape shape) {
                        Flush();
                        if (!RenderNativeShape(pdf, shape, spacingAfter: 0D) && options != null)
                            AddNativeExportWarning(options, "NativeMixedDrawableUnsupported", "mixed text-box paragraph",
                                "A sibling shape could not be mapped by the shared drawing renderer.");
                        continue;
                    }
                    if (view.Chart is WordChart chart) {
                        Flush();
                        if (PrepareNativeChart(chart, options, "mixed text-box paragraph chart") is { } drawing)
                            pdf.Drawing(drawing, ResolveNativeParagraphAlign(paragraph, allowJustify: false), spacingAfter: 0D);
                        continue;
                    }
                    pending.Add(view);
                    continue;
                }
                Flush();
                foreach (WordTextBox box in boxes)
                    RenderNativeTextBox(pdf, box, getMarker, notes, options, defaults, fontMap,
                        GetNativeTextBoxPlainText(paragraph, box), depth + 1);
            }
        }
        Flush();
        if (!ShouldSuppressNativeContextualSpacingAfter(paragraph, nextParagraph) && paragraphStyle.SpacingAfter is > 0D)
            pdf.ParagraphSpacingAfter(paragraphStyle.SpacingAfter.Value);

        void Flush() {
            bool visible = pending.Any(run => IsNativeRenderableTextRun(run, paragraph));
            List<int> references = GetNativeMixedTextBoxNoteNumbers(pending, notes, emittedNoteKeys);
            (int Level, string Marker)? currentMarker = first ? marker : null;
            if (!visible && references.Count == 0 && currentMarker is not { Marker.Length: > 0 }) {
                pending.Clear();
                return;
            }
            var style = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
            if (currentMarker != null) {
                ApplyNativeInlineListIndent(paragraph, style);
                ApplyNativeInlineListMarkerAlignment(paragraph, currentMarker.Value.Marker, style, defaults, fontMap);
            }
            style.SpacingBefore = 0D;
            style.SpacingAfter = 0D;
            WordParagraph[] segment = pending.ToArray();
            string text = string.Concat(segment.Select(run => run.Text));
            pdf.Paragraph(builder => AddNativeParagraphContent(builder, paragraph, currentMarker,
                segment, visible, text, references, notes, options, defaults, fontMap,
                inlineMarkerColumnWidth: currentMarker == null ? null : ResolveNativeInlineListMarkerColumnWidth(
                    paragraph, currentMarker.Value.Marker, style, defaults, fontMap), projectedContentOnly: true),
                ResolveNativeParagraphAlign(paragraph, allowJustify: false), style: style);
            first = false;
            pending.Clear();
        }
    }

    private static List<int> GetNativeMixedTextBoxNoteNumbers(IReadOnlyList<WordParagraph> runs,
        NativeNoteNumbering notes, HashSet<long> emittedKeys) {
        var numbers = new List<int>();
        foreach (OpenXmlElement child in runs.SelectMany(run => run.EnumerateEffectiveRunContent())) {
            long? key = child is W.FootnoteReference footnote && footnote.Id?.Value is long footnoteId
                ? GetNativeFootnoteKey(footnoteId)
                : child is W.EndnoteReference endnote && endnote.Id?.Value is long endnoteId
                    ? GetNativeEndnoteKey(endnoteId) : null;
            if (key.HasValue && notes.TryGetValue(key.Value, out int number) && emittedKeys.Add(key.Value))
                numbers.Add(number);
        }
        return numbers;
    }

    private static List<WordParagraph> GetNativeMixedTextBoxContentRuns(WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs) {
        // Reuse the canonical equation occurrences and visible run projection.
        // Include each equation in the same source order as its sibling drawing,
        // rather than asking the whole-paragraph equation helper to replay it.
        var equations = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph);
        if (equations.Count == 0) return runs.ToList();
        var positions = paragraph._paragraph.Descendants().Select((element, index) => (element, index))
            .ToDictionary(item => item.element, item => item.index);
        var ordered = new List<(int Position, WordParagraph Run)>();
        foreach (WordParagraph run in runs) {
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
                if (child is W.RunProperties) continue;
                if (child.Ancestors().Prepend(child).Any(element =>
                    equations.Any(equation => equation.Equation.IsBackingElement(element)))) continue;
                int position = positions.TryGetValue(child, out int childPosition) ? childPosition
                    : run._run != null && positions.TryGetValue(run._run, out int runPosition) ? runPosition : int.MaxValue;
                ordered.Add((position, CreateNativeRunContentView(run, new[] { child })));
            }
        }
        var segments = WordEquation.GetVisibleContentSegments(paragraph._paragraph, equations,
            element => element is not W.Run source ||
                !IsNativeHiddenTextRun(new WordParagraph(paragraph._document, paragraph._paragraph, source), paragraph));
        foreach (var equation in equations) {
            var segment = segments.FirstOrDefault(item => ReferenceEquals(item.Equation, equation.Equation));
            if (segment == null || string.IsNullOrEmpty(equation.Equation.Text)) continue;
            OpenXmlElement start = positions.OrderBy(item => item.Value)
                .First(item => equation.Equation.IsBackingElement(item.Key)).Key;
            WordParagraph source = segment.CreateSourceParagraph(paragraph._document, paragraph._paragraph, paragraph);
            if (source._run == null) {
                // A wrapper-backed equation has no Wordprocessing run. Give
                // its selected visible text a real run, retaining the wrapper's
                // applicable formatting and hyperlink context without changing XML.
                W.Run formattingRun = source._stdRun?.Descendants<W.Run>().FirstOrDefault()
                    ?? source._hyperlink?.Descendants<W.Run>().FirstOrDefault() ?? new W.Run();
                source = new WordParagraph(paragraph._document, paragraph._paragraph, formattingRun) {
                    _hyperlink = source._hyperlink
                };
            }
            ordered.Add((positions[start], CreateNativeRunContentView(source, new OpenXmlElement[] {
                new W.Text(equation.Equation.Text) { Space = SpaceProcessingModeValues.Preserve }
            })));
        }
        return ordered.OrderBy(item => item.Position).Select(item => item.Run).ToList();
    }

    private static void RecordNativeMixedTextBoxDiagnostic(IReadOnlyList<WordParagraph> runs,
        WordToPdfOptions? options, string source) {
        if (options != null && runs.Any(run => run.GetTextBoxes().Any()) &&
            runs.Any(run => run.EnumerateEffectiveRunContent().OfType<W.Text>().Any()))
            AddNativeExportWarning(options, "NativeMixedTextBoxLayoutApproximated", source,
                "Text surrounding peer or nested text boxes is retained in source order; anchored box positioning, wrapping and pagination across boxes are approximated by bounded flow.");
    }
}
