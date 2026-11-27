using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    // Text and a text box can share the same source run. Keep their original
    // order and formatting identities instead of treating the whole paragraph
    // as an object-only paragraph and losing its surrounding visible content.
    private static void RenderNativeMixedTextBoxes(INativePdfFlow pdf, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, (int Level, string Marker)? marker,
        Func<WordParagraph, (int Level, string Marker)?> getMarker, NativeNoteNumbering notes,
        WordPdfExportOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap, int depth,
        WordParagraph? nextParagraph = null) {
        CancellationToken cancellationToken = options?.CancellationToken ?? default;
        cancellationToken.ThrowIfCancellationRequested();
        RecordNativeMixedTextBoxDiagnostic(paragraph, runs, options,
            defaults.RunningContentContext == null ? "body paragraph" : "header/footer paragraph");
        int imageLimit = options?.MaxImagesPerParagraph ?? 1_000;
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(imageLimit, nameof(WordPdfExportOptions.MaxImagesPerParagraph));
        int imageCount = 0;
        var emittedNoteKeys = new HashSet<NativeNoteKey>();
        var pending = new List<WordParagraph>();
        var paragraphStyle = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
        if (paragraphStyle.SpacingBefore.ToPoints() is > 0D)
            pdf.ParagraphSpacingBefore(paragraphStyle.SpacingBefore.ToPoints());
        bool first = true;
        bool emittedOwnParagraph = false;
        foreach (WordParagraph run in GetNativeMixedTextBoxContentRuns(paragraph, runs, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OfficeXmlNode child in run.EffectiveRunContentNodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (child.IsW("rPr")) continue;
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
                        groupStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);
                        groupStyle.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
                        RenderNativeParagraphShapeGroups(pdf, paragraph, new[] { view },
                            ResolveNativeParagraphAlign(paragraph, allowJustify: false), options, groupStyle, ref imageCount, null,
                            runningStory: defaults.RunningContentContext != null);
                        if (groupStyle.AnchoredCanvas != null) {
                            pdf.Paragraph(builder => builder.AddLineBreak().SetFontSize(global::OfficeIMO.OfficeLength.Points(ResolveNativeParagraphFontSize(paragraph,
                                defaults, GetNativeParagraphStyleDefaults(paragraph)))), style: groupStyle);
                            emittedOwnParagraph = true;
                        }
                        continue;
                    }
                    if (view.Shape is WordShape shape) {
                        Flush();
                        if (!RenderNativeShape(pdf, shape, spacingAfter: 0D) && options != null)
                            AddNativeExportWarning(options, "native-mixed-drawable-unsupported", "mixed text-box paragraph",
                                "A sibling shape could not be mapped by the shared drawing renderer.");
                        continue;
                    }
                    if (view.Chart is WordChart chart) {
                        Flush();
                        if (PrepareNativeChart(chart, options, "mixed text-box paragraph chart") is { } drawing)
                            pdf.Drawing(drawing, ResolveNativeParagraphAlign(paragraph, allowJustify: false), spacingAfter: 0D,
                                style: NativeChartDrawingStyle());
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
        // Child boxes paint their own paragraphs. Classify this owner's rule
        // from the styled lines actually emitted, not recursively read box text.
        RecordNativeParagraphBarTabDiagnostics(paragraph, options, "mixed text-box paragraph",
            omitted: !emittedOwnParagraph);
        if (!ShouldSuppressNativeContextualSpacingAfter(paragraph, nextParagraph) && (paragraphStyle.SpacingAfter?.ToPoints()) is > 0D)
            pdf.ParagraphSpacingAfter(paragraphStyle.SpacingAfter.Value.ToPoints());

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
            style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
            WordParagraph[] segment = pending.ToArray();
            string text = string.Concat(segment.Select(run => run.Text));
            pdf.Paragraph(builder => AddNativeParagraphContent(builder, paragraph, currentMarker,
                segment, visible, text, references, notes, options, defaults, fontMap,
                inlineMarkerColumnWidth: currentMarker == null ? null : ResolveNativeInlineListMarkerColumnWidth(
                    paragraph, currentMarker.Value.Marker, style, defaults, fontMap), projectedContentOnly: true),
                ResolveNativeParagraphAlign(paragraph), style: style);
            emittedOwnParagraph = true;
            first = false;
            pending.Clear();
        }
    }

    // Running content needs its loss report before deferred page callbacks.
    // Inspect the same selected run children as the mixed flow owner: a child
    // box/object cannot supply its parent's text line. Numbering, references and
    // synthesized equations remain meaningful paragraph content without letters.
    private static bool HasNativeMixedTextBoxOwnContent(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs,
        WordSection section, WordPdfExportOptions? options) {
        if (WordDocumentTraversal.GetListInfo(paragraph)?.MarkerVisible == true) return true;
        foreach (WordParagraph run in GetNativeMixedTextBoxContentRuns(paragraph, runs, default)) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OfficeXmlNode child in run.EffectiveRunContentNodes()) {
                if (child.IsW("rPr")) continue;
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                if (view.GetTextBoxes().Any() || view.IsImage && view.Image != null) continue;
                if (HasNativeParagraphShapeGroups(new[] { view })) {
                    // Supported anchored groups emit an actual owner line. Inline,
                    // flowed or unsupported groups remain separate flow objects.
                    if (HasNativeParagraphShapeGroupAnchor(GetNativePageSize(section, options),
                        paragraph, new[] { view }, options)) return true;
                    continue;
                }
                if (view.Shape != null || view.Chart != null) continue;
                if (IsNativeRenderableTextRun(view, paragraph)) return true;
                if (child.IsW("footnoteReference") && paragraph._document.FindFootnoteNode(WordXml.ParseInt64(child.WAttr("id"))).Exists ||
                    child.IsW("endnoteReference") && paragraph._document.FindEndnoteNode(WordXml.ParseInt64(child.WAttr("id"))).Exists) return true;
            }
        }
        return false;
    }

    private static List<int> GetNativeMixedTextBoxNoteNumbers(IReadOnlyList<WordParagraph> runs,
        NativeNoteNumbering notes, HashSet<NativeNoteKey> emittedKeys) {
        var numbers = new List<int>();
        foreach (OfficeXmlNode child in runs.SelectMany(run => run.EffectiveRunContentNodes())) {
            NativeNoteKey? key = child.IsW("footnoteReference") && WordXml.ParseInt64(child.WAttr("id")) is long footnoteId
                ? GetNativeFootnoteKey(footnoteId)
                : child.IsW("endnoteReference") && WordXml.ParseInt64(child.WAttr("id")) is long endnoteId
                    ? GetNativeEndnoteKey(endnoteId) : null;
            if (key.HasValue && notes.TryGetValue(key.Value, out int number) && emittedKeys.Add(key.Value))
                numbers.Add(number);
        }
        return numbers;
    }

    private static List<WordParagraph> GetNativeMixedTextBoxContentRuns(WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, CancellationToken cancellationToken) {
        // Reuse the canonical equation occurrences and visible run projection.
        // Include each equation in the same source order as its sibling drawing,
        // rather than asking the whole-paragraph equation helper to replay it.
        cancellationToken.ThrowIfCancellationRequested();
        OfficeXmlNode paragraphNode = paragraph.Node();
        var equations = WordConverterNodes.EquationOccurrences(paragraph._document, paragraphNode);
        if (equations.Count == 0) return runs.ToList();
        var positions = new Dictionary<OfficeXmlNode, int>();
        foreach (OfficeXmlNode element in paragraphNode.DescendantElements()) {
            cancellationToken.ThrowIfCancellationRequested();
            positions.Add(element, positions.Count);
        }
        var ordered = new List<(int Position, WordParagraph Run)>();
        foreach (WordParagraph run in runs) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (OfficeXmlNode child in run.EffectiveRunContentNodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (child.IsW("rPr")) continue;
                if (child.Ancestors().Prepend(child).Any(element =>
                    equations.Any(equation => equation.Equation.IsBackingNode(element)))) continue;
                int position = positions.TryGetValue(child, out int childPosition) ? childPosition
                    : !run.RunNode().IsNone && positions.TryGetValue(run.RunNode(), out int runPosition) ? runPosition : int.MaxValue;
                ordered.Add((position, CreateNativeRunContentView(run, new[] { child })));
            }
        }
        var segments = WordConverterNodes.VisibleEquationContentSegments(paragraphNode, equations,
            element => !element.IsW("r") ||
                !IsNativeHiddenTextRun(WordConverterNodes.RunView(paragraph._document, paragraphNode, element), paragraph));
        foreach (var equation in equations) {
            cancellationToken.ThrowIfCancellationRequested();
            var segment = segments.FirstOrDefault(item => ReferenceEquals(item.Equation, equation.Equation));
            if (segment == null || string.IsNullOrEmpty(equation.Equation.Text)) continue;
            OfficeXmlNode start = paragraphNode.DescendantElements().First(element => {
                cancellationToken.ThrowIfCancellationRequested();
                return equation.Equation.IsBackingNode(element);
            });
            WordParagraph source = segment.CreateSourceParagraph(paragraph._document, paragraphNode, paragraph);
            if (source.RunNode().IsNone) {
                // A wrapper-backed equation has no Wordprocessing run. Give
                // its selected visible text a real run, retaining the wrapper's
                // applicable formatting and hyperlink context without changing XML.
                OfficeXmlNode formattingRun = source.SdtRunNode().FirstDescendantElement(WordXml.W, "r");
                if (formattingRun.IsNone) formattingRun = source.HyperlinkNode().FirstDescendantElement(WordXml.W, "r");
                if (formattingRun.IsNone) formattingRun = paragraphNode.CreateElement(WordXml.W, "r");
                source = WordConverterNodes.RunView(paragraph._document, paragraphNode, formattingRun, source.HyperlinkNode());
            }
            OfficeXmlNode text = source.RunNode().CreateElement(WordXml.W, "t");
            text.SetInnerText(equation.Equation.Text);
            text.SetAttribute(WordXml.Xml, "space", "preserve");
            ordered.Add((positions[start], CreateNativeRunContentView(source, new[] { text })));
        }
        return ordered.OrderBy(item => item.Position).Select(item => item.Run).ToList();
    }

    private static void RecordNativeMixedTextBoxDiagnostic(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs,
        WordPdfExportOptions? options, string source) {
        if (options != null && GetNativeVisibleTextBoxes(paragraph, runs).Count > 0)
            AddNativeExportWarning(options, "native-mixed-text-box-layout-approximated", source,
                "Text-box content and surrounding text are retained in source order; anchored positioning, shape geometry, wrapping and pagination are approximated by bounded flow.");
    }
}
