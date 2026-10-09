using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static void ReportUnselectedFrameText(OdgShape shape, OdfConversionReport report, string feature) {
        if (shape.ElementName != "frame") return;
        XElement selectedStory = shape.TextRoot;
        bool omittedStory = !ReferenceEquals(selectedStory, shape.Element) &&
            shape.Element.Elements().Any(IsFrameTextContainer);
        omittedStory |= shape.Element.Elements().Any(story =>
            (story.Name == OdfNamespaces.Draw + "image" || story.Name == OdfNamespaces.Draw + "text-box") &&
            !ReferenceEquals(selectedStory, story) && story.Elements().Any(IsFrameTextContainer));
        if (omittedStory)
            report.Add(feature + (shape.IsImage ? ":image-text" : ":frame-text"), OdfConversionMappingStatus.Skipped,
                message: "Text outside the selected image/text-box story is preserved in ODF but is not combined with the selected text.");
    }

    private static bool IsFrameTextContainer(XElement element) =>
        element.Name == OdfNamespaces.Table + "table" ||
        element.Name.Namespace == OdfNamespaces.Text && !OdfTextCodec.IsNonVisibleTextElement(element);

    private static bool IsUnmappedFrameTextContainer(XElement element) =>
        IsFrameTextContainer(element) && !OdfTextTraversal.IsParagraph(element) &&
        element.Name != OdfNamespaces.Text + "list";

    private static TextFrameProjection ProjectText(OdgShape shape, OfficeDrawing target, OdfConversionReport report, double width, double height,
        DrawingFieldContext fields, LineTextFrame? lineFrame = null, CancellationToken cancellationToken = default,
        OfficeDrawingTextMetrics? layoutMetrics = null) {
        string feature = "shape:" + shape.Name + ":text";
        var losses = new HashSet<string>(StringComparer.Ordinal);
        try {
            bool growHeight = ReadProjectionGrowth(shape, "auto-grow-height", losses);
            bool growWidth = ReadProjectionGrowth(shape, "auto-grow-width", losses);
            TextBoxConstraints? constraints = null;
            if (HasInstanceTextBoxSize(shape)) {
                constraints = ResolveTextBoxConstraints(shape, ref width, ref height);
                if (constraints == null) losses.Add("text-size-constraints");
                else {
                    target = ResizeTextCanvas(target, width, height);
                    report.Add(feature + ":text-size-constraints", OdfConversionMappingStatus.Approximated,
                        message: "Absolute instance minima replace saved frame dimensions. Matching-unit maxima cap growth; graphic-style creation defaults do not resize saved frames. Source ODF is unchanged; native producer enforcement is unqualified.");
                }
            }
            if (fields.TextFlowParticipants.Contains(shape.Element)) losses.Add("text-chain-flow");
            var paragraphs = new List<OfficeRichTextParagraph>();
            if (shape.TextRoot.Elements().Any(IsUnmappedFrameTextContainer))
                losses.Add("unmapped-text-container");
            int probeBudget = OfficeTextLayoutEngine.MaximumLayoutTextCharacters;
            if (!HasDrawingText(shape.TextRoot, ref probeBudget)) { ReportLosses(); return Result(); }
            ReportEnhancedTextArea(shape, report);
            var lists = new OdfDrawingListResolver(shape, losses);
            var fieldResolver = new DrawingFieldResolver(fields, report, feature);
            int remaining = OfficeTextLayoutEngine.MaximumLayoutTextCharacters, runCount = 0, visitedNodes = 0;
            foreach (XElement element in OdfTextTraversal.Paragraphs(shape.TextRoot)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (paragraphs.Count > 0) remaining--;
                if (remaining < 0 || paragraphs.Count >= OfficeTextLayoutEngine.MaximumLayoutLines)
                    throw new NotSupportedException("Text exceeds the shared paragraph/character layout limit.");
                var source = new OdfTextParagraph(shape.Document, element, shape.Element);
                // Graphic orientation is independent of paragraph defaults; a horizontal paragraph
                // fallback cannot qualify a vertical graphic text area.
                if (DrawingWritingMode(source.WritingMode, shape, paragraph: true) is not (null or "lr-tb") ||
                    DrawingWritingMode(shape.ReadGraphicWritingMode(), shape) is not (null or "lr-tb"))
                    losses.Add("writing-mode");
                var runs = ReadDrawingRuns(source, shape, ref remaining, ref visitedNodes, losses, fieldResolver);
                if (runs.Count == 0) runs.Add(CreateDrawingRun(string.Empty, source, losses));
                runCount += runs.Count + (paragraphs.Count > 0 ? 1 : 0);
                if (runCount > OfficeTextLayoutEngine.MaximumLayoutTextRuns)
                    throw new NotSupportedException("Text exceeds the shared styled-run layout limit.");
                double left = ParagraphMargin(source, "margin-left"), right = ParagraphMargin(source, "margin-right");
                double top = ParagraphMargin(source, "margin-top"), bottom = ParagraphMargin(source, "margin-bottom");
                OdfLength? nativeIndent = source.Resolve(style => style.TextIndent);
                double indent = TextLength(nativeIndent, allowNegative: true);
                OfficeTextParagraphIndent paragraphIndent = OfficeTextParagraphIndent.FirstLine(Math.Max(0, indent));
                OdfDrawingListResolver.Entry? listEntry = lists.For(element);
                if (listEntry != null) paragraphIndent = OfficeTextParagraphIndent.Empty;
                else if (indent < 0) { left += indent; paragraphIndent = OfficeTextParagraphIndent.Hanging(-indent); }
                double paragraphTabOrigin = left;
                double firstCharacterSize = runs.FirstOrDefault(run => run.Text.Length > 0)?.FontSize ?? runs[0].FontSize;
                OfficeTextParagraphLabel? label = ProjectListLabel(listEntry, source, firstCharacterSize, losses, ref left, ref paragraphIndent, ref paragraphTabOrigin);
                if (left < 0) throw new NotSupportedException("Hanging paragraph text outside the frame is not projected.");
                if (label != null) {
                    XElement listLevel = listEntry!.Level;
                    if (listLevel.Element(OdfNamespaces.Style + "text-properties")?.Attribute(OdfNamespaces.Fo + "font-size") == null &&
                        listLevel.Attribute(OdfNamespaces.Text + "style-name") == null && listLevel.Attribute(OdfNamespaces.Text + "bullet-relative-size") == null)
                        report.Add(feature + ":list-label-font-defaults", OdfConversionMappingStatus.Approximated,
                            message: "The list label uses the first displayed character's font size, or the paragraph font size for an empty body; native label defaults can differ. Explicit label fonts provide a narrower interoperability profile.");
                    remaining -= label.Run.Text.Length + (label.FollowedBy == OfficeTextParagraphLabelFollowedBy.Nothing ? 0 : 1); runCount++;
                    if (remaining < 0 || runCount > OfficeTextLayoutEngine.MaximumLayoutTextRuns)
                        throw new NotSupportedException("List labels exceed the shared text layout limit.");
                }
                string? nativeLineHeight = ParagraphValue(source, OdfNamespaces.Fo + "line-height");
                double? lineHeight = null, lineFactor = null;
                if (nativeLineHeight is not (null or "normal")) {
                    OdfLength length = OdfLength.Parse(nativeLineHeight);
                    if (nativeLineHeight.EndsWith("%", StringComparison.Ordinal)) lineFactor = ResolveTextLength(length, 1);
                    else lineHeight = ResolveTextLength(length, 1);
                }
                if (nativeLineHeight != "normal" && (ParagraphValue(source, OdfNamespaces.Style + "line-height-at-least") != null ||
                    ParagraphValue(source, OdfNamespaces.Style + "line-spacing") != null)) losses.Add("additional-line-spacing");
                string? lastAlignment = ParagraphValue(source, OdfNamespaces.Fo + "text-align-last");
                if (lastAlignment is not (null or "start" or "left")) losses.Add("last-line-alignment");
                OfficeTextAlignment alignment = source.TextAlign switch {
                    null or "start" or "left" => OfficeTextAlignment.Left,
                    "center" => OfficeTextAlignment.Center,
                    "end" or "right" => OfficeTextAlignment.Right,
                    "justify" => OfficeTextAlignment.Justify,
                    _ => throw new NotSupportedException("Unsupported native paragraph alignment.")
                };
                var margins = new OfficeTextPadding(left, top, right, bottom);
                var projected = label == null ? new OfficeRichTextParagraph(runs, alignment, lineHeight, margins, paragraphIndent, lineFactor) :
                    new OfficeRichTextParagraph(runs, label, alignment, lineHeight, margins, paragraphIndent, lineFactor);
                if (runs.Any(r => r.Text.Contains('\t'))) {
                    projected = projected.WithTabStops(ProjectTabStops(source, left, paragraphTabOrigin, losses, report, feature));
                    if (alignment != OfficeTextAlignment.Left) {
                        // Native full-width text areas align the complete tabbed line.
                        // Centered/right hanging indents and list-body selection
                        // remain outside the qualified profile.
                        if (listEntry != null || alignment != OfficeTextAlignment.Justify && indent < 0)
                            losses.Add("tab-paragraph-alignment");
                        else report.Add(feature + ":tab-paragraph-alignment", OdfConversionMappingStatus.Approximated,
                            message: "Paragraph alignment positions complete tabbed lines while retaining field alignment within the tab grid. Native wrapping, default spacing and font placement can differ.");
                    }
                }
                paragraphs.Add(projected);
                if (alignment == OfficeTextAlignment.Justify && runs.Any(r => r.Text.Contains('\n'))) losses.Add("hard-break-justification");
                string joined = string.Concat(runs.Select(r => r.Text));
                if (joined.Length > 0 && char.IsWhiteSpace(joined[0]) && joined[0] != '\t' ||
                    projected.TabStops == null && (joined.Contains("\n ") || joined.Contains("\n\t"))) losses.Add("leading-whitespace");
            }
            if (paragraphs.Count == 0) { ReportLosses(); return Result(); }
            OfficeTextPadding padding = new OfficeTextPadding(GraphicTextLength("padding-left"), GraphicTextLength("padding-top"),
                GraphicTextLength("padding-right"), GraphicTextLength("padding-bottom"));
            OfficeTextVerticalAlignment vertical = shape.ReadGraphicProperty(OdfNamespaces.Draw + "textarea-vertical-align") switch {
                null => lineFrame.HasValue ? OfficeTextVerticalAlignment.Center : OfficeTextVerticalAlignment.Top,
                "top" => OfficeTextVerticalAlignment.Top,
                "middle" => OfficeTextVerticalAlignment.Center,
                "bottom" => OfficeTextVerticalAlignment.Bottom,
                _ => throw new NotSupportedException("Unsupported native text-area vertical alignment.")
            };
            string? wrap = shape.ReadGraphicProperty(OdfNamespaces.Fo + "wrap-option");
            if (wrap == null) report.Add(feature + ":implicit-wrap-option", OdfConversionMappingStatus.Approximated,
                message: lineFrame.HasValue ? "An omitted line-label wrapping option uses unwrapped text; producer defaults can differ." :
                    "An omitted native wrapping option uses the shared wrapped-text default; producer defaults can differ.");
            if (paragraphs.Count > 0 && OdfTextTraversal.Paragraphs(shape.TextRoot).Any(p => {
                var source = new OdfTextParagraph(shape.Document, p, shape.Element);
                return source.Styles.FirstOrDefault(s => s.FontSize.HasValue)?.Family == OdfStyleFamily.Graphic ||
                    source.Styles.FirstOrDefault(s => s.FontFamily != null)?.Family == OdfStyleFamily.Graphic;
            })) report.Add(feature + ":graphic-text-defaults", OdfConversionMappingStatus.Approximated,
                message: "Graphic-level font inheritance is projected; native producers can ignore these defaults. Explicit paragraph fonts provide a narrower interoperability profile.");
            if (new[] { "padding-left", "padding-right", "padding-top", "padding-bottom" }.Any(name =>
                shape.ReadGraphicProperty(OdfNamespaces.Fo + name, OdfNamespaces.Fo + "padding") != shape.ReadGraphicProperty(OdfNamespaces.Fo + name)))
                report.Add(feature + ":padding-shorthand", OdfConversionMappingStatus.Approximated,
                    message: "The shared layout resolves padding shorthand; native producers can ignore it. Explicit side padding provides a narrower interoperability profile.");
            if (wrap is not (null or "wrap" or "no-wrap")) losses.Add("wrap-option");
            if (shape.ReadGraphicProperty(OdfNamespaces.Draw + "fit-to-contour") == "true") losses.Add("text-contour");
            bool autoGrow = growHeight || growWidth || losses.Contains("text-auto-size");
            bool shrinkToFit = ProjectTextFitting(shape, lineFrame.HasValue, autoGrow, losses);
            string? horizontalArea = shape.ReadGraphicProperty(OdfNamespaces.Draw + "textarea-horizontal-align");
            bool wrapText = wrap != "no-wrap";
            OfficeTextAreaAlignment areaAlignment = OfficeTextAreaAlignment.FullWidth;
            if (lineFrame.HasValue) {
                if (horizontalArea is not (null or "justify" or "left" or "center" or "right")) losses.Add("text-area-alignment");
            } else areaAlignment = ProjectTextArea(shape, horizontalArea, paragraphs, report, feature, losses, shrinkToFit, ref wrapText);
            if (shape.ReadGraphicProperty(OdfNamespaces.Style + "overflow-behavior") is not (null or "clip")) losses.Add("text-overflow");
            if (lineFrame.HasValue) {
                if (paragraphs.All(paragraph => paragraph.Runs.All(run => run.Text.Length == 0) &&
                    (paragraph.Label == null || paragraph.Label.Run.Text.Length == 0))) { ReportLosses(); return Result(); }
                if (autoGrow) losses.Add("text-auto-size");
                AddLineLabel(shape, target, paragraphs, lineFrame.Value, vertical, wrap, padding, report, feature,
                    cancellationToken, layoutMetrics);
                if (lineFrame.Value.Bounds.Left == lineFrame.Value.Bounds.Right &&
                    (horizontalArea is "left" or "right" || (horizontalArea is null or "justify") &&
                        paragraphs.Any(p => p.Alignment != OfficeTextAlignment.Center)))
                    report.Add(feature + ":label-degenerate-area", OdfConversionMappingStatus.Approximated,
                        message: "The caption is retained near its zero-width route. Native text-area expansion can place left/right areas or noncentered justified paragraphs far off-page; that producer-specific placement is not reproduced.");
                if (wrap == "wrap") report.Add(feature + ":label-wrapping", OdfConversionMappingStatus.Unsupported,
                    message: "The complete line/connector caption is projected unwrapped. Native Draw line labels can ignore wrapping declarations; width-constrained wrapping is not reproduced. Use explicit paragraph or line breaks when native caption layout matters.");
            }
            else {
                target = GrowTextBox(shape, target, paragraphs, ref width, ref height, padding, areaAlignment, vertical, wrapText,
                    growHeight, growWidth, losses, report, feature, cancellationToken, layoutMetrics, constraints,
                    out bool normalizeHorizontalPaint);
                AddFixedFrameText(target, paragraphs, width, height, areaAlignment, vertical, wrapText, padding,
                    shrinkToFit, report, feature, cancellationToken, layoutMetrics, reportHorizontalClipping: growWidth,
                    normalizeHorizontalPaint: normalizeHorizontalPaint);
            }
            report.Add(feature, OdfConversionMappingStatus.Approximated,
                message: lineFrame.HasValue ? "An unwrapped label uses shared paragraph measurement and saved route bounds; native font metrics, rerouting and later font-provider changes can alter placement." :
                    "Styled runs, paragraph alignment, margins, indentation, line spacing and frame padding use shared render-time font layout; native metrics, auto-sizing and overflow can differ.");
            ReportLosses();

            void ReportLosses() {
                foreach (string loss in losses) report.Add(feature + ":" + loss, OdfConversionMappingStatus.Unsupported,
                    message: "Native " + loss + " is retained in ODF but is not fully reproduced in this text projection.");
            }

            double GraphicTextLength(string name) {
                string? value = shape.ReadGraphicProperty(OdfNamespaces.Fo + name, OdfNamespaces.Fo + "padding");
                return value == null ? 0 : TextLength(OdfLength.Parse(value));
            }
        } catch (Exception exception) when (exception is ArgumentException or InvalidDataException or NotSupportedException or OverflowException or FormatException) {
            report.Add(feature, OdfConversionMappingStatus.Skipped, message: exception.Message);
        }
        return Result();

        TextFrameProjection Result() => new TextFrameProjection(target, width, height);
    }

    private static bool HasDrawingText(XElement root, ref int budget) {
        foreach (XElement paragraph in OdfTextTraversal.Paragraphs(root)) {
            if (OdfTextCodec.ReadNodes(paragraph.Nodes(), ref budget).Length > 0) return true;
            if (paragraph.Parent?.Name == OdfNamespaces.Text + "list-item") return true;
            // Empty field caches can gain visible text when a native application refreshes them.
            if (paragraph.Descendants().Any(e => OdfTextField.IsField(e.Name) &&
                !e.Ancestors().TakeWhile(a => a != paragraph).Any(OdfTextCodec.IsNonVisibleTextElement))) return true;
        }
        return false;
    }

    private static string? DrawingWritingMode(string? mode, OdgShape shape, bool paragraph = false) {
        if (mode != "page") return mode;
        if (paragraph) {
            string? graphic = shape.ReadGraphicWritingMode();
            if (graphic != null && graphic != "page") return graphic;
        }
        // A graphic/paragraph page token inherits the containing page layout, including master selection.
        XElement? page = shape.Element.Ancestors(OdfNamespaces.Draw + "page").FirstOrDefault();
        if (shape.Document is not OdgDocument document) return "page";
        if (page != null) return (string?)new OdgPage(document, page).LayoutProperties.Attribute(OdfNamespaces.Style + "writing-mode");
        XElement? master = shape.Element.Ancestors(OdfNamespaces.Style + "master-page").FirstOrDefault();
        return master == null ? "page" : (string?)new OdgPage(document, master)
            .ResolveLayoutProperties(master).Attribute(OdfNamespaces.Style + "writing-mode");
    }

    private static string? ParagraphValue(OdfTextParagraph source, XName name, XName? shorthand = null) => source.Styles
        .Select(s => (string?)s.ParagraphProperties?.Attribute(name) ?? (shorthand == null ? null : (string?)s.ParagraphProperties?.Attribute(shorthand)))
        .FirstOrDefault(v => v != null);

    private static double ParagraphMargin(OdfTextParagraph source, string name) {
        string? value = ParagraphValue(source, OdfNamespaces.Fo + name, OdfNamespaces.Fo + "margin");
        return value == null ? 0 : TextLength(OdfLength.Parse(value));
    }

    private static double TextLength(OdfLength? length, bool allowNegative = false) {
        if (!length.HasValue) return 0;
        if (!length.Value.TryToPoints(out double points) || !allowNegative && points < 0)
            throw new NotSupportedException("Text margins, padding and indentation require supported absolute lengths.");
        return points;
    }

    private static double ResolveTextLength(OdfLength length, double inherited) {
        string value = length.ToString();
        double result;
        if (value.EndsWith("%", StringComparison.Ordinal) && double.TryParse(value.Substring(0, value.Length - 1),
            NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)) result = inherited * percent / 100D;
        else result = length.ToPoints();
        if (result <= 0 || double.IsNaN(result) || double.IsInfinity(result)) throw new NotSupportedException("Text size and line spacing must be finite positive lengths.");
        return result;
    }
}
