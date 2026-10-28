using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int? RegisterInlineFigureStructureElement(
        LayoutResult.Page? page,
        PdfCreateOptions options,
        string? alternativeText,
        int? parentElementIndex) {
        if (page == null ||
            options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers ||
            string.IsNullOrWhiteSpace(alternativeText)) {
            return null;
        }

        int markedContentId = page.NextMarkedContentId++;
        page.StructElements.Add(new PageStructElement {
            MarkedContentId = markedContentId,
            StructureType = "Figure",
            AlternativeText = alternativeText!,
            ParentElementIndex = parentElementIndex
        });
        return markedContentId;
    }

    private static void AppendInlineElement(
        StringBuilder sb,
        PdfInlineElement inlineElement,
        double x,
        double baseline,
        PdfCreateOptions options,
        LayoutResult.Page? page,
        int? parentElementIndex) {
        double bottom = baseline + inlineElement.BaselineOffset;
        if (inlineElement is PdfInlineImage inlineImage) {
            if (page == null) {
                throw new InvalidOperationException("Inline images require an active output page.");
            }

            ImageBlock block = inlineImage.Block;
            PageImage pageImage = CreatePageImage(
                block,
                block.Style ?? new PdfImageStyle(),
                inlineImage.TableTextRotation == 0 ? x : x + (inlineElement.Width - inlineElement.Height) / 2D,
                inlineImage.TableTextRotation == 0 ? bottom : bottom + (inlineElement.Height - inlineElement.Width) / 2D,
                inlineImage.TableTextRotation == 0 ? inlineElement.Width : inlineElement.Height,
                inlineImage.TableTextRotation == 0 ? inlineElement.Height : inlineElement.Width);
            if (inlineImage.TableTextRotation != 0) {
                pageImage.RotationAngle -= inlineImage.TableTextRotation;
                pageImage.RotationCenterX = x + inlineElement.Width / 2D;
                pageImage.RotationCenterY = bottom + inlineElement.Height / 2D;
            }
            pageImage.IsInlineDecoration = string.IsNullOrWhiteSpace(inlineElement.AlternativeText);
            page.Images.Add(pageImage);
            pageImage.InlineDrawToken = AllocateInlineImageDrawToken(page);
            if (!string.IsNullOrWhiteSpace(inlineElement.AlternativeText)) {
                pageImage.MarkedContentId = RegisterInlineFigureStructureElement(page, options, inlineElement.AlternativeText, parentElementIndex);
                pageImage.StructElementIndex = FindStructElementIndex(page, pageImage.MarkedContentId, "Figure");
            }

            sb.Append(pageImage.InlineDrawToken);
            return;
        }

        if (inlineElement is PdfInlineFormWidget inlineWidget) {
            if (page == null) throw new InvalidOperationException("Inline form fields require an active output page.");
            PrepareInlineCheckBoxGlyphUsage(inlineWidget.Field, options);
            double left = x, right = x + inlineWidget.Width, top = bottom + inlineWidget.Height;
            if (inlineWidget.SourceRun is { } source) {
                double cursor = x;
                foreach (PdfTextRun run in NormalizeFallbackRuns(new[] { source }, ChooseNormal(options.DefaultFont), options)) {
                    RichSeg segment = CreateInlineAppearanceSegment(run, options);
                    IncludeInlineFormPaintBounds(page, segment, cursor, baseline, options,
                        ref left, ref bottom, ref right, ref top);
                    cursor += GetRichSegmentWidth(segment);
                }
            }
            FormFieldAnnotation annotation = CreateInlineFormAnnotation(inlineWidget.Field, left, bottom,
                right, top, parentElementIndex);
            annotation.InlineWidgetBaseline = baseline - bottom;
            annotation.InlineWidgetTextX = x - left;
            page.FormFields.Add(annotation);
            return;
        }

        PdfInlineBox inlineBox = (PdfInlineBox)inlineElement;
        bool tagged = options.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers;
        int? figureMarkedContentId = RegisterInlineFigureStructureElement(page, options, inlineElement.AlternativeText, parentElementIndex);
        bool hasAlternativeText = !string.IsNullOrWhiteSpace(inlineElement.AlternativeText);
        if (hasAlternativeText) {
            sb.Append("/Figure << /Alt ")
                .Append(PdfSyntaxEscaper.TextString(inlineElement.AlternativeText!));
            if (figureMarkedContentId.HasValue) {
                sb.Append(" /MCID ")
                    .Append(figureMarkedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }

            sb.Append(" >> BDC\n");
        } else {
            AppendArtifactBegin(sb, tagged);
        }

        ContentStreamBuilder boxContent = new ContentStreamBuilder(sb).SaveState();
        if (inlineBox.SourceRule is { } rule) {
            if (inlineBox.Width > 0D) {
                foreach (PdfInlineRule.InkBand band in rule.Bands) {
                    boxContent
                        .FillColor(band.Color)
                        .Rectangle(x, baseline + band.BaselineOffset, inlineBox.Width, band.Height)
                        .FillPath();
                }
            }
        } else if (inlineBox.Background.HasValue) {
            boxContent
                .FillColor(inlineBox.Background.Value)
                .Rectangle(x, bottom, inlineBox.Width, inlineBox.Height)
                .FillPath();
        }

        if (inlineBox.BorderColor.HasValue && inlineBox.BorderWidth > 0D) {
            boxContent
                .StrokeColor(inlineBox.BorderColor.Value)
                .LineWidth(inlineBox.BorderWidth)
                .Rectangle(x, bottom, inlineBox.Width, inlineBox.Height)
                .StrokePath();
        }

        boxContent.RestoreState();
        if (hasAlternativeText) {
            sb.Append("EMC\n");
        } else {
            AppendArtifactEnd(sb, tagged);
        }
    }

    private static void WriteRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfCreateOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double? xOverride = null, double? widthOverride = null, double? firstLineXOverride = null, double? firstLineWidthOverride = null, string? structureType = null, int? markedContentId = null, LayoutResult.Page? structurePage = null, System.Collections.Generic.IReadOnlyList<PdfAlign?>? lineAlignments = null, System.Collections.Generic.IReadOnlyList<double>? lineXOffsets = null, System.Collections.Generic.IReadOnlyList<double>? lineWidths = null, bool suppressActualText = false, System.Collections.Generic.IReadOnlyList<double>? lineTopGaps = null, PdfStandardFont? baselineFont = null, string? foregroundGraphicsState = null, string? decorationGraphicsState = null, bool captureInlineForms = true, bool hasFollowingParagraphLines = false, LayoutResult.Page? fontPage = null, double paragraphSpacingAfter = 0D, TableCellContentFrame? barClipFrame = null, double paragraphSpacingBefore = 0D, Action<Action>? foregroundGroup = null) {
        List<List<RichSeg>> logicalRows = lines;
        RichPaintRows? paintRows = ExpandRichPaintRows(lines, lineHeights, defaultLeading,
            xOverride ?? opts.MarginLeft, widthOverride ?? opts.PageWidth - opts.MarginLeft - opts.MarginRight,
            firstLineXOverride, firstLineWidthOverride, lineXOffsets, lineWidths, lineTopGaps);
        if (paintRows != null) {
            lines = paintRows.Lines;
            lineHeights = paintRows.Heights;
            lineXOffsets = paintRows.Offsets;
            lineWidths = paintRows.Widths;
            lineTopGaps = paintRows.Gaps;
            firstLineXOverride = null;
            firstLineWidthOverride = null;
        }
        // Canvas artifacts can suppress accessibility structure while still
        // owning page glyphs and their font resources.
        fontPage ??= structurePage;
        double widthContent = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double widthUsed = widthOverride ?? widthContent;
        System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color, OfficeIMO.Drawing.OfficeTextDecorationStyle Style)>? underlines = null;
        System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color, OfficeIMO.Drawing.OfficeTextDecorationStyle Style)>? strikes = null;
        System.Collections.Generic.List<(double X, double Y, double Width, double Height, PdfColor Color)>? backgrounds = null;

        void AddBackground(double x, double y, double width, double height, PdfColor color) {
            if (width <= 0.001D || height <= 0.001D) {
                return;
            }

            backgrounds ??= new System.Collections.Generic.List<(double X, double Y, double Width, double Height, PdfColor Color)>();
            if (backgrounds.Count > 0) {
                var previous = backgrounds[backgrounds.Count - 1];
                if (previous.Color.Equals(color) &&
                    Math.Abs(previous.Y - y) <= 0.001D &&
                    Math.Abs(previous.Height - height) <= 0.001D &&
                    x <= previous.X + previous.Width + 0.25D) {
                    double left = Math.Min(previous.X, x);
                    double right = Math.Max(previous.X + previous.Width, x + width);
                    backgrounds[backgrounds.Count - 1] = (left, previous.Y, right - left, previous.Height, previous.Color);
                    return;
                }
            }

            backgrounds.Add((x, y, width, height, color));
        }

        double backgroundYOffset = 0D;
        double xOrigin = xOverride ?? opts.MarginLeft;
        int?[] bidiParagraphLevels = ResolveRichParagraphBidiLevels(lines);
        var justification = new RichLineJustification[lines.Count];
        for (int index = 0; index < lines.Count; index++) {
            justification[index] = PlanRichLineJustification(lines[index], ResolveRichLineAlignment(block.Align, lineAlignments, index),
                block.Style ?? opts.DefaultParagraphStyleSnapshot,
                ResolveRichLineWidth(widthUsed, firstLineWidthOverride, lineWidths, index),
                index == lines.Count - 1 && !hasFollowingParagraphLines, opts);
        }
        for (int li = 0; li < lines.Count; li++) {
            var segs = justification[li].Segments;
            bool hasLineBackground = false;
            for (int si = 0; si < segs.Count; si++) {
                if (segs[si].BackgroundColor.HasValue && (!captureInlineForms || segs[si].FormField == null)) {
                    hasLineBackground = true;
                    break;
                }
            }

            if (!hasLineBackground) {
                backgroundYOffset += li < lineHeights.Count ? lineHeights[li] : defaultLeading;
                continue;
            }

            double lineY = AdjustRichLineBaseline(startY - backgroundYOffset - (lineTopGaps != null ? lineTopGaps[li] : 0D), lines[li], opts, fontSize, baselineFont);
            double lineWidthUsed = ResolveRichLineWidth(widthUsed, firstLineWidthOverride, lineWidths, li);
            double lineXOrigin = ResolveRichLineXOrigin(xOrigin, firstLineXOverride, lineXOffsets, li);
            PdfAlign lineAlign = ResolveRichLineAlignment(block.Align, lineAlignments, li);
            double wordSpacing = justification[li].WordSpacing;
            double lineWForAlign = justification[li].Width;
            double dx = 0;
            if (lineAlign == PdfAlign.Center) dx = Math.Max(0, (lineWidthUsed - lineWForAlign) / 2);
            else if (lineAlign == PdfAlign.Right) dx = Math.Max(0, lineWidthUsed - lineWForAlign);

            double xCursor = dx;
            BidiLineLayout? backgroundBidi = bidiParagraphLevels[li].HasValue ? CreateBidiLineLayout(segs, bidiParagraphLevels[li]!.Value, wordSpacing) : null;
            for (int si = 0; si < segs.Count; si++) {
                var s = segs[si];
                double leadingAdvance = 0D;
                if (s.LeadingSpace) {
                    double baseGap = s.LeadingAdvance;
                    leadingAdvance = baseGap + (s.LeadingSpaceIsExpandable ? wordSpacing * s.LeadingSpaceCount : 0);
                    xCursor += leadingAdvance;
                }

                double wSeg = GetRichSegmentWidth(s);
                if (backgroundBidi != null) {
                    // Word and gap occupy separate visual positions; highlight each.
                    if (s.BackgroundColor.HasValue && (!captureInlineForms || s.FormField == null) && leadingAdvance > 0D) {
                        double gapFontSize = EffectiveRichFontSize(s.FontSize, s.Baseline);
                        double gapBaseline = lineY + TextRiseForBaseline(s.FontSize, s.Baseline) + s.LineAlignmentRise;
                        double gapAscent = GetAscenderForOptions(s.Font, s.NamedFont, gapFontSize, opts);
                        double gapDescent = GetDescenderForOptions(s.Font, s.NamedFont, gapFontSize, opts);
                        double gapPad = GetRichTextBackgroundPadding(gapFontSize).Vertical;
                        AddBackground(lineXOrigin + dx + backgroundBidi.GapX[si], gapBaseline - gapDescent - gapPad,
                            leadingAdvance, gapAscent + gapDescent + gapPad * 2D, s.BackgroundColor.Value);
                    }
                    xCursor = dx + backgroundBidi.VisualX[si];
                    leadingAdvance = 0D;
                }
                if (s.BackgroundColor.HasValue && (!captureInlineForms || s.FormField == null) && (wSeg > 0 || leadingAdvance > 0)) {
                    double runFontSize = EffectiveRichFontSize(s.FontSize, s.Baseline);
                    double textRise = TextRiseForBaseline(s.FontSize, s.Baseline) + s.LineAlignmentRise;
                    double asc = GetAscenderForOptions(s.Font, s.NamedFont, runFontSize, opts);
                    double desc = GetDescenderForOptions(s.Font, s.NamedFont, runFontSize, opts);
                    var (padX, padY) = GetRichTextBackgroundPadding(runFontSize);
                    double baselineY = lineY + textRise;
                    AddBackground(
                        lineXOrigin + xCursor - leadingAdvance - padX,
                        baselineY - desc - padY,
                        wSeg + leadingAdvance + (padX * 2D),
                        asc + desc + (padY * 2D),
                        s.BackgroundColor.Value);
                }

                xCursor += wSeg;
            }

            backgroundYOffset += li < lineHeights.Count ? lineHeights[li] : defaultLeading;
        }

        if (backgrounds != null) {
            AppendArtifactBegin(sb, markedContentId.HasValue);
            foreach (var bg in backgrounds) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .FillColor(bg.Color)
                    .Rectangle(bg.X, bg.Y, bg.Width, bg.Height)
                    .FillPath()
                    .RestoreState();
            }

            AppendArtifactEnd(sb, markedContentId.HasValue);
        }

        int foregroundStart = sb.Length;
        // Drawing foreground alpha belongs to glyph paint, independently of run
        // backgrounds and the decoration paint emitted after the text object.
        if (foregroundGraphicsState != null) {
            new ContentStreamBuilder(sb).SaveState().GraphicsState(foregroundGraphicsState);
        }
        // The logical replacement encloses all paragraph/link marked-content
        // transitions, so a hyperlink cannot close the identifier span early.
        bool optionalBreakText = TryBeginOptionalBreakText(sb, block, logicalRows, suppressActualText, opts, xOrigin, startY, widthUsed, fontSize);
        if (optionalBreakText) suppressActualText = true;
        AppendMarkedContentBegin(sb, structureType, markedContentId);
        bool textMarkedContentOpen = markedContentId.HasValue;
        int? textStructElementIndex = FindStructElementIndex(structurePage, markedContentId, structureType);
        ContentStreamBuilder content = new ContentStreamBuilder(sb)
            .BeginText()
            .TextLeading(defaultLeading);

        double yOffset = 0D;
        StringBuilder? uprightCellImages = null;
        for (int li = 0; li < lines.Count; li++) {
            double lineY = AdjustRichLineBaseline(startY - yOffset - (lineTopGaps != null ? lineTopGaps[li] : 0D), lines[li], opts, fontSize, baselineFont);
            double lineWidthUsed = ResolveRichLineWidth(widthUsed, firstLineWidthOverride, lineWidths, li);
            double lineXOrigin = ResolveRichLineXOrigin(xOrigin, firstLineXOverride, lineXOffsets, li);
            var segs = justification[li].Segments;
            PdfAlign lineAlign = ResolveRichLineAlignment(block.Align, lineAlignments, li);
            double wordSpacing = justification[li].WordSpacing;
            // Tw applies only to single-byte code 32. Composite-font lines are positioned explicitly.
            double pdfWordSpacing = wordSpacing != 0D && !segs.All(seg => UsesSimpleFontText(seg, opts)) ? 0D : wordSpacing;
            double lineWForAlign = justification[li].Width;
            double dx = 0;
            if (lineAlign == PdfAlign.Center) dx = Math.Max(0, (lineWidthUsed - lineWForAlign) / 2);
            else if (lineAlign == PdfAlign.Right) dx = Math.Max(0, lineWidthUsed - lineWForAlign);
            content
                .TextMatrix(lineXOrigin + dx, lineY)
                .WordSpacing(pdfWordSpacing);

            double xCursor = dx;
            // Right-to-left text keeps logical content order; words move to their UAX #9 positions.
            BidiLineLayout? bidi = bidiParagraphLevels[li].HasValue ? CreateBidiLineLayout(segs, bidiParagraphLevels[li]!.Value, wordSpacing) : null;
            double currentTextRise = 0;
            for (int si = 0; si < segs.Count; si++) {
                var s = segs[si];
                string fontRes = GetFontResourceName(s.Font, s.NamedFont, ChooseNormal(opts.DefaultFont));
                double runFontSize = EffectiveRichFontSize(s.FontSize, s.Baseline);
                double textRise = TextRiseForBaseline(s.FontSize, s.Baseline) + s.LineAlignmentRise;
                bool syntheticOblique = opts.NeedsSyntheticOblique(s.Font, s.NamedFont);
                bool syntheticBold = opts.NeedsSyntheticBold(s.Font, s.NamedFont);
                bool inlineSeparator = s.InlineElement != null && s.InlineElement is not PdfInlineFormWidget && s.LeadingSpace;
                string? inlineLeader = inlineSeparator && s.LeadingTabLeader != PdfTabLeaderStyle.None
                    ? BuildTabLeaderText(s.LeadingAdvance + (s.LeadingSpaceIsExpandable ? wordSpacing * s.LeadingSpaceCount : 0D), s, opts)
                    : null;
                bool paintsPageText = s.InlineElement == null || inlineSeparator &&
                    (s.LeadingTabLeader == PdfTabLeaderStyle.None || inlineLeader!.Length > 0);
                if (paintsPageText) {
                    // Inline geometry needs font metrics, but only a real page
                    // glyph needs a Tf resource. Its separator may own that glyph.
                    if (s.InlineElement != null && fontPage != null)
                        MarkPageRichTextFont(fontPage, opts, s.Font, s.NamedFont);
                    content.Font(fontRes, runFontSize, syntheticOblique, syntheticBold: syntheticBold);
                    if (Math.Abs(textRise - currentTextRise) > 0.0001) {
                        content.TextRise(textRise);
                        currentTextRise = textRise;
                    }
                }

                var color = s.Color ?? block.DefaultColor ?? opts.DefaultTextColor;
                content.FillColor(color ?? PdfColor.Black);
                bool hasLinkTarget = !string.IsNullOrEmpty(s.Uri) || !string.IsNullOrEmpty(s.DestinationName);
                if (paintsPageText && (!hasLinkTarget || s.LeadingSpace)) {
                    EnsureTextMarkedContentOpen(
                        sb,
                        ref content,
                        ref textMarkedContentOpen,
                        structurePage,
                        textStructElementIndex,
                        structureType,
                        defaultLeading,
                        lineXOrigin + xCursor,
                        lineY,
                        pdfWordSpacing,
                        fontRes,
                        syntheticOblique,
                        syntheticBold,
                        runFontSize,
                        textRise,
                        color ?? PdfColor.Black,
                        hasLinkTarget && s.LeadingSpace && s.LeadingTabLeader == PdfTabLeaderStyle.None && s.InlineElement == null);
                }

                double inlineFormLeadingAdvance = 0D;
                if (s.LeadingSpace) {
                    double baseGap = s.LeadingAdvance;
                    double gap = baseGap + (s.LeadingSpaceIsExpandable ? wordSpacing * s.LeadingSpaceCount : 0);
                    inlineFormLeadingAdvance = gap;
                    if (bidi != null) {
                        xCursor = dx + bidi.GapX[si];
                        content.TextMatrix(lineXOrigin + xCursor, lineY);
                    }
                    if (s.LeadingUnderlineStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.None && s.LeadingTabLeader == PdfTabLeaderStyle.None && s.LeadingSpaceIsExpandable) {
                        PdfColor underlineColor = s.LeadingDecorationColor ?? block.DefaultColor ?? opts.DefaultTextColor ?? PdfColor.Black;
                        double underlineY = lineY + s.LeadingDecorationTextRise - s.LeadingDecorationFontSize * 0.15;
                        (underlines ??= new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color, OfficeIMO.Drawing.OfficeTextDecorationStyle Style)>())
                            .Add((lineXOrigin + xCursor, lineXOrigin + xCursor + gap, underlineY, underlineColor, s.LeadingUnderlineStyle));
                    }

                    if (s.InlineElement is PdfInlineFormWidget) {
                        // A control's measured separator has no ink. Icon-only
                        // fonts need not contain a SPACE glyph to advance it.
                        xCursor += gap;
                        content.BreakTextRun().TextMatrix(lineXOrigin + xCursor, lineY);
                    } else if (s.LeadingTabLeader != PdfTabLeaderStyle.None) {
                        string leader = inlineLeader ?? BuildTabLeaderText(gap, s, opts);
                        // Tab stops stay explicit text positions so readers keep the columns.
                        content.BreakTextRun();
                        if (leader.Length > 0) {
                            ApplyRichTextSpacing(content, s.HorizontalTextScaling, s.CharacterSpacing, pdfWordSpacing);
                            content
                                .TextMatrix(lineXOrigin + xCursor, lineY)
                                .ShowText(EncodeTextShowCommand(leader, s.Font, s.NamedFont, opts, s.FeatureSettings, s.TextDirection, s.FontMetricScale), runFontSize, textRise, suppressActualText);
                            ResetRichTextSpacing(content, s.HorizontalTextScaling, s.CharacterSpacing, pdfWordSpacing);
                        }
                        xCursor += gap;
                        content.BreakTextRun().TextMatrix(lineXOrigin + xCursor, lineY);
                    } else if (s.LeadingIsTab) {
                        // A plain tab is geometry; its source font may give SPACE visible ink.
                        xCursor += gap;
                        content.BreakTextRun().TextMatrix(lineXOrigin + xCursor, lineY);
                    } else if (!s.LeadingSpaceIsExpandable) {
                        content
                            .TextMatrix(lineXOrigin + xCursor, lineY)
                            .ShowText(EncodeTextShowCommand(" ", s.Font, s.NamedFont, opts, s.FeatureSettings, fontMetricScale: s.FontMetricScale), runFontSize, textRise, suppressActualText);
                        xCursor += gap;
                        content.TextMatrix(lineXOrigin + xCursor, lineY);
                    } else {
                        content.ShowText(EncodeTextShowCommand(" ", s.Font, s.NamedFont, opts, s.FeatureSettings, fontMetricScale: s.FontMetricScale), runFontSize, textRise, suppressActualText);
                        xCursor += gap;
                        // A separator can inherit metrics from the previous run. Its
                        // measured advance also handles CID spaces to which Tw does not apply.
                        double paintedGap = MeasureRichText(" ", s.Font, s.NamedFont, s.FontSize, s.Baseline, opts, s.FeatureSettings);
                        if (Math.Abs(wordSpacing) > 0.0001 || Math.Abs(gap - paintedGap) > 0.0001)
                            content.TextMatrix(lineXOrigin + xCursor, lineY);
                    }
                }
                if (s.InlineElement != null) {
                    content.EndText();
                    if (textMarkedContentOpen) {
                        AppendMarkedContentEnd(sb, markedContentId);
                        textMarkedContentOpen = false;
                    }

                    if (structurePage != null) PromoteTextStructureContainer(structurePage, textStructElementIndex);
                    // Upright pictures keep the paragraph's top-to-bottom advance even
                    // when counterclockwise cell text advances from bottom to top.
                    double inlineX = s.InlineElement is PdfInlineImage { TableTextRotation: > 0 }
                        ? lineXOrigin + lineWidthUsed - xCursor - s.InlineElement.Width
                        : lineXOrigin + xCursor;
                    bool pictureClip = BeginImportedPictureClip(sb, s.InlineElement,
                        (block.Style ?? opts.DefaultParagraphStyleSnapshot)?.LineSpacing, inlineX, lineY,
                        startY + GetAscenderForOptions(baselineFont ?? ChooseNormal(opts.DefaultFont), fontSize, opts), opts.PageWidth);
                    AppendInlineElement(
                        s.InlineElement is PdfInlineImage { TableTextRotation: not 0 }
                            ? uprightCellImages ??= new StringBuilder()
                            : sb,
                        s.InlineElement,
                        inlineX,
                        s.InlineElement is PdfInlineFormWidget or PdfInlineBox { UsesRunTextMetrics: true }
                            ? lineY + s.LineAlignmentRise : lineY,
                        opts,
                        structurePage,
                        textStructElementIndex);
                    if (pictureClip) new ContentStreamBuilder(sb).RestoreState();
                    xCursor += s.InlineElement.Width + s.JustificationAdvance;
                    content = new ContentStreamBuilder(sb)
                        .BeginText()
                        .TextLeading(defaultLeading)
                        .TextMatrix(lineXOrigin + xCursor, lineY)
                        .WordSpacing(pdfWordSpacing);
                    currentTextRise = 0D;
                    continue;
                }

                double wSeg = GetRichSegmentWidth(s);
                if (captureInlineForms && s.FormField != null) {
                    if (structurePage == null) throw new InvalidOperationException("Inline form text requires an active output page.");
                    if (bidi != null) xCursor = dx + bidi.VisualX[si];
                    CaptureInlineFormText(structurePage, s, lineXOrigin + xCursor, lineY,
                        block.DefaultColor ?? opts.DefaultTextColor, opts, textStructElementIndex, inlineFormLeadingAdvance);
                    xCursor += wSeg;
                    ResetRichTextSpacing(content, s.HorizontalTextScaling, s.CharacterSpacing, pdfWordSpacing);
                    content.TextMatrix(lineXOrigin + xCursor, lineY);
                    continue;
                }

                int? linkMarkedContentId = null;
                int? linkStructElementIndex = null;
                if (hasLinkTarget && wSeg > 0D && opts.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers && structurePage != null) {
                    linkMarkedContentId = structurePage.NextMarkedContentId++;
                    PromoteTextStructureContainer(structurePage, textStructElementIndex);
                    linkStructElementIndex = structurePage.StructElements.Count;
                    structurePage.StructElements.Add(new PageStructElement {
                        MarkedContentId = linkMarkedContentId,
                        StructureType = "Link",
                        ParentElementIndex = textStructElementIndex
                    });
                }

                if (bidi != null) {
                    xCursor = dx + bidi.VisualX[si];
                    content.TextMatrix(lineXOrigin + xCursor, lineY);
                }
                double segmentStartX = xCursor;
                PdfTextShowCommand textCommand = s.KerningCommand ?? EncodeTextShowCommand(s.Text, s.Font, s.NamedFont, opts, s.FeatureSettings, s.TextDirection, s.FontMetricScale,
                    preservesPaintedGlyphs: s.PreservesPaintedGlyphs);
                // Imported layout uses design-unit advances while PDF widths are
                // rounded. Anchor each run before those differences accumulate
                // across paint boundaries or reach later inline decorations.
                if (opts.UseFontUnitTextLayout || HasRichTextSpacing(s.HorizontalTextScaling, s.CharacterSpacing)) {
                    content.TextMatrix(lineXOrigin + xCursor, lineY);
                }
                if (HasRichTextSpacing(s.HorizontalTextScaling, s.CharacterSpacing)) {
                    ApplyRichTextSpacing(content, s.HorizontalTextScaling, s.CharacterSpacing, pdfWordSpacing);
                }
                if (linkMarkedContentId.HasValue) {
                    content.EndText();
                    if (textMarkedContentOpen) {
                        AppendMarkedContentEnd(sb, markedContentId);
                        textMarkedContentOpen = false;
                    }

                    AppendMarkedContentBegin(sb, "Link", linkMarkedContentId);
                    content
                        .BeginText()
                        .TextLeading(defaultLeading)
                        .TextMatrix(lineXOrigin + xCursor, lineY)
                        .WordSpacing(pdfWordSpacing)
                        .Font(fontRes, runFontSize, syntheticOblique, syntheticBold: syntheticBold);
                    if (Math.Abs(textRise) > 0.0001) {
                        content.TextRise(textRise);
                    }

                    content
                        .FillColor(color ?? PdfColor.Black)
                        .ShowText(textCommand, runFontSize, textRise, suppressActualText);
                    ResetRichTextSpacing(content, s.HorizontalTextScaling, s.CharacterSpacing, pdfWordSpacing);
                    content.EndText();
                    AppendMarkedContentEnd(sb, linkMarkedContentId);
                    content
                        .BeginText()
                        .TextLeading(defaultLeading)
                        .TextMatrix(lineXOrigin + xCursor + wSeg, lineY)
                        .WordSpacing(pdfWordSpacing);
                    if (Math.Abs(textRise) > 0.0001) {
                        content.TextRise(0);
                    }

                    currentTextRise = 0;
                } else {
                    content.ShowText(textCommand, runFontSize, textRise, suppressActualText);
                    ResetRichTextSpacing(content, s.HorizontalTextScaling, s.CharacterSpacing, pdfWordSpacing);
                }

                double baselineY = lineY + textRise;

                if (s.Underline) {
                    var ulColor = (s.DecorationColor ?? s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = baselineY - runFontSize * 0.15;
                    underlines ??= new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color, OfficeIMO.Drawing.OfficeTextDecorationStyle Style)>();
                    if (s.UnderlineStyle == OfficeIMO.Drawing.OfficeTextDecorationStyle.Words) {
                        VisitRichWordDecorationAdvances(s, textCommand, opts,
                            (start, end) => underlines.Add((lineXOrigin + segmentStartX + start,
                                lineXOrigin + segmentStartX + end, yLine, ulColor, OfficeIMO.Drawing.OfficeTextDecorationStyle.Single)));
                    } else {
                        underlines.Add((lineXOrigin + segmentStartX, lineXOrigin + segmentStartX + wSeg, yLine, ulColor, s.UnderlineStyle));
                    }
                }
                if (s.Strike) {
                    var stColor = (s.DecorationColor ?? s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = baselineY + runFontSize * 0.32;
                    (strikes ??= new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color, OfficeIMO.Drawing.OfficeTextDecorationStyle Style)>())
                        .Add((lineXOrigin + segmentStartX, lineXOrigin + segmentStartX + wSeg, yLine, stColor, s.StrikeStyle));
                }
                if (hasLinkTarget) {
                    var fontForMetrics = s.Font;
                    double asc = GetAscenderForOptions(fontForMetrics, s.NamedFont, runFontSize, opts);
                    double desc = GetDescenderForOptions(fontForMetrics, s.NamedFont, runFontSize, opts);
                    double x1 = lineXOrigin + segmentStartX;
                    double x2 = x1 + wSeg;
                    double y1 = baselineY - desc;
                    double y2 = baselineY + asc;
                    AddRichTextLinkAnnotation(annots, structurePage, x1, y1, x2, y2, s.Uri, s.DestinationName, s.Contents, linkStructElementIndex);
                }
                xCursor += wSeg;
            }

            if (segs.Count > 0 && segs.Any(seg => seg.EndsWithTextSeparator)) {
                RichSeg last = segs[segs.Count - 1];
                if (last.InlineElement != null && fontPage != null)
                    MarkPageRichTextFont(fontPage, opts, last.Font, last.NamedFont);
                string separatorFontResource = GetFontResourceName(last.Font, last.NamedFont, ChooseNormal(opts.DefaultFont));
                double separatorFontSize = EffectiveRichFontSize(last.FontSize, last.Baseline);
                double separatorTextRise = TextRiseForBaseline(last.FontSize, last.Baseline) + last.LineAlignmentRise;
                content.Font(separatorFontResource, separatorFontSize, opts.NeedsSyntheticOblique(last.Font, last.NamedFont), syntheticBold: opts.NeedsSyntheticBold(last.Font, last.NamedFont));
                if (Math.Abs(separatorTextRise - currentTextRise) > 0.0001) {
                    content.TextRise(separatorTextRise);
                    currentTextRise = separatorTextRise;
                }

                content.ShowText(EncodeTextShowCommand(" ", last.Font, last.NamedFont, opts, last.FeatureSettings, fontMetricScale: last.FontMetricScale), separatorFontSize, separatorTextRise, suppressActualText);
            }

            if (Math.Abs(currentTextRise) > 0.0001) {
                content.TextRise(0);
            }

            yOffset += li < lineHeights.Count ? lineHeights[li] : defaultLeading;
        }
        content
            .WordSpacing(0)
            .EndText();
        if (textMarkedContentOpen) {
            AppendMarkedContentEnd(sb, markedContentId);
        }
        if (optionalBreakText) EndOptionalBreakText(sb, block, logicalRows, opts, xOrigin, startY, widthUsed, fontSize);
        if (foregroundGraphicsState != null) {
            new ContentStreamBuilder(sb).RestoreState();
        }

        if (foregroundGroup != null && sb.Length > foregroundStart) {
            string paint = sb.ToString(foregroundStart, sb.Length - foregroundStart);
            sb.Length = foregroundStart;
            foregroundGroup(() => sb.Append(paint));
        }

        bool hasDecorationState = decorationGraphicsState != null && (underlines != null || strikes != null);
        if (hasDecorationState) {
            new ContentStreamBuilder(sb).SaveState().GraphicsState(decorationGraphicsState!);
        }
        if (underlines != null) {
            foreach (var ul in underlines) {
                AppendArtifactBegin(sb, markedContentId.HasValue);
                AppendPageTextDecorationLine(sb, ul.X1, ul.X2, ul.Y, 0.5D, ul.Color, ul.Style);
                AppendArtifactEnd(sb, markedContentId.HasValue);
            }
        }

        if (strikes != null) {
            foreach (var st in strikes) {
                AppendArtifactBegin(sb, markedContentId.HasValue);
                AppendPageTextDecorationLine(sb, st.X1, st.X2, st.Y, 0.5D, st.Color, st.Style);
                AppendArtifactEnd(sb, markedContentId.HasValue);
            }
        }
        if (hasDecorationState) {
            new ContentStreamBuilder(sb).RestoreState();
        }

        AppendRichBarTabDecorations(sb, lines, lineHeights, opts, startY, fontSize, defaultLeading,
            xOrigin, firstLineXOverride, lineXOffsets, lineTopGaps, baselineFont,
            paragraphSpacingAfter, paragraphSpacingBefore, markedContentId.HasValue, decorationGraphicsState, barClipFrame);

        // A picture can overlap later text after the cell turns its text axes.
        // Preserve its logical structure position while painting the upright image above that text.
        if (uprightCellImages != null) sb.Append(uprightCellImages);
    }

    private static bool UsesSimpleFontText(RichSeg segment, PdfCreateOptions options) {
        if (segment.InlineElement != null) return false;
        if (segment.NamedFont is PdfNamedFontFace named &&
            (options.TryGetNamedFontProgram(named, out _) || options.TryGetNamedOpenTypeCffFontProgram(named, out _))) return false;
        return !options.TryGetEmbeddedStandardFontProgram(segment.Font, out _) &&
            !options.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out _);
    }

    private static int? FindStructElementIndex(LayoutResult.Page? structurePage, int? markedContentId, string? structureType) {
        if (structurePage == null || !markedContentId.HasValue || string.IsNullOrWhiteSpace(structureType)) {
            return null;
        }

        for (int i = 0; i < structurePage.StructElements.Count; i++) {
            PageStructElement element = structurePage.StructElements[i];
            if (element.MarkedContentId == markedContentId &&
                string.Equals(element.StructureType, structureType, StringComparison.Ordinal)) {
                return i;
            }
        }

        return null;
    }

    private static void EnsureTextMarkedContentOpen(
        StringBuilder sb,
        ref ContentStreamBuilder content,
        ref bool textMarkedContentOpen,
        LayoutResult.Page? structurePage,
        int? textStructElementIndex,
        string? structureType,
        double defaultLeading,
        double x,
        double y,
        double wordSpacing,
        string fontRes,
        bool syntheticOblique,
        bool syntheticBold,
        double fontSize,
        double textRise,
        PdfColor fillColor,
        bool isLinkLeadingWhitespace) {
        if (textMarkedContentOpen ||
            structurePage == null ||
            !textStructElementIndex.HasValue ||
            string.IsNullOrWhiteSpace(structureType)) {
            return;
        }

        if (textStructElementIndex.Value < 0 || textStructElementIndex.Value >= structurePage.StructElements.Count) {
            return;
        }

        content.EndText();
        int markedContentId = structurePage.NextMarkedContentId++;
        PageStructElement element = structurePage.StructElements[textStructElementIndex.Value];
        if (!element.MarkedContentId.HasValue) {
            structurePage.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = "Span",
                ParentElementIndex = textStructElementIndex,
                IsLinkLeadingWhitespace = isLinkLeadingWhitespace
            });
        } else {
            if (element.AdditionalMarkedContentIds == null) {
                element.AdditionalMarkedContentIds = new System.Collections.Generic.List<int>();
            }
            element.AdditionalMarkedContentIds.Add(markedContentId);
        }

        AppendMarkedContentBegin(sb, structureType, markedContentId);
        content = new ContentStreamBuilder(sb)
            .BeginText()
            .TextLeading(defaultLeading)
            .TextMatrix(x, y)
            .WordSpacing(wordSpacing)
            .Font(fontRes, fontSize, syntheticOblique, syntheticBold: syntheticBold)
            .FillColor(fillColor);
        if (Math.Abs(textRise) > 0.0001) {
            content.TextRise(textRise);
        }

        textMarkedContentOpen = true;
    }

    private static void AppendMarkedContentBegin(StringBuilder sb, string? structureType, int? markedContentId) {
        if (!markedContentId.HasValue || string.IsNullOrWhiteSpace(structureType)) {
            return;
        }

        sb.Append('/')
            .Append(structureType)
            .Append(" << /MCID ")
            .Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture))
            .Append(" >> BDC\n");
    }

    private static void AppendMarkedContentEnd(StringBuilder sb, int? markedContentId) {
        if (markedContentId.HasValue) {
            sb.Append("EMC\n");
        }
    }

    private static void AddRichTextLinkAnnotation(System.Collections.Generic.List<LinkAnnotation> annots, LayoutResult.Page? structurePage, double x1, double y1, double x2, double y2, string? uri, string? destinationName, string? contents, int? structElementIndex) {
        // Optional breaks and other non-painting runs retain their logical text,
        // but have no clickable area. Do not serialize or merge empty rectangles.
        if (x2 <= x1 || y2 <= y1) { return; }
        if (annots.Count > 0) {
            LinkAnnotation previous = annots[annots.Count - 1];
            double gap = x1 - previous.X2;
            bool sameTarget =
                string.Equals(previous.Uri, uri, System.StringComparison.Ordinal) &&
                string.Equals(previous.DestinationName, destinationName, System.StringComparison.Ordinal) &&
                string.Equals(previous.Contents, contents, System.StringComparison.Ordinal);
            bool sameLine =
                Math.Abs(previous.Y1 - y1) <= 0.5D &&
                Math.Abs(previous.Y2 - y2) <= 0.5D;
            if (sameTarget && sameLine && gap >= -0.25D && gap <= 18D) {
                if (structElementIndex.HasValue && previous.StructElementIndex.HasValue && structurePage != null) {
                    if (!TryMergeLinkStructureElements(structurePage, previous.StructElementIndex.Value, structElementIndex.Value)) {
                        annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = uri, DestinationName = destinationName, Contents = contents, StructElementIndex = structElementIndex });
                        return;
                    }
                } else if (structElementIndex.HasValue || previous.StructElementIndex.HasValue) {
                    annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = uri, DestinationName = destinationName, Contents = contents, StructElementIndex = structElementIndex });
                    return;
                }

                annots[annots.Count - 1] = new LinkAnnotation {
                    X1 = previous.X1,
                    Y1 = Math.Min(previous.Y1, y1),
                    X2 = Math.Max(previous.X2, x2),
                    Y2 = Math.Max(previous.Y2, y2),
                    Uri = uri,
                    DestinationName = destinationName,
                    Contents = contents,
                    StructElementIndex = previous.StructElementIndex
                };
                return;
            }
        }

        annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = uri, DestinationName = destinationName, Contents = contents, StructElementIndex = structElementIndex });
    }

    private static bool TryMergeLinkStructureElements(LayoutResult.Page structurePage, int targetStructElementIndex, int mergedStructElementIndex) {
        if (targetStructElementIndex < 0 || targetStructElementIndex >= structurePage.StructElements.Count ||
            mergedStructElementIndex < 0 || mergedStructElementIndex >= structurePage.StructElements.Count ||
            mergedStructElementIndex > targetStructElementIndex + 2 ||
            mergedStructElementIndex <= targetStructElementIndex ||
            mergedStructElementIndex != structurePage.StructElements.Count - 1) {
            return false;
        }

        PageStructElement target = structurePage.StructElements[targetStructElementIndex];
        PageStructElement merged = structurePage.StructElements[mergedStructElementIndex];
        // Geometric proximity does not make links logically adjacent. Moving a later
        // MCID across a Span or into another paragraph would reorder tagged content.
        if (target.ParentElementIndex != merged.ParentElementIndex) {
            return false;
        }
        PageStructElement? whitespace = null;
        if (mergedStructElementIndex == targetStructElementIndex + 2) {
            whitespace = structurePage.StructElements[targetStructElementIndex + 1];
            if (!whitespace.IsLinkLeadingWhitespace || !whitespace.MarkedContentId.HasValue ||
                whitespace.ParentElementIndex != target.ParentElementIndex) {
                return false;
            }
        }
        if (merged.MarkedContentId.HasValue) {
            if (target.AdditionalMarkedContentIds == null) {
                target.AdditionalMarkedContentIds = new System.Collections.Generic.List<int>();
            }

            // A word's own leading space belongs between the two word MCIDs. It
            // can join that Link, but unrelated text must keep its separate owner.
            if (whitespace != null) target.AdditionalMarkedContentIds.Add(whitespace.MarkedContentId!.Value);
            target.AdditionalMarkedContentIds.Add(merged.MarkedContentId.Value);
        }

        structurePage.StructElements.RemoveAt(mergedStructElementIndex);
        if (whitespace != null) structurePage.StructElements.RemoveAt(targetStructElementIndex + 1);
        return true;
    }

}
