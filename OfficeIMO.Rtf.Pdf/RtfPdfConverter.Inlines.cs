using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void RenderParagraph(RtfDocument document, RtfParagraph paragraph, PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state) {
        RecordParagraphFrameDiagnostic(document, paragraph, options, "Paragraph/Frame");
        RecordParagraphBarTabDiagnostic(document, paragraph, options);
        if (RtfPdfMapping.HasPageBreakBefore(document, paragraph)) {
            pdf.PageBreak();
        }

        PdfCore.PdfAlign align = RtfPdfMapping.ToPdfAlign(document.GetParagraphFormatting(paragraph).Alignment);
        List<PdfCore.PdfTextRun> pendingRuns = new List<PdfCore.PdfTextRun>();
        bool emitted = false;
        AppendListMarker(paragraph, pendingRuns, state, options);

        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, document.Styles.Count > 0 || run.UseDefaultCharacterFormatting
                        ? document.GetRunFormatting(paragraph, run) : run, pendingRuns, options, state);
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Page || rtfBreak.Kind == RtfBreakKind.SoftPage:
                    FlushParagraph(document, paragraph, pdf, pendingRuns, align, state);
                    emitted = true;
                    pdf.PageBreak();
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Column:
                    FlushParagraph(document, paragraph, pdf, pendingRuns, align, state);
                    emitted = true;
                    if (state.InColumns) pdf.AddColumnBreak();
                    else pdf.PageBreak();
                    break;
                case RtfBreak:
                    pendingRuns.Add(PdfCore.PdfTextRun.LineBreak());
                    break;
                case RtfField field:
                    AppendParagraphRuns(
                        document,
                        field.Result,
                        pendingRuns,
                        options,
                        state,
                        inheritedLinkUri: field.Hyperlink?.ToString(),
                        inheritedLinkDestinationName: GetFieldLinkDestinationName(field),
                        inheritedLinkContents: field.HyperlinkField?.ScreenTip);
                    break;
                case RtfGeneratedText generatedText:
                    AppendGeneratedText(generatedText, pendingRuns, state);
                    break;
                case RtfImage image:
                    FlushParagraph(document, paragraph, pdf, pendingRuns, align, state);
                    emitted = true;
                    RenderImage(image, pdf, options);
                    break;
                case RtfObject rtfObject:
                    AddConversionWarning(options, "ObjectFlattened", "Paragraph/Object", "RTF object was flattened to its visible text result.", RtfConversionAction.Flattened);
                    AppendPlainText(rtfObject.ToPlainText(), pendingRuns);
                    break;
                case RtfShape shape:
                    AddConversionWarning(options, "ShapeFlattened", "Paragraph/Shape", "RTF shape was flattened to its visible text result.", RtfConversionAction.Flattened);
                    AppendPlainText(shape.ToPlainText(), pendingRuns);
                    break;
                case RtfBookmarkMarker marker when marker.Kind == RtfBookmarkMarkerKind.Start:
                    FlushParagraph(document, paragraph, pdf, pendingRuns, align, state);
                    emitted = true;
                    pdf.Bookmark(marker.Name);
                    break;
            }
        }

        if (pendingRuns.Count > 0 || !emitted) {
            FlushParagraph(document, paragraph, pdf, pendingRuns, align, state);
        }
    }

    private static void FlushParagraph(RtfDocument document, RtfParagraph source, PdfCore.PdfDocument pdf,
        List<PdfCore.PdfTextRun> runs, PdfCore.PdfAlign align, PdfRenderState state) {
        List<PdfCore.PdfTextRun> snapshot = new List<PdfCore.PdfTextRun>(runs);
        EnsureParagraphLineBoxStrut(document, source, snapshot, state);
        runs.Clear();
        PdfCore.PdfParagraphStyle style = RtfPdfMapping.ToPdfParagraphStyle(document, source, snapshot, state.FontOptions);
        pdf.Paragraph(paragraph => paragraph.Runs(snapshot), align, style: style);
    }

    /// <summary>Keeps empty and control-only source lines' available character boxes without introducing visible text.</summary>
    private static void EnsureParagraphLineBoxStrut(RtfDocument document, RtfParagraph source,
        List<PdfCore.PdfTextRun> runs, PdfRenderState state, int bodyStartIndex = 0) {
        PdfCore.PdfTextRun? controlStyle = null;
        for (int index = bodyStartIndex; index < runs.Count; index++) {
            PdfCore.PdfTextRun run = runs[index];
            if (run.IsLineBoxStrut || run.InlineElement != null ||
                run.Text.Any(character => character is not ('\r' or '\n' or '\t'))) return;
            if (run.Text.Length > 0 && (run.FontSize?.ToPoints()) is double size && size > 0D &&
                !double.IsNaN(size) && !double.IsInfinity(size)) controlStyle = run;
        }

        PdfCore.PdfTextRun strut;
        if (controlStyle != null) {
            // Tabs retain their source character formatting. Keep that known
            // box even when wrapping discards their unpainted trailing advance.
            strut = new PdfCore.PdfTextRun(string.Empty, bold: controlStyle.Bold, italic: controlStyle.Italic,
                fontSize: controlStyle.FontSize, font: controlStyle.Font, fontFamily: controlStyle.FontFamily,
                baseline: controlStyle.Baseline);
        } else {
            // RtfBreak carries no character state. Use only the source defaults
            // or inherited style the existing semantic model can resolve.
            RtfRun inherited = document.GetRunFormatting(source, new RtfRun(string.Empty) {
                UseDefaultCharacterFormatting = source.UseParsedParagraphDefaults
            });
            strut = new PdfCore.PdfTextRun(string.Empty, bold: inherited.Bold, italic: inherited.Italic,
                fontSize: ((double?)(inherited.FontSize) is double physicalLength1 ? global::OfficeIMO.OfficeLength.Points(physicalLength1) : (global::OfficeIMO.OfficeLength?)null), font: state.ResolveFont(inherited.FontId, inherited.Bold, inherited.Italic));
        }
        strut.IsLineBoxStrut = true;
        if ((strut.FontSize?.ToPoints()) is double sourceSize && sourceSize > 0D &&
            !double.IsNaN(sourceSize) && !double.IsInfinity(sourceSize)) {
            for (int index = bodyStartIndex; index < runs.Count; index++) {
                PdfCore.PdfTextRun control = runs[index];
                if ((control.FontSize?.ToPoints()).HasValue || control.Text.Length == 0) continue;
                // Every blank source line needs its own known font box. A note
                // label before the body must not become a break's fallback size.
                runs[index] = new PdfCore.PdfTextRun(control.Text, bold: strut.Bold,
                    underline: control.Underline, color: control.Color, italic: strut.Italic,
                    strike: control.Strike, fontSize: global::OfficeIMO.OfficeLength.Points(sourceSize), font: strut.Font,
                    linkUri: control.LinkUri, linkContents: control.LinkContents,
                    baseline: strut.Baseline, linkDestinationName: control.LinkDestinationName,
                    tabLeader: control.TabLeader, tabAlignment: control.TabAlignment,
                    backgroundColor: control.BackgroundColor, fontFamily: strut.FontFamily,
                    underlineStyle: control.UnderlineStyle, strikeStyle: control.StrikeStyle,
                    decorationColor: control.DecorationColor).WithSpacingFrom(control);
            }
        }
        runs.Add(strut);
    }

    private enum PdfInlineProjectionContext { Text, RunningStory }

    private static void AppendParagraphRuns(RtfDocument document, RtfParagraph paragraph, List<PdfCore.PdfTextRun> runs, RtfToPdfOptions options, PdfRenderState state, bool collectNotes = true, string? inheritedLinkUri = null, string? inheritedLinkDestinationName = null, string? inheritedLinkContents = null, PdfInlineProjectionContext context = PdfInlineProjectionContext.Text) {
        RecordParagraphFrameDiagnostic(document, paragraph, options, "Paragraph/Frame");
        AppendListMarker(paragraph, runs, state, options);
        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, document.Styles.Count > 0 || run.UseDefaultCharacterFormatting
                        ? document.GetRunFormatting(paragraph, run) : run, runs, options, state, collectNotes,
                        inheritedLinkUri, inheritedLinkDestinationName, inheritedLinkContents);
                    break;
                case RtfBreak:
                    runs.Add(PdfCore.PdfTextRun.LineBreak());
                    break;
                case RtfField field:
                    string? fieldLinkUri = field.Hyperlink?.ToString();
                    AppendParagraphRuns(
                        document,
                        field.Result,
                        runs,
                        options,
                        state,
                        collectNotes,
                        fieldLinkUri ?? inheritedLinkUri,
                        fieldLinkUri == null ? GetFieldLinkDestinationName(field) ?? inheritedLinkDestinationName : null,
                        field.HyperlinkField?.ScreenTip ?? inheritedLinkContents,
                        context);
                    break;
                case RtfGeneratedText generatedText:
                    AppendGeneratedText(generatedText, runs, state, collectNotes);
                    break;
                case RtfImage image when context == PdfInlineProjectionContext.RunningStory:
                    AppendInlineImage(image, runs, options, "HeaderFooter/Paragraph/Image");
                    break;
                case RtfObject rtfObject:
                    AddConversionWarning(options, "ObjectFlattened", "Inline/Object", "RTF object was flattened to its visible text result.", RtfConversionAction.Flattened);
                    AppendPlainText(rtfObject.ToPlainText(), runs);
                    break;
                case RtfShape shape:
                    AddConversionWarning(options, "ShapeFlattened", "Inline/Shape", "RTF shape was flattened to its visible text result.", RtfConversionAction.Flattened);
                    AppendPlainText(shape.ToPlainText(), runs);
                    break;
            }
        }
    }

    private static void AppendRun(RtfDocument document, RtfRun run, List<PdfCore.PdfTextRun> runs, RtfToPdfOptions options, PdfRenderState state, bool collectNotes = true, string? inheritedLinkUri = null, string? inheritedLinkDestinationName = null, string? inheritedLinkContents = null) {
        if (run.Hidden && !options.IncludeHiddenText) {
            AddConversionWarning(
                options,
                "HiddenTextSkipped",
                "Run",
                "An RTF hidden text run was skipped because IncludeHiddenText is false.",
                new Dictionary<string, string> {
                    ["Length"] = (run.Text ?? string.Empty).Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
                });
            return;
        }

        ReportUnsupportedRunFormatting(run, options);

        string text = run.Text ?? string.Empty;
        if (collectNotes && run.Note != null) {
            state.AddNote(run.Note, text);
        }

        if (text.Length == 0) {
            return;
        }

        global::OfficeIMO.OfficeColor? foreground = RtfPdfMapping.ToPdfColor(document, run.ForegroundColorIndex);
        global::OfficeIMO.OfficeColor? background = RtfPdfMapping.ToPdfColor(document, run.HighlightColorIndex)
            ?? RtfPdfMapping.ToPdfColor(document, run.CharacterBackgroundColorIndex);
        PdfCore.PdfStandardFont? font = state.ResolveFont(run.FontId, run.Bold, run.Italic);
        string? linkUri = run.Hyperlink?.ToString() ?? inheritedLinkUri;
        string? linkDestinationName = run.Hyperlink == null && linkUri == null ? inheritedLinkDestinationName : null;
        string? linkContents = (linkUri != null || linkDestinationName != null) && run.Hyperlink == null ? inheritedLinkContents : null;

        var pdfRun = new PdfCore.PdfTextRun(
            text,
            bold: run.Bold,
            underline: run.Underline,
            color: foreground,
            italic: run.Italic,
            strike: run.Strike || run.DoubleStrike,
            fontSize: ((double?)(run.FontSize) is double physicalLength3 ? global::OfficeIMO.OfficeLength.Points(physicalLength3) : (global::OfficeIMO.OfficeLength?)null),
            font: font,
            linkUri: linkUri,
            linkContents: linkContents,
            baseline: RtfPdfMapping.ToPdfBaseline(run.VerticalPosition),
            linkDestinationName: linkDestinationName,
            backgroundColor: background);
        if (run.KerningHalfPoints is > 0) {
            double fontSize = run.FontSize ?? state.FontOptions.DefaultFontSize.ToPoints();
            pdfRun = pdfRun.WithFeatureSettings(state.FontOptions.DefaultTextFeatureSettings.With("kern",
                fontSize >= run.KerningHalfPoints.Value / 2D ? 1 : 0));
        }
        runs.Add(pdfRun);
    }

    private static void ReportUnsupportedRunFormatting(RtfRun run, RtfToPdfOptions options) {
        var features = new List<string>();
        if (run.Hidden && options.IncludeHiddenText) features.Add("hidden-visibility");
        if ((run.UnderlineStyle != RtfUnderlineStyle.None && run.UnderlineStyle != RtfUnderlineStyle.Single)
            || run.UnderlineColorIndex.HasValue) features.Add("underline-style-or-color");
        if (run.DoubleStrike) features.Add("double-strike");
        if (run.Outline || run.Shadow || run.Emboss || run.Imprint || run.CapsStyle != RtfCapsStyle.None) {
            features.Add("text-effects");
        }
        if (run.HighlightColorIndex.HasValue && run.CharacterBackgroundColorIndex.HasValue) {
            features.Add("overlapping-background-colors");
        }
        if (run.CharacterShadingForegroundColorIndex.HasValue || run.CharacterShadingPatternPercent.HasValue
            || run.CharacterShadingPattern != RtfShadingPattern.None) features.Add("character-shading");
        if (run.CharacterBorder.HasAnyValue) features.Add("character-border");
        if (run.CharacterSpacingTwips.HasValue || run.CharacterScalePercent.HasValue
            || run.CharacterOffsetHalfPoints.HasValue) features.Add("character-metrics");
        if (run.StyleId.HasValue || run.Direction.HasValue || run.LanguageId.HasValue) {
            features.Add("style-direction-or-language");
        }
        if (run.RevisionKind != RtfRevisionKind.None || run.RevisionAuthorIndex.HasValue
            || run.RevisionTimestampValue.HasValue || run.CharacterRevisionSaveId.HasValue
            || run.InsertionRevisionSaveId.HasValue || run.DeletionRevisionSaveId.HasValue) {
            features.Add("revision-metadata");
        }
        if (features.Count == 0) return;
        AddConversionWarning(
            options,
            "RunFormattingFlattened",
            "Run",
            "RTF run formatting outside the PDF text-run subset was flattened.",
            RtfConversionAction.Flattened,
            new Dictionary<string, string> { ["Features"] = string.Join(",", features) });
    }

    private static string? GetFieldLinkDestinationName(RtfField field) {
        if (field.Hyperlink != null || string.IsNullOrWhiteSpace(field.HyperlinkField?.SubAddress)) {
            return null;
        }

        return field.HyperlinkField!.SubAddress;
    }

    private static void AppendPlainText(string text, List<PdfCore.PdfTextRun> runs) {
        if (!string.IsNullOrEmpty(text)) {
            runs.Add(PdfCore.PdfTextRun.Normal(text));
        }
    }

    private static void AppendGeneratedText(RtfGeneratedText generatedText, List<PdfCore.PdfTextRun> runs, PdfRenderState state, bool collectNotes = true) {
        string text = generatedText.ToPlainText();
        if (collectNotes && generatedText.Note != null) {
            state.AddNote(generatedText.Note, text);
        }

        AppendPlainText(text, runs);
    }

    private static void AppendListMarker(RtfParagraph paragraph, List<PdfCore.PdfTextRun> runs, PdfRenderState state, RtfToPdfOptions options) {
        RtfListMarker? marker = state.NextListMarker(paragraph);
        if (marker == null) return;
        if (!marker.IsNumberFormatSupported) AddConversionWarning(options, "ListNumberFormatFlattened", "Paragraph/List", "The requested numbering format was represented by decimal marker text.", RtfConversionAction.Flattened);
        if (marker.Formatting.Level.PictureIndex.HasValue) AddConversionWarning(options, "ListPictureFlattened", "Paragraph/List", "The picture bullet was represented by its text marker.", RtfConversionAction.Flattened);
        if (marker.Text.Length > 0) {
            runs.Add(PdfCore.PdfTextRun.Normal(marker.Text + marker.Separator));
        }
    }
}
