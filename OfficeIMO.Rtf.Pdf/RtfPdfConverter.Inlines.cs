using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void RenderParagraph(RtfDocument document, RtfParagraph paragraph, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        if (RtfPdfMapping.HasPageBreakBefore(document, paragraph)) {
            pdf.PageBreak();
        }

        PdfCore.PdfAlign align = RtfPdfMapping.ToPdfAlign(paragraph.Alignment);
        PdfCore.PdfParagraphStyle? style = RtfPdfMapping.ToPdfParagraphStyle(document, paragraph);
        List<PdfCore.TextRun> pendingRuns = new List<PdfCore.TextRun>();
        bool emitted = false;
        AppendListMarker(paragraph, pendingRuns, state);

        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, run, pendingRuns, options, state);
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Page || rtfBreak.Kind == RtfBreakKind.SoftPage:
                    FlushParagraph(pdf, pendingRuns, align, style);
                    emitted = true;
                    pdf.PageBreak();
                    break;
                case RtfBreak:
                    pendingRuns.Add(PdfCore.TextRun.LineBreak());
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
                    FlushParagraph(pdf, pendingRuns, align, style);
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
                    FlushParagraph(pdf, pendingRuns, align, style);
                    emitted = true;
                    pdf.Bookmark(marker.Name);
                    break;
            }
        }

        if (pendingRuns.Count > 0 || !emitted) {
            FlushParagraph(pdf, pendingRuns, align, style);
        }
    }

    private static void FlushParagraph(PdfCore.PdfDocument pdf, List<PdfCore.TextRun> runs, PdfCore.PdfAlign align, PdfCore.PdfParagraphStyle? style) {
        List<PdfCore.TextRun> snapshot = runs.Count == 0
            ? new List<PdfCore.TextRun> { PdfCore.TextRun.Normal(string.Empty) }
            : new List<PdfCore.TextRun>(runs);
        runs.Clear();
        pdf.Paragraph(paragraph => paragraph.Runs(snapshot), align, style: style);
    }

    private static void AppendParagraphRuns(RtfDocument document, RtfParagraph paragraph, List<PdfCore.TextRun> runs, RtfPdfSaveOptions options, PdfRenderState state, bool collectNotes = true, string? inheritedLinkUri = null, string? inheritedLinkDestinationName = null, string? inheritedLinkContents = null) {
        AppendListMarker(paragraph, runs, state);
        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, run, runs, options, state, collectNotes, inheritedLinkUri, inheritedLinkDestinationName, inheritedLinkContents);
                    break;
                case RtfBreak:
                    runs.Add(PdfCore.TextRun.LineBreak());
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
                        field.HyperlinkField?.ScreenTip ?? inheritedLinkContents);
                    break;
                case RtfGeneratedText generatedText:
                    AppendGeneratedText(generatedText, runs, state, collectNotes);
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

    private static void AppendRun(RtfDocument document, RtfRun run, List<PdfCore.TextRun> runs, RtfPdfSaveOptions options, PdfRenderState state, bool collectNotes = true, string? inheritedLinkUri = null, string? inheritedLinkDestinationName = null, string? inheritedLinkContents = null) {
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

        string text = run.Text ?? string.Empty;
        if (collectNotes && run.Note != null) {
            state.AddNote(run.Note, text);
        }

        if (text.Length == 0) {
            return;
        }

        PdfCore.PdfColor? foreground = RtfPdfMapping.ToPdfColor(document, run.ForegroundColorIndex);
        PdfCore.PdfColor? background = RtfPdfMapping.ToPdfColor(document, run.HighlightColorIndex)
            ?? RtfPdfMapping.ToPdfColor(document, run.CharacterBackgroundColorIndex);
        PdfCore.PdfStandardFont? font = state.ResolveFont(run.FontId, run.Bold, run.Italic);
        string? linkUri = run.Hyperlink?.ToString() ?? inheritedLinkUri;
        string? linkDestinationName = run.Hyperlink == null && linkUri == null ? inheritedLinkDestinationName : null;
        string? linkContents = (linkUri != null || linkDestinationName != null) && run.Hyperlink == null ? inheritedLinkContents : null;

        runs.Add(new PdfCore.TextRun(
            text,
            bold: run.Bold,
            underline: run.Underline,
            color: foreground,
            italic: run.Italic,
            strike: run.Strike || run.DoubleStrike,
            fontSize: run.FontSize,
            font: font,
            linkUri: linkUri,
            linkContents: linkContents,
            baseline: RtfPdfMapping.ToPdfBaseline(run.VerticalPosition),
            linkDestinationName: linkDestinationName,
            backgroundColor: background));
    }

    private static string? GetFieldLinkDestinationName(RtfField field) {
        if (field.Hyperlink != null || string.IsNullOrWhiteSpace(field.HyperlinkField?.SubAddress)) {
            return null;
        }

        return field.HyperlinkField!.SubAddress;
    }

    private static void AppendPlainText(string text, List<PdfCore.TextRun> runs) {
        if (!string.IsNullOrEmpty(text)) {
            runs.Add(PdfCore.TextRun.Normal(text));
        }
    }

    private static void AppendGeneratedText(RtfGeneratedText generatedText, List<PdfCore.TextRun> runs, PdfRenderState state, bool collectNotes = true) {
        string text = generatedText.ToPlainText();
        if (collectNotes && generatedText.Note != null) {
            state.AddNote(generatedText.Note, text);
        }

        AppendPlainText(text, runs);
    }

    private static void AppendListMarker(RtfParagraph paragraph, List<PdfCore.TextRun> runs, PdfRenderState state) {
        string? marker = GetListMarker(paragraph, state);
        if (marker != null && marker.Length > 0) {
            runs.Add(PdfCore.TextRun.Normal(marker));
        }
    }

    private static string? GetListMarker(RtfParagraph paragraph, PdfRenderState state) {
        if (paragraph.ListKind == RtfListKind.None) {
            return null;
        }

        if (paragraph.ListText != null) {
            string markerText = NormalizeListMarkerText(paragraph.ListText.ToPlainText());
            if (paragraph.ListKind == RtfListKind.Decimal) {
                state.AdvanceDecimalList(paragraph, markerText);
            }

            return EnsureMarkerSeparator(markerText);
        }

        if (paragraph.ListKind == RtfListKind.Bullet) {
            return "\u2022 ";
        }

        return state.NextDecimalMarker(paragraph).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". ";
    }

    private static string NormalizeListMarkerText(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        return text
            .Replace("\r\n", " ")
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('\f', ' ')
            .Replace('\v', ' ')
            .Replace('\t', ' ')
            .Trim();
    }

    private static string EnsureMarkerSeparator(string marker) {
        if (marker.Length == 0 || char.IsWhiteSpace(marker[marker.Length - 1])) {
            return marker;
        }

        return marker + " ";
    }

    private static bool TryReadLeadingIntegerMarker(string marker, out int value) {
        value = 0;
        int index = 0;
        while (index < marker.Length && char.IsWhiteSpace(marker[index])) {
            index++;
        }

        int start = index;
        while (index < marker.Length && char.IsDigit(marker[index])) {
            int digit = marker[index] - '0';
            if (value > (int.MaxValue - digit) / 10) {
                value = 0;
                return false;
            }

            value = (value * 10) + digit;
            index++;
        }

        return index > start;
    }
}
