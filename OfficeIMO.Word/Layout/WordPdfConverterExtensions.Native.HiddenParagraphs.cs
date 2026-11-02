using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // A hidden paragraph mark removes the boundary, while each source run retains
    // its own paragraph/style context. Copying runs into the first XML paragraph
    // would change inherited fonts, visibility and hyperlink formatting.
    private static bool TryRenderNativeJoinedParagraphs(
        INativePdfFlow pdf, IReadOnlyList<WordElement> elements, ref int index, ref int rejectedJoinUntil,
        Func<WordParagraph, (int Level, string Marker)?> getMarker,
        NativeNoteNumbering footnoteNumbersById, WordToPdfOptions? options,
        NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
        if (index <= rejectedJoinUntil ||
            elements[index] is not WordParagraph candidate ||
            ShouldRenderNativeEmptyParagraphLineBox(candidate) ||
            index + 1 >= elements.Count || elements[index + 1] is not WordParagraph) return false;
        if (!CanJoinNativeTextParagraph(candidate, getMarker, nativeDefaults, nativeFontMap)) {
            ReportUnsupportedJoin();
            return false;
        }
        if (!TryGetNativeParagraphJoinChain(elements, index, out var paragraphs, out int lastIndex)) return false;
        WordParagraph first = paragraphs[0];
        int unsupportedOffset = FirstUnsupportedNativeJoinIndex(paragraphs, getMarker, nativeDefaults, nativeFontMap);
        if (unsupportedOffset >= 0) {
            rejectedJoinUntil = index + unsupportedOffset - 1;
            ReportUnsupportedJoin();
            return false;
        }

        void ReportUnsupportedJoin() {
            if (options != null) AddNativeExportWarning(options, "NativeHiddenParagraphJoinUnsupported",
                nativeDefaults.RunningContentContext == null ? "body" : "header/footer",
                "Hidden paragraph marks beside lists, headings, decorated paragraphs, flow breaks or objects retain separate PDF paragraph boundaries.");
        }

        WordParagraph last = paragraphs[paragraphs.Count - 1];
        PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(first, nativeDefaults, nativeFontMap);
        PdfCore.PdfParagraphStyle finalStyle = CreateNativeParagraphStyle(last, nativeDefaults, nativeFontMap);
        if (ShouldSuppressNativeContextualSpacingBefore(first)) style.SpacingBefore = 0D;
        // Seed rich line measurement with the smallest visible source font, just
        // as an ordinary rich paragraph does. Hidden marks and preceding large
        // runs must not enlarge later wrapped or explicitly broken lines.
        List<WordParagraph> visibleParagraphs = paragraphs.Where(paragraph => GetNativeRunningContentRuns(paragraph, nativeDefaults.RunningContentContext).Any(run =>
            IsNativeRenderableTextRun(run, paragraph))).ToList();
        if (visibleParagraphs.Count > 0) {
            style.FontSize = visibleParagraphs.Min(paragraph => ResolveNativeParagraphLayoutFontSize(paragraph,
                nativeDefaults, GetNativeParagraphStyleDefaults(paragraph)));
            double naturalLineHeight = visibleParagraphs.Max(paragraph => ResolveNativeParagraphSingleLineHeight(paragraph,
                nativeDefaults, GetNativeParagraphStyleDefaults(paragraph), nativeFontMap: nativeFontMap));
            style.LineSpacing = ResolveNativeParagraphLineSpacing(first, GetNativeParagraphStyleDefaults(first),
                nativeDefaults).ToPdfLineSpacing(naturalLineHeight);
        }
        style.SpacingAfter = ShouldSuppressNativeContextualSpacingAfter(last,
            GetNextNativeRenderableElement(elements, lastIndex) as WordParagraph) ? 0D : finalStyle.SpacingAfter;
        style.KeepWithNext = finalStyle.KeepWithNext;
        if (HasNativePageBreakBefore(first)) pdf.PageBreak();
        foreach (WordParagraph paragraph in paragraphs) {
            RecordNativeBodyParagraphDiagnostics(paragraph, options, "joined body paragraph",
                mapsCheckBoxes: false, mapsFormFields: false, mapsPictureControls: false, mapsRepeatingSections: false);
            if (nativeDefaults.RunningContentContext == null && !string.IsNullOrEmpty(paragraph.Bookmark?.Name))
                pdf.Bookmark(paragraph.Bookmark!.Name!);
        }
        bool hasNoteReferences = paragraphs.Any(paragraph => GetNativeParagraphFootnoteNumbers(paragraph,
            GetNativeRunningContentRuns(paragraph, nativeDefaults.RunningContentContext), Array.Empty<int>(), footnoteNumbersById).Count > 0);
        if (!hasNoteReferences && !paragraphs.Any(paragraph => GetNativeRunningContentRuns(paragraph, nativeDefaults.RunningContentContext).Any(run =>
                IsNativeRenderableTextRun(run, paragraph) ||
                (IsNativeTextWrappingBreak(run) && !IsNativeHiddenTextRun(run, paragraph))))) {
            RenderNativeEmptyParagraph(pdf, last, style, nativeDefaults, nativeFontMap);
            if (style.SpacingAfter == 0D && pdf is NativeSpacingCollapseFlow emptyCollapse)
                emptyCollapse.RetainSuppressedSpacingAfter(finalStyle.SpacingAfter);
            index = lastIndex;
            return true;
        }
        pdf.Paragraph(builder => {
            foreach (WordParagraph paragraph in paragraphs) {
                List<WordParagraph> runs = GetNativeRunningContentRuns(paragraph, nativeDefaults.RunningContentContext);
                bool hasRuns = runs.Any(run => IsNativeRenderableTextRun(run, paragraph) ||
                    (IsNativeTextWrappingBreak(run) && !IsNativeHiddenTextRun(run, paragraph)));
                string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                if (nativeDefaults.RunningContentContext != null) content = string.Concat(runs.Select(run => run.Text));
                string renderContent = hasRuns || ShouldRenderNativeDirectText(paragraph, runs, content) ? content : string.Empty;
                AddNativeParagraphContent(builder, paragraph, null, runs, hasRuns, renderContent,
                    GetNativeParagraphFootnoteNumbers(paragraph, runs, Array.Empty<int>(), footnoteNumbersById),
                    footnoteNumbersById, options, nativeDefaults, nativeFontMap);
            }
        }, ResolveNativeParagraphAlign(first), ResolveNativeParagraphDefaultColor(first), style);
        if (style.SpacingAfter == 0D && pdf is NativeSpacingCollapseFlow collapseFlow)
            collapseFlow.RetainSuppressedSpacingAfter(finalStyle.SpacingAfter);
        index = lastIndex;
        return true;
    }

    private static bool TryGetNativeParagraphJoinChain(IReadOnlyList<WordElement> elements, int index,
        out List<WordParagraph> paragraphs, out int lastIndex) {
        paragraphs = new List<WordParagraph>();
        lastIndex = index;
        if (elements[index] is not WordParagraph first || ShouldRenderNativeEmptyParagraphLineBox(first) ||
            index + 1 >= elements.Count || elements[index + 1] is not WordParagraph) return false;
        paragraphs.Add(first);
        while (!ShouldRenderNativeEmptyParagraphLineBox(paragraphs[paragraphs.Count - 1]) &&
               lastIndex + 1 < elements.Count && elements[lastIndex + 1] is WordParagraph next) {
            paragraphs.Add(next);
            lastIndex++;
        }
        return true;
    }

    private static bool CanJoinNativeParagraphChain(IReadOnlyList<WordParagraph> paragraphs,
        Func<WordParagraph, (int Level, string Marker)?> getMarker,
        NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) =>
        FirstUnsupportedNativeJoinIndex(paragraphs, getMarker, nativeDefaults, nativeFontMap) < 0;

    private static int FirstUnsupportedNativeJoinIndex(IReadOnlyList<WordParagraph> paragraphs,
        Func<WordParagraph, (int Level, string Marker)?> getMarker,
        NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
        for (int offset = 0; offset < paragraphs.Count; offset++) {
            if ((offset > 0 && HasNativePageBreakBefore(paragraphs[offset])) ||
                !CanJoinNativeTextParagraph(paragraphs[offset], getMarker, nativeDefaults, nativeFontMap)) return offset;
        }
        return -1;
    }

    private static bool CanJoinNativeTextParagraph(WordParagraph paragraph,
        Func<WordParagraph, (int Level, string Marker)?> getMarker,
        NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
        W.Paragraph source = paragraph._paragraph;
        if (WordParagraph.IsSectionMarkOnly(source) || source.ParagraphProperties?.SectionProperties != null ||
            getMarker(paragraph) != null || GetHeadingLevel(paragraph) > 0 || HasNativeParagraphBarTabs(paragraph) ||
            source.ParagraphProperties?.GetFirstChild<W.ParagraphBorders>() != null ||
            source.ParagraphProperties?.GetFirstChild<W.Shading>() != null ||
            source.Descendants<W.Drawing>().Any() || source.Descendants<W.Picture>().Any() ||
            source.Descendants<W.EmbeddedObject>().Any() ||
            WordEquation.GetOccurrences(paragraph._document, source).Count > 0 ||
            GetNativeCheckBoxControls(paragraph).Count > 0 || GetNativeFormFieldControls(paragraph).Count > 0 ||
            GetNativeRepeatingSectionControls(paragraph).Count > 0) return false;
        PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph, nativeDefaults, nativeFontMap);
        if (CreateNativeParagraphPanelStyle(paragraph, style) != null ||
            CreateNativeTopBorderRuleStyle(paragraph, style) != null ||
            CreateNativeBottomBorderRuleStyle(paragraph, style) != null) return false;
        foreach (WordParagraph run in GetNativeRunningContentRuns(paragraph, nativeDefaults.RunningContentContext)) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> children = run._visibleRunSourceChildren ?? run._run!.ChildElements;
            if (children.OfType<W.Break>().Any(boundary => boundary.Type?.Value == W.BreakValues.Page ||
                boundary.Type?.Value == W.BreakValues.Column)) return false;
        }
        return true;
    }
}
