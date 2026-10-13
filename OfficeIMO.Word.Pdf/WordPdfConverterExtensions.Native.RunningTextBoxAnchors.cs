using DocumentFormat.OpenXml;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Keeps simple non-intersecting page-anchored running-story boxes out of paragraph
    /// flow. Unsupported peers, nested boxes and intersecting wrap remain on the
    /// explicitly approximated mixed-content renderer.
    /// </summary>
    private static bool TryRenderNativeAnchoredRunningTextBoxParagraph(INativePdfFlow pdf, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, (int Level, string Marker)? marker, NativeNoteNumbering notes,
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap, WordParagraph? nextParagraph) {
        if (defaults.RunningContentContext == null || !TryPrepareNativeAnchoredRunningTextBoxParagraph(pdf.PageSize,
                paragraph, runs, marker, notes, options, defaults, fontMap, out var prepared)) return false;
        WordParagraph[] selected = prepared!.Runs;
        string text = string.Concat(selected.Select(run => run.Text));
        if (ShouldSuppressNativeContextualSpacingAfter(paragraph, nextParagraph)) prepared.Style.SpacingAfter = 0D;
        if (ShouldSuppressNativeContextualSpacingBefore(paragraph)) prepared.Style.SpacingBefore = 0D;
        pdf.Paragraph(builder => AddNativeParagraphContent(builder, paragraph, null, selected, true, text,
            Array.Empty<int>(), notes, options, defaults, fontMap, projectedContentOnly: true),
            PdfCore.PdfAlign.Left, style: prepared.Style);
        return true;
    }

    private sealed record NativeAnchoredRunningTextBoxParagraph(WordParagraph[] Runs, PdfCore.PdfParagraphStyle Style);

    private static bool TryPrepareNativeAnchoredRunningTextBoxParagraph(PdfCore.PageSize pageSize, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, (int Level, string Marker)? marker, NativeNoteNumbering? notes,
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap,
        out NativeAnchoredRunningTextBoxParagraph? prepared) {
        prepared = null;
        if (defaults.RunningContentSection == null || marker != null || WordDocumentTraversal.GetListInfo(paragraph).HasValue ||
            WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0 ||
            ResolveNativeParagraphAlign(paragraph, allowJustify: false) != PdfCore.PdfAlign.Left) return false;
        options?.CancellationToken.ThrowIfCancellationRequested();
        List<WordTextBox> boxes = runs.SelectMany(run => run.GetTextBoxes()).ToList();
        if (boxes.Count == 0) return false;

        var surrounding = new List<WordParagraph>();
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
                if (child is W.RunProperties) continue;
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                if (view.GetTextBoxes().Any()) continue;
                if (child is not W.Text) return false;
                surrounding.Add(view);
            }
        }
        if (!surrounding.Any(run => IsNativeRenderableTextRun(run, paragraph))) return false;
        PdfCore.PdfParagraphStyle paragraphStyle = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
        PdfCore.PageMargins margins = GetNativeMargins(defaults.RunningContentSection, options, (0D, 0D));
        double textLeft = margins.Left + paragraphStyle.LeftIndent + paragraphStyle.FirstLineIndent;
        double textWidth = 0D;
        foreach (WordParagraph run in surrounding) {
            double? measured = fontMap.MeasureText(CreateNativeCellTextRun(run.Text, run,
                NativeTableStyleDefaults.Empty, defaults, fontMap));
            if (!measured.HasValue) return false;
            textWidth += measured.Value;
        }
        double contentWidth = defaults.RunningContentContext?.ContentWidth ?? pageSize.Width - margins.Left - margins.Right;
        if (textLeft + textWidth > margins.Left + contentWidth - paragraphStyle.RightIndent) return false;
        var canvas = new PdfCore.PdfPageCanvas();
        foreach (WordTextBox box in boxes.OrderBy(box => box.ZOrder)) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (!TryAddNativeRunningTextBox(canvas, box, pageSize, textLeft, textWidth,
                    notes, options, defaults, fontMap)) return false;
        }
        paragraphStyle.AnchoredCanvas = new PdfCore.PdfCanvasBlock(canvas.Items);
        prepared = new NativeAnchoredRunningTextBoxParagraph(surrounding.ToArray(), paragraphStyle);
        return true;
    }
}
