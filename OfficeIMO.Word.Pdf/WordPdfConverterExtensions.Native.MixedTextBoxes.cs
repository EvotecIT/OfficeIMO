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
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap, int depth) {
        RecordNativeMixedTextBoxDiagnostic(runs, options,
            defaults.RunningContentContext == null ? "body paragraph" : "header/footer paragraph");
        var pending = new List<WordParagraph>();
        bool first = true;
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
                if (child is W.RunProperties) continue;
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                List<WordTextBox> boxes = view.GetTextBoxes().ToList();
                if (boxes.Count == 0) {
                    if (view.IsImage && view.Image != null) {
                        Flush();
                        RenderNativeImage(pdf, view.Image, options: options, source: "mixed text-box paragraph image");
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

        void Flush() {
            bool visible = pending.Any(run => IsNativeRenderableTextRun(run, paragraph));
            (int Level, string Marker)? currentMarker = first ? marker : null;
            if (!visible && currentMarker is not { Marker.Length: > 0 }) {
                pending.Clear();
                return;
            }
            var style = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
            if (currentMarker != null) {
                ApplyNativeInlineListIndent(paragraph, style);
                ApplyNativeInlineListMarkerAlignment(paragraph, currentMarker.Value.Marker, style, defaults, fontMap);
            }
            if (!first) style.SpacingBefore = 0D;
            style.SpacingAfter = 0D;
            WordParagraph[] segment = pending.ToArray();
            string text = string.Concat(segment.Select(run => run.Text));
            List<int> references = GetNativeParagraphFootnoteNumbers(segment.FirstOrDefault() ?? paragraph, segment,
                Array.Empty<int>(), notes);
            pdf.Paragraph(builder => AddNativeParagraphContent(builder, paragraph, currentMarker,
                segment, visible, text, references, notes, options, defaults, fontMap,
                inlineMarkerColumnWidth: currentMarker == null ? null : ResolveNativeInlineListMarkerColumnWidth(
                    paragraph, currentMarker.Value.Marker, style, defaults, fontMap)),
                ResolveNativeParagraphAlign(paragraph, allowJustify: false), style: style);
            first = false;
            pending.Clear();
        }
    }

    private static void RecordNativeMixedTextBoxDiagnostic(IReadOnlyList<WordParagraph> runs,
        WordToPdfOptions? options, string source) {
        if (options != null && runs.Any(run => run.GetTextBoxes().Any()) &&
            runs.Any(run => run.EnumerateEffectiveRunContent().OfType<W.Text>().Any()))
            AddNativeExportWarning(options, "NativeMixedTextBoxLayoutApproximated", source,
                "Text surrounding peer or nested text boxes is retained in source order; anchored box positioning and text wrapping are approximated by bounded flow.");
    }
}
