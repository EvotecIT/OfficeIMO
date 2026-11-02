using System.Collections.Generic;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Projects each visible diagram once, leaving its placement and width constraint to native PDF flow.</summary>
    private static List<OfficeDrawing> PrepareNativeRunSmartArt(WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, WordPdfExportOptions? options, string source,
        bool isHeaderFooter = false, bool reportDiagnostics = true) {
        var drawings = new List<OfficeDrawing>();
        var seen = new HashSet<W.Drawing>();
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            // A source run may contain several diagrams or a picture before a diagram.
            // Content views let the semantic SmartArt accessor resolve every selected drawing.
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                if (view.SmartArt is not WordSmartArt smartArt || !seen.Add(smartArt._drawing)) continue;
                OfficeDrawing? drawing;
                string projection;
                try {
                    drawing = WordSmartArtDrawingRenderer.CreateDrawing(smartArt, out projection,
                        options?.CancellationToken ?? default);
                } catch (Exception exception) when (exception is OpenXmlPackageException or InvalidOperationException or
                    KeyNotFoundException or ArgumentOutOfRangeException or XmlException) {
                    // Diagram parts outside the semantic reader's supported relationship layout,
                    // or malformed diagram data, must not discard the rest of the paragraph.
                    drawing = null;
                    projection = "unavailable";
                }
                if (drawing != null) drawings.Add(drawing);
                if (!reportDiagnostics || options == null) continue;
                string code = drawing == null
                    ? isHeaderFooter ? "native-header-footer-smart-art-unsupported" : "native-body-smart-art-unsupported"
                    : projection == "persisted" ? "native-smart-art-persisted-layout" : "native-smart-art-layout-approximated";
                string message = drawing == null
                    ? "SmartArt has no readable editable nodes or supported persisted geometry within the semantic reader's supported diagram relationships."
                    : projection == "persisted"
                        ? "Painted supported persisted SmartArt preset geometry and node positions with simplified colors and typography; full diagram styles, effects and connector geometry are not reconstructed."
                        : "Painted SmartArt node text using an approximate " + projection + " layout; authored diagram geometry, styles and relationships are not reproduced.";
                if (!options.Warnings.Any(warning => warning.Code == code && warning.Source == source && warning.Message == message))
                    AddNativeExportWarning(options, code, source, message);
            }
        }
        return drawings;
    }

    /// <summary>Records running-content diagram coverage before its deferred drawing callbacks execute.</summary>
    private static void RecordNativeHeaderFooterSmartArtDiagnostics(WordParagraph paragraph, WordPdfExportOptions options, string source) {
        List<WordParagraph> runs = GetNativeRuns(paragraph);
        if (paragraph._paragraph.Ancestors<W.TableCell>().Any() || paragraph._paragraph.Ancestors<W.TextBoxContent>().Any()) {
            if (runs.Any(run => run.IsSmartArt && !IsNativeHiddenTextRun(run, paragraph)))
                AddNativeExportWarning(options, "native-header-footer-smart-art-unsupported", source,
                    "SmartArt nested in a header/footer table cell or text box is not mapped by the shared diagram flow.");
            return;
        }
        PrepareNativeRunSmartArt(paragraph, runs, options, source, isHeaderFooter: true);
    }

    /// <summary>Enters semantic diagram scenes into the same measured flow as other body and running-content objects.</summary>
    private static bool RenderNativeSmartArt(INativePdfFlow pdf, IReadOnlyList<OfficeDrawing> drawings,
        PdfCore.PdfAlign align, NativeObjectParagraphSpacing? paragraphSpacing = null) {
        foreach (OfficeDrawing drawing in drawings) {
            RenderNativeFlowObject(pdf, paragraphSpacing, flow => flow.Drawing(drawing, align,
                spacingBefore: 0D, spacingAfter: 0D,
                style: new PdfCore.PdfDrawingStyle { ConstrainToContentWidth = true }), alignToLineTop: true);
        }
        return drawings.Count > 0;
    }
}
