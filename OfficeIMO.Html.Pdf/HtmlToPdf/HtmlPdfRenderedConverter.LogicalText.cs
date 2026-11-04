using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    /// <summary>Wraps outlined ink in its logical source without inventing a zero-advance spacing rectangle.</summary>
    private static void AddOutlinedLogicalTextReplacement(PdfCore.PdfPageCanvas canvas,
        HtmlRenderText visual, double textX, double advance, bool asSpan, ClipBounds? logicalClip,
        Action<PdfCore.PdfPageCanvas> paint) {
        double logicalHeight = Math.Max(0.01D, Math.Min(visual.Height, visual.Font.Size));
        var carrier = logicalClip?.ConstrainLogicalRectangle(visual.X + textX, visual.Y, advance, logicalHeight)
            ?? (X: visual.X + textX, Y: visual.Y, Width: advance, Height: logicalHeight);
        PdfCore.PdfCanvasTextStructureRole role = asSpan
            ? PdfCore.PdfCanvasTextStructureRole.Span
            : MapStructureRole(visual.SemanticRole);
        void AddReplacement(PdfCore.PdfPageCanvas target) {
            if (carrier.Width > 0D && carrier.Height > 0D) {
                target.ActualText(visual.Text,
                    carrier.X * PointsPerCssPixel, carrier.Y * PointsPerCssPixel,
                    carrier.Width * PointsPerCssPixel, carrier.Height * PointsPerCssPixel,
                    paint, logicalOrder: visual.SemanticSourceOrder);
            } else {
                // Zero advance still has a position and can have visible ink.
                target.ActualText(visual.Text, carrier.X * PointsPerCssPixel,
                    (carrier.Y + carrier.Height) * PointsPerCssPixel,
                    paint, logicalOrder: visual.SemanticSourceOrder);
            }
        }
        if (role == PdfCore.PdfCanvasTextStructureRole.Span) {
            AddReplacement(canvas);
        } else {
            canvas.Structure(MapOutlinedTextStructureRole(role), AddReplacement,
                new PdfCore.PdfCanvasStructureOptions { LogicalOrder = visual.SemanticSourceOrder });
        }
    }

    private static bool TryResolveReorderedLogicalText(IEnumerable<HtmlRenderVisual> visuals, out string logicalText) =>
        HtmlRenderLogicalText.TryResolveReorderedText(visuals, out logicalText);
    private static void AddLogicalTextGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderLogicalTextGroup group, RegisteredWebFonts webFonts, PdfImageResourceCache imageResources, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip, bool logicalTextOwned, HtmlPdfPagePaintContext? pagePaint) {
        void AddChildren(PdfCore.PdfPageCanvas target) {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(target, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned: true, pagePaint: pagePaint);
            }
        }
        if (!group.Visuals.Any(child => ContainsPdfRenderableVisual(child, webFonts, surfaceWidth, surfaceHeight, activeClip, cancellationToken))) {
            AddChildren(canvas);
            return;
        }
        var content = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
        AddChildren(content);
        if (!HasCanvasContent(content.Items)) {
            canvas.AddItems(content.Items);
            return;
        }
        string replacementText = group.Text;
        IEnumerable<HtmlRenderVisual> logicalPaint = group.Visuals;
        bool scopedText = group.LogicalScope != null && pagePaint != null;
        if (scopedText) {
            if (pagePaint!.IsClaimed(group.LogicalScope!)) {
                canvas.SuppressTextExtraction(nested => nested.AddItems(content.Items));
                return;
            }
            // Resolve coverage from all page-local fragments, rather than assuming
            // the first source layer can paint or knows other layers' font failures.
            logicalPaint = pagePaint!.GetLogicalPaint(group.LogicalScope!);
            HtmlRenderLogicalText.TryResolveSourceText(logicalPaint.Where(visual =>
                ContainsPdfRenderableVisual(visual, webFonts, surfaceWidth, surfaceHeight, activeClip, cancellationToken)), out replacementText,
                preserveBlockSeparators: group.LogicalScope!.PreserveBlockSeparators);
        }
        if (replacementText.Length == 0) {
            // Artifact marking excludes tagged reading order, but independent
            // extractors still read scalar glyphs. An empty replacement owns the
            // secondary paint without inventing another text value or hiding links.
            canvas.SuppressTextExtraction(nested => nested.AddItems(content.Items));
            return;
        }
        string? logicalText = FilterLogicalPrivateUseGlyphs(replacementText, logicalPaint, webFonts, cancellationToken);
        if (logicalText == null) {
            canvas.AddItems(content.Items);
            return;
        }
        if (logicalText.Length == 0) {
            canvas.AddItems(content.Items);
            return;
        }
        if (scopedText && !pagePaint!.TryClaim(group.LogicalScope!)) {
            canvas.SuppressTextExtraction(nested => nested.AddItems(content.Items));
            return;
        }
        if (logicalTextOwned) canvas.AddItems(content.Items);
        else {
            double logicalHeight = Math.Min(group.Height, 12D);
            ClipBounds clip = activeClip ?? new ClipBounds(0D, 0D, surfaceWidth, surfaceHeight);
            var carrier = clip.ConstrainLogicalRectangle(group.X, group.Y, group.Width, logicalHeight);
            if (carrier.X != group.X || carrier.Y != group.Y || carrier.Width != group.Width || carrier.Height != logicalHeight) {
                canvas.ActualText(logicalText, carrier.X * PointsPerCssPixel, carrier.Y * PointsPerCssPixel,
                    carrier.Width * PointsPerCssPixel, carrier.Height * PointsPerCssPixel,
                    nested => nested.AddItems(content.Items), logicalOrder: group.SemanticSourceOrder);
            } else canvas.ActualText(logicalText, group.X * PointsPerCssPixel,
                (group.Y + logicalHeight) * PointsPerCssPixel, nested => nested.AddItems(content.Items), logicalOrder: group.SemanticSourceOrder);
        }
    }

}
