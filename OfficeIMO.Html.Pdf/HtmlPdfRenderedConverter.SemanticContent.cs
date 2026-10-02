using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool HasCanvasContent(IReadOnlyList<PdfCore.PdfCanvasItem> items) {
        foreach (PdfCore.PdfCanvasItem item in items) {
            bool hasContent = item switch {
                PdfCore.PdfCanvasOutlineItem => false,
                PdfCore.PdfCanvasNamedDestinationItem => false,
                PdfCore.PdfCanvasNamedDestinationLinkItem => false,
                PdfCore.PdfCanvasArtifactItem artifact => HasCanvasContent(artifact.Items),
                PdfCore.PdfCanvasFigureItem figure => HasCanvasContent(figure.Items),
                PdfCore.PdfCanvasStructureItem structure => HasCanvasContent(structure.Items),
                PdfCore.PdfCanvasActualTextItem actualText => HasCanvasContent(actualText.Items),
                PdfCore.PdfCanvasClipItem clip => HasCanvasContent(clip.Items),
                PdfCore.PdfCanvasEffectItem effect => HasCanvasContent(effect.Items),
                _ => true
            };
            if (hasContent) return true;
        }
        return false;
    }

    private static void AddSemanticGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderSemanticGroup group, RegisteredWebFonts webFonts, PdfImageResourceCache imageResources, PdfCore.PdfConversionReport conversionReport, double surfaceWidth, double surfaceHeight, bool interactiveFormControls, CancellationToken cancellationToken, bool textAsSpan, ClipBounds? activeClip, bool logicalTextOwned, HtmlPdfPagePaintContext? pagePaint) {
        if (!group.Visuals.Any(child => ContainsPdfRenderableVisual(child, webFonts, surfaceWidth, surfaceHeight, activeClip, cancellationToken))) {
            // Navigation-only groups still carry named destinations. They cannot create
            // an empty structure element. The same path handles groups whose paint is
            // entirely outside the page, while non-painting children still reach the
            // page canvas so empty anchors remain valid link targets.
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                AddVisual(canvas, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls, cancellationToken, textAsSpan, activeClip, logicalTextOwned, pagePaint);
            }
            return;
        }
        if (group.Role == HtmlRenderSemanticGroupRole.Artifact) {
            var artifactContent = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(artifactContent, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls: false, cancellationToken, textAsSpan: true, activeClip: activeClip, logicalTextOwned: logicalTextOwned, pagePaint: pagePaint);
            }
            if (HasCanvasContent(artifactContent.Items))
                canvas.Artifact(nested => nested.AddItems(artifactContent.Items));
            else canvas.AddItems(artifactContent.Items);
            return;
        }
        var options = new PdfCore.PdfCanvasStructureOptions {
            AlternativeText = string.IsNullOrWhiteSpace(group.AlternativeText) ? null : group.AlternativeText,
            ColumnSpan = group.ColumnSpan,
            RowSpan = group.RowSpan,
            HeaderScope = MapTableHeaderScope(group.HeaderScope),
            StructureElementKey = group.StructureElementKey
        };
        if (group.Role == HtmlRenderSemanticGroupRole.Formula && group.MathMlSource != null)
            pagePaint?.MathMlFiles.Attach(options, group.MathMlSource, cancellationToken);
        bool childTextAsSpan = textAsSpan || IsTextContentGroup(group.Role);
        string logicalText = string.Empty;
        bool hasLogicalText = IsTextContentGroup(group.Role)
            && TryResolveReorderedLogicalText(group.Visuals, out logicalText);
        string? printableText = hasLogicalText
            ? FilterLogicalPrivateUseGlyphs(logicalText, group.Visuals, webFonts, cancellationToken)
            : null;
        bool outlineLimitReachedBeforeChildren = webFonts.OutlineBudget.IsPathLimitReached;
        var content = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
        foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(content, child, webFonts, imageResources, conversionReport, surfaceWidth, surfaceHeight, interactiveFormControls,
                cancellationToken, childTextAsSpan, activeClip, logicalTextOwned || !string.IsNullOrEmpty(printableText), pagePaint);
        }
        if (hasLogicalText && !outlineLimitReachedBeforeChildren && webFonts.OutlineBudget.IsPathLimitReached)
            printableText = FilterLogicalPrivateUseGlyphs(logicalText, group.Visuals, webFonts, cancellationToken);
        if (!HasCanvasContent(content.Items)) {
            canvas.AddItems(content.Items);
            return;
        }
        canvas.Structure(MapSemanticGroupRole(group.Role), nested => {
            if (printableText is { Length: > 0 } actualText)
                nested.ActualText(actualText, target => target.AddItems(content.Items));
            else nested.AddItems(content.Items);
        }, options);
    }

    private static PdfCore.PdfCanvasStructureRole MapSemanticGroupRole(HtmlRenderSemanticGroupRole role) {
        if (role == HtmlRenderSemanticGroupRole.Section) return PdfCore.PdfCanvasStructureRole.Section;
        if (role == HtmlRenderSemanticGroupRole.Division) return PdfCore.PdfCanvasStructureRole.Division;
        if (role == HtmlRenderSemanticGroupRole.Paragraph) return PdfCore.PdfCanvasStructureRole.Paragraph;
        if (role == HtmlRenderSemanticGroupRole.Heading1) return PdfCore.PdfCanvasStructureRole.Heading1;
        if (role == HtmlRenderSemanticGroupRole.Heading2) return PdfCore.PdfCanvasStructureRole.Heading2;
        if (role == HtmlRenderSemanticGroupRole.Heading3) return PdfCore.PdfCanvasStructureRole.Heading3;
        if (role == HtmlRenderSemanticGroupRole.Heading4) return PdfCore.PdfCanvasStructureRole.Heading4;
        if (role == HtmlRenderSemanticGroupRole.Heading5) return PdfCore.PdfCanvasStructureRole.Heading5;
        if (role == HtmlRenderSemanticGroupRole.Heading6) return PdfCore.PdfCanvasStructureRole.Heading6;
        if (role == HtmlRenderSemanticGroupRole.List) return PdfCore.PdfCanvasStructureRole.List;
        if (role == HtmlRenderSemanticGroupRole.ListItem) return PdfCore.PdfCanvasStructureRole.ListItem;
        if (role == HtmlRenderSemanticGroupRole.ListLabel) return PdfCore.PdfCanvasStructureRole.ListLabel;
        if (role == HtmlRenderSemanticGroupRole.ListBody) return PdfCore.PdfCanvasStructureRole.ListBody;
        if (role == HtmlRenderSemanticGroupRole.Table) return PdfCore.PdfCanvasStructureRole.Table;
        if (role == HtmlRenderSemanticGroupRole.TableRow) return PdfCore.PdfCanvasStructureRole.TableRow;
        if (role == HtmlRenderSemanticGroupRole.TableHeaderCell) return PdfCore.PdfCanvasStructureRole.TableHeaderCell;
        if (role == HtmlRenderSemanticGroupRole.TableCell) return PdfCore.PdfCanvasStructureRole.TableCell;
        if (role == HtmlRenderSemanticGroupRole.Footnote) return PdfCore.PdfCanvasStructureRole.Note;
        if (role == HtmlRenderSemanticGroupRole.Formula) return PdfCore.PdfCanvasStructureRole.Formula;
        return PdfCore.PdfCanvasStructureRole.Caption;
    }

}
