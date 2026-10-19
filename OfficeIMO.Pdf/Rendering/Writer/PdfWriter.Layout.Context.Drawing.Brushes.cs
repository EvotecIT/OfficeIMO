using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private bool _suppressDrawingLinks;
        private bool _drawingIsOpacityData;

        private void DrawDrawingEffectAt(OfficeDrawingEffectGroup effect, double originX, double originTopY, OfficeDrawingTextMetrics metrics) {
            OfficeTransform transform = ToTopLeftPageTransform(effect.Transform, originX, originTopY);
            PageEffectGroup? maskGroup = null;
            if (effect.SoftMask != null) {
                OfficeDrawingSoftMask mask = effect.SoftMask;
                if (mask.BackdropColor.A != 0) {
                    throw new NotSupportedException("PDF drawing masks require a transparent backdrop.");
                }
                if (mask.Mode == OfficeSoftMaskMode.Luminosity &&
                    mask.LuminosityStandard != OfficeSoftMaskLuminosityStandard.PdfDeviceRgb) {
                    throw new NotSupportedException("PDF luminosity drawing masks require the PdfDeviceRgb luminosity standard.");
                }
                int start = sb.Length;
                bool previousLinks = _suppressDrawingLinks;
                _suppressDrawingLinks = true;
                bool previousWrappers = _suppressCanvasAccessibilityWrappers;
                bool previousText = _suppressCanvasActualTextChildren;
                bool previousOpacityData = _drawingIsOpacityData;
                OfficeTransform previousEffectToPage = _canvasEffectToPage;
                _suppressCanvasAccessibilityWrappers = true;
                _suppressCanvasActualTextChildren = true;
                // Luminosity colors encode opacity, so all nested mask resources
                // retain their numeric values instead of following print conversion.
                _drawingIsOpacityData |= mask.Mode == OfficeSoftMaskMode.Luminosity;
                _canvasEffectToPage = ConvertTopLeftCanvasTransform(transform, currentOpts.PageHeight).Then(previousEffectToPage);
                try {
                    RenderEffectGroup(OfficeTransform.Identity, 1D, OfficeBlendMode.Normal, () =>
                        RenderEffectGroup(ToTopLeftPageTransform(mask.Transform, originX, originTopY), 1D,
                            () => DrawDrawingElements(mask.InnerDrawing, originX, originTopY, metrics)), forceForm: true);
                    maskGroup = currentPage!.EffectGroups[currentPage.EffectGroups.Count - 1];
                } finally {
                    sb.Length = start;
                    _suppressDrawingLinks = previousLinks;
                    _suppressCanvasAccessibilityWrappers = previousWrappers;
                    _suppressCanvasActualTextChildren = previousText;
                    _drawingIsOpacityData = previousOpacityData;
                    _canvasEffectToPage = previousEffectToPage;
                }
            }
            int contentStart = sb.Length;
            RenderEffectGroup(transform, maskGroup == null ? effect.Opacity : 1D,
                maskGroup == null ? effect.BlendMode : OfficeBlendMode.Normal,
                () => DrawDrawingElements(effect.InnerDrawing, originX, originTopY, metrics), forceForm: maskGroup != null);
            if (maskGroup != null) {
                PageEffectGroup contentGroup = currentPage!.EffectGroups[currentPage.EffectGroups.Count - 1];
                contentGroup.SoftMask = maskGroup;
                contentGroup.SoftMaskMode = effect.SoftMask!.Mode;
                // Composite the completed masked result before applying group opacity.
                // Sharing its alpha state with the mask invocation attenuates the mask
                // as well as the result in PDF consumers.
                if (effect.Opacity < 1D || effect.BlendMode != OfficeBlendMode.Normal) {
                    string maskedContent = sb.ToString(contentStart, sb.Length - contentStart);
                    sb.Length = contentStart;
                    RenderEffectGroup(OfficeTransform.Identity, effect.Opacity, effect.BlendMode,
                        () => sb.Append(maskedContent), forceForm: true);
                }
            }
        }
    }
}
