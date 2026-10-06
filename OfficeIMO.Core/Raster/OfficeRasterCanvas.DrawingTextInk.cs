using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Embedded drawings paint into their own raster viewport before HTML places
    // them. Keep font/profile changes and group transforms in the Drawing owner.
    internal void InspectDrawingTextInk(OfficeDrawing drawing, OfficeTransform transform,
        IReadOnlyList<OfficeTextInkClip> outerClips,
        Action<(double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped), string?> report) {
        var clips = new List<OfficeTextInkClip>(outerClips);
        int work = 4096, remainingLayoutCharacters = 65536;
        VisitSurface(drawing, this, transform, 0);

        void Unmeasured(string reason) => report((0D, 0D, 0D, 0D, false, false, false), reason);
        void Charge() {
            _cancellationToken.ThrowIfCancellationRequested();
            if (--work < 0) throw new NotSupportedException("Drawing text ink inspection exceeds its 4096-element/run work limit.");
        }
        void VisitSurface(OfficeDrawing surface, OfficeRasterCanvas parent, OfficeTransform placement, int depth) {
            if (!placement.TryInvert(out OfficeTransform inverse)) {
                if (MayContainText(surface, depth)) Unmeasured("A singular drawing transform cannot establish text ink geometry."); return;
            }
            var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: surface.Fonts,
                textShapingProvider: surface.TextShapingProvider ?? parent.TextShapingProvider,
                textShapingLanguage: surface.TextShapingLanguage ?? parent.TextShapingLanguage,
                diagnosticSink: _diagnosticSink, diagnosticSource: _diagnosticSource, cancellationToken: _cancellationToken);
            clips.Add(new OfficeTextInkClip(0D, 0D, surface.Width, surface.Height, true, true, inverse));
            try { Visit(surface, canvas, placement, depth); } finally { clips.RemoveAt(clips.Count - 1); }
        }
        void Visit(OfficeDrawing current, OfficeRasterCanvas canvas, OfficeTransform placement, int depth) {
            if (depth > 64 || clips.Count > 64) throw new NotSupportedException("Drawing text ink inspection exceeds its 64 nested groups/clips limit.");
            foreach (OfficeDrawingElement element in current.Elements) {
                Charge();
                if (element is OfficeDrawingText text) InspectText(text, canvas, placement);
                else if (element is OfficeDrawingEffectGroup effect) {
                    if (effect.Opacity <= 0D) continue;
                    if (effect.SoftMask != null) {
                        if (MayContainText(effect.InnerDrawing, depth + 1) || MayContainText(effect.SoftMask.InnerDrawing, depth + 1))
                            Unmeasured("Masked drawing text visibility is not inspected.");
                        continue;
                    }
                    VisitSurface(effect.InnerDrawing, canvas, effect.Transform.Then(placement), depth + 1);
                } else if (element is OfficeDrawingGroup group) {
                    if (group.ClipPath.Kind == OfficeClipPathKind.Empty) continue;
                    if (group.ActualText != null) { Unmeasured("Text represented by vector outlines is not inspected as positioned text."); continue; }
                    OfficeTransform frame = group.FrameTransform?.CreateDestinationTransform() ?? OfficeTransform.Identity;
                    OfficeTransform clipPlacement = OfficeTransform.Translate(group.X, group.Y).Then(frame).Then(placement);
                    if (!OfficeTextInkClip.TryCreatePath(group.ClipPath, clipPlacement, _cancellationToken, out OfficeTextInkClip clip)) {
                        if (MayContainText(group.InnerDrawing, depth + 1)) Unmeasured("Unsupported, non-finite or over-budget drawing text clips are not inspected."); continue;
                    }
                    clips.Add(clip);
                    try {
                        Visit(group.InnerDrawing, canvas.WithDrawingTextProfile(group.InnerDrawing),
                            OfficeTransform.Translate(group.X + group.ContentOffsetX, group.Y + group.ContentOffsetY).Then(frame).Then(placement), depth + 1);
                    } finally { clips.RemoveAt(clips.Count - 1); }
                } else if (element is OfficeDrawingRichText rich) {
                    ChargeLayout(rich.PlainText.Length);
                    foreach (var run in rich.Runs) Charge();
                    canvas.InspectLaidOutTextInk(() => OfficeDrawingRasterRenderer.RenderRichText(canvas, rich, 1D),
                        placement, clips, Charge, report);
                }
                else if (element is OfficeDrawingTilingPattern pattern && pattern.Opacity > 0D && MayContainText(pattern.InnerTile, depth + 1))
                    Unmeasured("Repeated vector-pattern text is not inspected.");
                else if (element is OfficeDrawingImage || element is OfficeDrawingImagePattern)
                    Unmeasured("Text inside embedded image resources is not inspected.");
            }
        }
        // This is a conservative content check, not a visibility test. Opaque images
        // and outline metadata may represent text; unsupported text keeps its warning.
        bool MayContainText(OfficeDrawing current, int depth) {
            if (depth > 64) throw new NotSupportedException("Drawing text ink inspection exceeds its 64 nested groups/clips limit.");
            foreach (var element in current.Elements) {
                Charge();
                if (element is OfficeDrawingText text) {
                    if (text.RasterText.Length != 0) return true;
                } else if (element is OfficeDrawingRichText || element is OfficeDrawingImage || element is OfficeDrawingImagePattern) return true;
                else if (element is OfficeDrawingEffectGroup effect) {
                    if (effect.Opacity > 0D && (MayContainText(effect.InnerDrawing, depth + 1) ||
                        (effect.SoftMask != null && MayContainText(effect.SoftMask.InnerDrawing, depth + 1)))) return true;
                } else if (element is OfficeDrawingGroup group) {
                    if (group.ClipPath.Kind != OfficeClipPathKind.Empty &&
                        (group.ActualText != null || MayContainText(group.InnerDrawing, depth + 1))) return true;
                } else if (element is OfficeDrawingTilingPattern pattern && pattern.Opacity > 0D && MayContainText(pattern.InnerTile, depth + 1)) return true;
            }
            return false;
        }

        void ChargeLayout(int length) {
            remainingLayoutCharacters -= length;
            if (remainingLayoutCharacters < 0) throw new NotSupportedException("Drawing text ink inspection exceeds its 65536 laid-out character limit.");
        }

        void InspectText(OfficeDrawingText text, OfficeRasterCanvas canvas, OfficeTransform placement) {
            if (text.RasterText.Length == 0 || (text.Color ?? OfficeColor.Black).A == 0) return;
            bool positioned = !text.WrapText && !text.ShrinkToFit && !text.StackedText && !text.HasPadding
                && text.VerticalAlignment == OfficeTextVerticalAlignment.Top && text.TextDirection != OfficeTextDirection.TopToBottom;
            bool usesPositionedPaint = text.TextAdvanceWidth.HasValue || text.OverflowBehavior == OfficeTextOverflowBehavior.Clip
                || text.BaselineScale != 1D || text.BaselineOffset != 0D || !text.FeatureSettings.IsDefault
                || !string.Equals(text.FontPalette, "normal", StringComparison.OrdinalIgnoreCase);
            if (!positioned && text.TextDirection != OfficeTextDirection.TopToBottom) {
                ChargeLayout(text.RasterText.Length);
                bool saved = canvas.PreservePaintedGlyphOrder;
                canvas.PreservePaintedGlyphOrder = text.PreservesPaintedGlyphs;
                try {
                    canvas.InspectLaidOutTextInk(() => OfficeDrawingRasterRenderer.RenderText(canvas, text, 1D, 1L),
                        placement, clips, Charge, report);
                } finally { canvas.PreservePaintedGlyphOrder = saved; }
                return;
            }
            if (!positioned || !usesPositionedPaint || (text.HasFrameTransform && !text.TextAdvanceWidth.HasValue)) {
                Unmeasured("This drawing text layout does not use the supported positioned-paint path."); return;
            }
            OfficeTransform inkTransform = text.HasFrameTransform
                ? text.CreateFrameTransform().CreateDestinationTransform().Then(placement) : placement;
            double sourceSize = Math.Max(1D, text.Font.Size), size = sourceSize * text.BaselineScale;
            double lineHeight = text.LineHeight ?? text.Font.Size * 1.2D;
            string[] lines = text.RasterText.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            bool preserve = canvas.PreservePaintedGlyphOrder;
            canvas.PreservePaintedGlyphOrder = text.PreservesPaintedGlyphs;
            using var faceScope = canvas.PushTextFace(text.Font.Face);
            try {
                for (int index = 0; index < lines.Length; index++) {
                    Charge();
                    double offset = index * lineHeight;
                    if (offset >= text.Height) break;
                    if (lines[index].Length == 0) continue;
                    double advance = lines.Length == 1 && text.TextAdvanceWidth.HasValue ? text.TextAdvanceWidth.Value
                        : Math.Max(.001D, canvas.MeasurePositionedText(lines[index], size, text.Font.FamilyName,
                            text.Font.Style, text.FeatureSettings, text.TextDirection));
                    report(canvas.MeasurePositionedTextBounds(lines[index], text.X, text.Y + offset + text.BaselineOffset,
                        text.Width, text.Height - offset, size, text.Font, advance, text.Alignment, text.FeatureSettings,
                        text.FontPalette, sourceSize, text.UnderlineStyle, text.StrikethroughStyle, text.TextDirection,
                        inkOnly: true, inkTransform: inkTransform, color: text.Color, decorationColor: text.DecorationColor,
                        inkClips: clips), null);
                }
            } finally { canvas.PreservePaintedGlyphOrder = preserve; }
        }
    }
}
