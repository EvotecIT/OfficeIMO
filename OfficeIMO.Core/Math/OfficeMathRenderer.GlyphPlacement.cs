using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathRenderer {
    private sealed partial class LayoutEngine {
        /// <summary>OpenType's two-height corner-pair algorithm, in drawing units above each baseline.</summary>
        private static double ScriptKern(LayoutBox basis, LayoutBox script, double shift, bool over, bool left) {
            int baseCorner = over ? (left ? 1 : 0) : (left ? 3 : 2);
            int scriptCorner = over ? (left ? 2 : 3) : (left ? 0 : 1);
            double scriptEdge = over ? -(script.Height - script.Baseline) : script.Baseline;
            double baseEdge = over ? basis.Baseline : -(basis.Height - basis.Baseline);
            double first = KernAt(basis, baseCorner, scriptEdge + shift) + KernAt(script, scriptCorner, scriptEdge);
            double second = KernAt(basis, baseCorner, baseEdge) + KernAt(script, scriptCorner, baseEdge - shift);
            return Math.Min(first, second);
        }

        private static double KernAt(LayoutBox box, int corner, double height) {
            if (box.GlyphData == null || box.GlyphUnit <= 0D || !box.GlyphData.Kerns.TryGetValue(box.GlyphId, out var kerns)) return 0D;
            return (kerns[corner]?.At(height / box.GlyphUnit) ?? 0) * box.GlyphUnit;
        }

        private LayoutBox FontAccent(OfficeMathExpression expression, double scale) {
            LayoutBox content = CrampedLayout(expression.Children[0], scale);
            string accentText = _options.TokenPaintText?.Invoke(expression) ?? expression.Character ?? "^";
            // Accents frequently sit entirely above their baseline. Use the actual ink
            // bottom rather than reserving the otherwise empty baseline descent.
            LayoutBox accent = (expression.Stretchy == false ? null : StretchGlyph(accentText, scale,
                content.Width, horizontal: true, tightInk: true)) ?? Text(accentText, scale, expression.Character, tightInk: true);
            double accentX = (content.AccentAttachment ?? content.Width / 2D)
                - (accent.AccentAttachment ?? accent.Width / 2D);
            double left = Math.Min(0D, accentX);
            double width = Math.Max(content.Width, accentX + accent.Width) - left;
            double gap = Math.Max(0D, MathValue(OfficeMathConstant.AccentBaseHeight, scale) - content.Baseline);
            double contentY = accent.Height + gap;
            var box = new LayoutBox(width, contentY + content.Height, contentY + content.Baseline);
            box.Add(accent, accentX - left, 0D); box.Add(content, -left, contentY);
            box.AccentAttachment = (content.AccentAttachment ?? content.Width / 2D) - left;
            return box;
        }
    }
}
