using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingPortableExporter {
    private sealed partial class ProjectionContext {
        internal Element ProjectText(OfficeDrawingText text, OfficeTransform transform, OfficeFontFaceCollection fonts, string location) {
            CountText(text.Text.Length, location);
            if (!string.Equals(text.Font.FamilyName, Options.FontFamily, StringComparison.OrdinalIgnoreCase)) {
                // SVG producers may append generic fallbacks. The qualified face must
                // remain first; the face and glyph checks below still forbid substitution.
                System.Collections.Generic.List<string> families = OfficeFontFamilyParser.Parse(text.Font.FamilyName, requirePrimaryFamily: true);
                if (families.Count == 0 || !string.Equals(families[0], Options.FontFamily, StringComparison.OrdinalIgnoreCase))
                    Fail("DRAWING_FONT_FAMILY_UNSUPPORTED", "Every text run must use the explicitly selected font family.", location);
            }
            if (text.WrapText || text.ShrinkToFit || text.StackedText || text.HasFrameTransform || text.GlyphTransform.HasValue || text.HasPadding || text.HasParagraphIndent ||
                text.VerticalAlignment != OfficeTextVerticalAlignment.Top || text.Baseline != OfficeTextBaseline.Normal ||
                text.BaselineLevel != 0 || text.BaselineScale != 1 || text.BaselineOffset != 0 || text.FontMetricScale != 1 ||
                text.TextDirection != OfficeTextDirection.Auto && text.TextDirection != OfficeTextDirection.LeftToRight ||
                text.PreservesPaintedGlyphs || text.RasterText != text.Text || !text.FeatureSettings.IsDefault || text.FontPalette != "normal" ||
                text.UnderlineStyle != OfficeTextDecorationStyle.None || text.StrikethroughStyle != OfficeTextDecorationStyle.None ||
                text.Font.IsUnderline || text.Font.IsStrikethrough || text.OverflowBehavior != OfficeTextOverflowBehavior.Clip)
                Fail("DRAWING_TEXT_LAYOUT_UNSUPPORTED", "Only plain single-line, unrotated positioned text is supported.", location);
            OfficeFontFaceDescriptor descriptor = text.Font.Face;
            if ((descriptor.Weight != 400 && descriptor.Weight != 700) || descriptor.StretchPercent != 100 || descriptor.Slant == OfficeFontSlant.Oblique)
                Fail("DRAWING_FONT_FACE_UNSUPPORTED", "Only explicit regular, bold, italic and bold-italic static faces are supported.", location);
            foreach (char c in text.Text) {
                bool latin = c >= 0x20 && c <= 0x7e || c >= 0xa0 && c <= 0x024f && c != 0xad ||
                    c >= 0x2010 && c <= 0x2027 || c >= 0x2030 && c <= 0x205e || c == 0x20ac || c == 0x2212;
                if (!latin) Fail("DRAWING_TEXT_SCRIPT_UNSUPPORTED", "Text is outside the qualified single-line Latin cohort.", location);
            }
            OfficeFontFace? face = null;
            foreach (OfficeFontFace candidate in fonts.Faces)
                if (string.Equals(candidate.FamilyName, Options.FontFamily, StringComparison.OrdinalIgnoreCase) && candidate.Descriptor == descriptor) { face = candidate; break; }
            if (face == null || !face.CanEmbedAsStaticPdfFont || face.ContainerFormat != OfficeFontContainerFormat.OpenType ||
                !(face.ParsedFont is OfficeTrueTypeFont trueType) || trueType.IsVariableFont || trueType.HasColorGlyphs || !face.Covers(text.Text))
                Fail("DRAWING_FONT_UNRESOLVED", "Register a matching static TrueType face containing the complete text; fallback and synthetic styles are not supported.", location);
            double size = Length(text.Font.Size, location);
            double advance = face!.Program.Measure(text.Text, size);
            Coordinate(advance, location);
            if (text.TextAdvanceWidth.HasValue && Math.Abs(text.TextAdvanceWidth.Value - advance) > 0.000001)
                Fail("DRAWING_TEXT_ADVANCE_UNSUPPORTED", "A source run with a forced advance cannot be preserved by the portable text profile.", location);
            double x = text.X;
            if (text.Alignment == OfficeTextAlignment.Center) x += (text.Width - advance) / 2;
            else if (text.Alignment == OfficeTextAlignment.Right) x += text.Width - advance;
            else if (text.Alignment != OfficeTextAlignment.Left) Fail("DRAWING_TEXT_LAYOUT_UNSUPPORTED", "Unknown text alignment.", location);
            OfficePoint baseline = Point(new OfficePoint(x, text.Y + size), transform, location);
            double outputSize = Length(size * transform.M11, location);
            if (outputSize > 144) Fail("DRAWING_FONT_SIZE_LIMIT", "Portable text size cannot exceed 144 points.", location);
            bool bold = descriptor.Weight == 700, italic = descriptor.Slant == OfficeFontSlant.Italic;
            var result = new Element { Kind = "text", Text = text.Text, X = baseline.X, Y = baseline.Y,
                FontSize = outputSize, Font = bold ? italic ? "boldItalic" : "bold" : italic ? "italic" : "regular",
                Color = Paint(text.Color ?? OfficeColor.Black, null, location) };
            var program = (OfficeTrueTypeFont)face.ParsedFont;
            var positions = program.PositionPortableScalars(text.Text, size);
            double scalarCursor = 0;
            bool needsPositions = false;
            for (int i = 0; i < text.Text.Length; i++) {
                if (Math.Abs(positions[i] - scalarCursor) > 0.000001) needsPositions = true;
                program.TryGetGlyphMetrics(text.Text[i], out _, out int nominalAdvance);
                scalarCursor += nominalAdvance * size / face.Program.UnitsPerEm;
            }
            System.Collections.Generic.List<Element> children;
            if (needsPositions) {
                RequireElements(text.Text.Length, location);
                children = new System.Collections.Generic.List<Element>(text.Text.Length);
                for (int i = 0; i < text.Text.Length; i++) {
                    CountElement(location);
                    double characterX = baseline.X + positions[i] * transform.M11;
                    Coordinate(characterX, location);
                    children.Add(new Element { Kind = "text", Text = text.Text[i].ToString(), X = characterX, Y = baseline.Y,
                        FontSize = outputSize, Font = result.Font, Color = result.Color });
                }
            } else children = new System.Collections.Generic.List<Element> { result };
            if (!text.ClipToFrame && !needsPositions) return result;
            if (!needsPositions) CountElement(location);
            return new Element { Kind = "group",
                Clip = text.ClipToFrame ? Clip(OfficeClipPath.Rectangle(text.Width, text.Height), text.X, text.Y, transform, location) : (Rect?)null,
                Children = children };
        }
    }
}
