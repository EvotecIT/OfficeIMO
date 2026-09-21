using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly Dictionary<byte[], IOfficeBoundedFontProgram> verticalOutlineFonts = new Dictionary<byte[], IOfficeBoundedFontProgram>();

        private bool TryDrawNativeVerticalText(OfficeDrawingText text, double originX, double originTopY) {
            IOfficeTextShapingProvider? provider = currentOpts.TextShapingProviderSnapshot;
            if (provider == null || (_suppressCanvasAccessibilityWrappers && !_suppressCanvasActualTextChildren) ||
                text.WrapText || text.ShrinkToFit || text.HasPadding ||
                text.HasParagraphIndent || text.Alignment != OfficeTextAlignment.Center ||
                text.VerticalAlignment != OfficeTextVerticalAlignment.Top ||
                text.Font.IsUnderline || text.Font.IsStrikethrough ||
                text.UnderlineStyle != OfficeTextDecorationStyle.None ||
                text.StrikethroughStyle != OfficeTextDecorationStyle.None ||
                text.Baseline != OfficeTextBaseline.Normal || text.BaselineLevel != 0 ||
                text.TextAdvanceWidth.HasValue ||
                !string.Equals(text.FontPalette, "normal", StringComparison.OrdinalIgnoreCase) ||
                text.Text.Contains('\r') || text.Text.Contains('\n')) return false;

            double size = text.Font.Size * text.BaselineScale;
            var run = new PdfTextRun(text.Text, text.Font.IsBold, text.Font.IsUnderline,
                    ToPdfColor(text.Color ?? OfficeColor.Black), text.Font.IsItalic, text.Font.IsStrikethrough,
                    size, ResolveDrawingTextFont(text.Font.FamilyName), fontFamily: text.Font.FamilyName,
                    underlineStyle: text.UnderlineStyle, strikeStyle: text.StrikethroughStyle,
                    decorationColor: ToPdfColor(text.DecorationColor))
                .WithFeatureSettings(text.FeatureSettings)
                .WithTextDirection(OfficeTextDirection.TopToBottom);
            PdfStandardFont font = ResolveFontForRun(run, ChooseNormal(currentOpts.DefaultFont));
            PdfNamedFontFace? namedFont = currentOpts.TryResolveNamedFontFace(
                run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedFace)
                ? resolvedFace : null;
            PdfTextShapingOptions shapingOptions = PdfTextShapingOptions.ForRendering(
                namedFont?.FamilyName ?? text.Font.FamilyName ?? font.ToString(),
                currentOpts.TextShapingModeSnapshot, provider,
                currentOpts.RecordProviderShapedTextRunDelegate, currentOpts.Language,
                text.FeatureSettings, OfficeTextDirection.TopToBottom);
            PdfGlyphRun glyphRun;
            byte[] fontData;
            int fontUnitsPerEm;
            if (namedFont.HasValue && currentOpts.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? trueType) && trueType != null) {
                if (!PdfExternalTextShaper.TryShapeText(text.Text, trueType, shapingOptions, out glyphRun)) return false;
                fontData = trueType.FontDataForInspection;
                fontUnitsPerEm = trueType.UnitsPerEm;
            } else if (namedFont.HasValue && currentOpts.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? cff) && cff != null) {
                if (!PdfExternalTextShaper.TryShapeText(text.Text, cff, shapingOptions, out glyphRun)) return false;
                fontData = cff.FontDataForInspection;
                fontUnitsPerEm = cff.UnitsPerEm;
            } else if (!namedFont.HasValue && currentOpts.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? standardTrueType) && standardTrueType != null) {
                if (!PdfExternalTextShaper.TryShapeText(text.Text, standardTrueType, shapingOptions, out glyphRun)) return false;
                fontData = standardTrueType.FontDataForInspection;
                fontUnitsPerEm = standardTrueType.UnitsPerEm;
            } else if (!namedFont.HasValue && currentOpts.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? standardCff) && standardCff != null) {
                if (!PdfExternalTextShaper.TryShapeText(text.Text, standardCff, shapingOptions, out glyphRun)) return false;
                fontData = standardCff.FontDataForInspection;
                fontUnitsPerEm = standardCff.UnitsPerEm;
            } else {
                return false;
            }
            if (!glyphRun.HasCompleteVerticalAdvances || glyphRun.Glyphs.Count == 0 ||
                !HasCompleteLogicalCoverage(text.Text, glyphRun.Glyphs)) return false;
            bool movesVertically = false;
            for (int index = 0; index < glyphRun.Glyphs.Count; index++) {
                movesVertically |= glyphRun.Glyphs[index].AdvanceHeight1000 != 0;
            }
            if (!movesVertically) return false;
            if (!TryGetVerticalInkTop(text.Text, glyphRun, fontData, fontUnitsPerEm, size, out double inkTop)) return false;

            double frameX = originX + text.X;
            double frameTopY = originTopY - text.Y;
            double penX = frameX + text.Width / 2D;
            // Core contours use a downward Y axis; PDF uses an upward Y axis.
            // Shaping providers disagree on vertical baseline conventions. Move ink
            // below the top clip edge only when the supplied origin would crop it.
            double penY = frameTopY + Math.Min(0D, inkTop);
            string fontResource = GetFontResourceName(font, namedFont, ChooseNormal(currentOpts.DefaultFont));
            OfficeColor paintColor = text.Color ?? OfficeColor.Black;
            string? opacityState = EnsureGraphicsState(paintColor.A / 255D, 1D);
            void Paint() {
                var content = new ContentStreamBuilder(sb).SaveState();
                if (opacityState != null) content.GraphicsState(opacityState);
                content
                    .Rectangle(frameX, frameTopY - text.Height, text.Width, text.Height)
                    .ClipPath()
                    .EndPath()
                    .FillColor(ToPdfColor(paintColor) ?? PdfColor.Black)
                    .BeginText()
                    .Font(fontResource, size);
                for (int index = 0; index < glyphRun.Glyphs.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfGlyphInfo glyph = glyphRun.Glyphs[index];
                    content.TextMatrix(
                        penX + glyph.OffsetX1000 * size / 1000D,
                        penY + glyph.OffsetY1000 * size / 1000D)
                        .ShowHexText(glyph.GlyphId.ToString("X4", CultureInfo.InvariantCulture));
                    penX += glyph.AdvanceWidth1000 * size / 1000D;
                    penY += glyph.AdvanceHeight1000 * size / 1000D;
                }
                content.EndText().RestoreState();
            }
            if (text.HasFrameTransform) {
                OfficeTransform transform = ToTopLeftPageTransform(
                    text.CreateFrameTransform().CreateDestinationTransform(), originX, originTopY);
                RenderEffectGroup(transform, 1D, Paint);
            } else {
                Paint();
            }
            MarkRichFonts(new[] { run });
            pageDirty = true;
            return true;
        }

        private bool TryGetVerticalInkTop(string text, PdfGlyphRun glyphRun, byte[] fontData,
            int unitsPerEm, double size, out double inkTop) {
            inkTop = 0D;
            OfficeTextShapingResult? result = glyphRun.SourceShapingResult;
            if (result == null) return false;
            if (!verticalOutlineFonts.TryGetValue(fontData, out IOfficeBoundedFontProgram? font)) {
                var faces = new OfficeFontFaceCollection();
                if (!faces.TryAdd("PDF vertical ink", fontData) || faces.Faces.Count != 1 ||
                    faces.Faces[0].Program is not IOfficeBoundedFontProgram parsed) return false;
                font = parsed;
                verticalOutlineFonts.Add(fontData, font);
            }
            if (font.UnitsPerEm != unitsPerEm) return false;
            List<List<OfficePoint>> contours;
            try {
                contours = font.GetShapedTextContoursBounded(text, result, 0D, 0D, size,
                    maximumPointCount: 100_000, cancellationToken);
            } catch (OperationCanceledException) {
                throw;
            } catch (ArgumentException) {
                return false;
            } catch (InvalidOperationException) {
                return false;
            } catch (NotSupportedException) {
                return false;
            } catch (InvalidDataException) {
                return false;
            } catch (OverflowException) {
                return false;
            }
            double minimumY = double.PositiveInfinity;
            foreach (List<OfficePoint> contour in contours) {
                foreach (OfficePoint point in contour) {
                    if (double.IsNaN(point.Y) || double.IsInfinity(point.Y)) return false;
                    minimumY = Math.Min(minimumY, point.Y);
                }
            }
            if (double.IsPositiveInfinity(minimumY)) return false;
            inkTop = minimumY;
            return true;
        }

        private static bool HasCompleteLogicalCoverage(string source, IReadOnlyList<PdfGlyphInfo> glyphs) {
            var ranges = new List<(int Start, int End)>(glyphs.Count);
            for (int index = 0; index < glyphs.Count; index++) {
                PdfGlyphInfo glyph = glyphs[index];
                if (glyph.UnicodeText.Length == 0) continue;
                if (glyph.TextIndex < 0 || glyph.TextIndex > source.Length - glyph.UnicodeText.Length ||
                    string.CompareOrdinal(source, glyph.TextIndex, glyph.UnicodeText, 0,
                        glyph.UnicodeText.Length) != 0) return false;
                ranges.Add((glyph.TextIndex, glyph.TextIndex + glyph.UnicodeText.Length));
            }
            ranges.Sort((left, right) => left.Start.CompareTo(right.Start));
            int covered = 0;
            for (int index = 0; index < ranges.Count; index++) {
                if (ranges[index].Start > covered) return false;
                covered = Math.Max(covered, ranges[index].End);
            }
            return covered == source.Length;
        }
    }
}
