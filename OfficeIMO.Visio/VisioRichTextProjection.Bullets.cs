using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed partial class VisioRichTextProjection {
    // Bullet cells retain their original native values. This adapter creates only the
    // render-time label; it never inserts marker characters into Text or Label.
    private static bool TryParagraphLabel(XElement row, OfficeRichTextRun firstRun, VisioDocument? document,
        double scale, double position, int remainingCharacters, out OfficeTextParagraphLabel? label) {
        label = null;
        string? value = Cell(row, "Bullet");
        if (value == null) return true;
        if (!TryInt(value, out int kind) || kind < 0) return false;
        if (kind == 0) return true;
        string? custom = Cell(row, "BulletStr");
        // Visio 2002 writes U+E000 for an empty custom marker.
        if (custom == "\ue000") custom = null;
        bool hasCustom = !string.IsNullOrEmpty(custom);
        string? text = hasCustom ? custom : kind switch {
            1 => "\u2022", 2 => "\u25c6", 3 => "\u25a0", 4 => "\u25a1",
            5 => "\u2756", 6 => "\u27a2", 7 => "\u2714", _ => null
        };
        if (text == null || text.Length > 4096 || text.Length > remainingCharacters
            || text.IndexOfAny(new[] { '\r', '\n', '\t' }) >= 0) return false;
        if (!Finite(position) || position < 0D) return false;
        double distance = .25D;
        if (Cell(row, "TextPosAfterBullet") != null) {
            if (!TryParagraphNumber(row, "TextPosAfterBullet", out distance) || distance < 0D) return false;
            if (distance == 0D) distance = .25D;
        }
        distance *= scale;
        if (!Finite(distance) || !Finite(position + distance)) return false;
        double size = firstRun.FontSize;
        if (Cell(row, "BulletFontSize") != null) {
            if (!TryParagraphNumber(row, "BulletFontSize", out double nativeSize)) return false;
            // Native inches for absolute sizes; negative cached values are ratios (−1 = 100%).
            if (nativeSize > 0D) size = nativeSize * scale;
            else if (nativeSize < 0D) size *= -nativeSize;
        }
        if (!Finite(size) || size <= 0D) return false;
        string family = firstRun.FontFamily;
        if (hasCustom) {
            string? font = Cell(row, "BulletFont");
            // Zero selects the paragraph's first character font, rather than FaceNames ID 0.
            if (font != null && TryInt(font, out int fontId)) {
                if (fontId < 0) return false;
                if (fontId > 0) {
                    string? name = document?.PreservedFaceNamesElements.FirstOrDefault(face =>
                        TryInt((string?)face.Attribute("ID"), out int id) && id == fontId)?.Attribute("Name")?.Value;
                    if (string.IsNullOrWhiteSpace(name)) return false;
                    family = Font(font, document, family);
                }
            } else if (font != null) family = Font(font, document, family);
        }
        var run = new OfficeRichTextRun(text, size, firstRun.Color, firstRun.Bold, firstRun.Italic, fontFamily: family);
        label = OfficeTextParagraphLabel.AtPosition(run, position, textPosition: position + distance);
        return true;
    }
}
