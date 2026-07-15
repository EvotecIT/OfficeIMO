using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const uint OfficeArtBackgroundShapeFlag = 1U << 10;

        private LegacyPptBackground? ReadBackground(LegacyPptRecord ownerRecord,
            LegacyPptColorScheme? colorScheme, LegacyPptImportOptions options) {
            LegacyPptRecord? drawing = ownerRecord.Children.FirstOrDefault(record =>
                record.Type == RecordDrawing);
            LegacyPptRecord? container = drawing?.DescendantsAndSelf()
                .Where(record => record.Type == OfficeArtSpContainer)
                .LastOrDefault(IsBackgroundShapeContainer);
            if (container == null) return null;

            LegacyPptRecord? fopt = container.Children.FirstOrDefault(record =>
                record.Type == OfficeArtFopt);
            OfficeArtShapeStyle style = ReadShapeStyle(fopt);
            uint fillType = style.FillType.GetValueOrDefault();
            LegacyPptBackgroundKind kind = style.FillEnabled == false
                ? LegacyPptBackgroundKind.None
                : MapBackgroundKind(fillType);
            string? foreground = ResolveShapeColor(style.FillColor, colorScheme)
                ?? colorScheme?.Background;
            string? background = ResolveShapeColor(style.FillBackColor, colorScheme)
                ?? colorScheme?.Fill ?? foreground;
            int? pictureStoreIndex = style.FillBlipStoreIndex;
            OfficeArtBlipStoreEntry? picture = ResolvePicture(pictureStoreIndex);
            LegacyPptGradientStop[] gradientStops = style.FillGradientStops
                .Select(stop => new LegacyPptGradientStop(
                    ResolveShapeColor(stop.Color, colorScheme), stop.Position))
                .ToArray();
            bool hasUnresolvedGradientStop = gradientStops.Any(stop => stop.Color == null);
            var result = new LegacyPptBackground(kind, fillType, foreground, background,
                style.FillOpacity, style.FillBackOpacity, style.FillAngleDegrees,
                style.FillFocusPercent, pictureStoreIndex, picture, gradientStops,
                style.IsFillGradientStopTableTruncated || hasUnresolvedGradientStop);

            if (options.ReportUnsupportedContent && NeedsBackgroundDiagnostic(result)) {
                AddDiagnostic("PPT-BACKGROUND-PARTIAL", LegacyPptDiagnosticSeverity.Warning,
                    GetBackgroundDiagnostic(result), container.Offset);
            }
            return result;
        }

        private static bool IsBackgroundShapeContainer(LegacyPptRecord container) {
            LegacyPptRecord? fsp = container.Children.FirstOrDefault(record =>
                record.Type == OfficeArtFsp);
            return fsp != null && fsp.PayloadLength >= 8
                && (fsp.ReadUInt32(4) & OfficeArtBackgroundShapeFlag) != 0;
        }

        private static LegacyPptBackgroundKind MapBackgroundKind(uint fillType) => fillType switch {
            0 => LegacyPptBackgroundKind.Solid,
            1 => LegacyPptBackgroundKind.Pattern,
            2 => LegacyPptBackgroundKind.Texture,
            3 => LegacyPptBackgroundKind.Picture,
            4 => LegacyPptBackgroundKind.LinearGradient,
            5 => LegacyPptBackgroundKind.CenterGradient,
            6 => LegacyPptBackgroundKind.ShapeGradient,
            7 => LegacyPptBackgroundKind.ScaleGradient,
            8 => LegacyPptBackgroundKind.TitleGradient,
            9 => LegacyPptBackgroundKind.Inherited,
            _ => LegacyPptBackgroundKind.Unsupported
        };

        private static bool NeedsBackgroundDiagnostic(LegacyPptBackground background) =>
            !background.HasProjectableFill || background.IsGradientStopTableTruncated
            || background.FocusPercent.GetValueOrDefault() != 0
            || background.Kind == LegacyPptBackgroundKind.Pattern;

        private static string GetBackgroundDiagnostic(LegacyPptBackground background) {
            if (background.IsGradientStopTableTruncated) {
                return "The background contains a malformed, truncated, or unresolved OfficeArt gradient-stop table; its endpoint colors are projected and the exact table remains preserve-only.";
            }
            if (background.FocusPercent.GetValueOrDefault() != 0) {
                return "The background uses a focused OfficeArt gradient; its colors and direction are projected while the exact focus remains preserve-only.";
            }
            if ((background.Kind is LegacyPptBackgroundKind.Texture
                    or LegacyPptBackgroundKind.Picture)
                && background.Picture == null) {
                return "The image-backed background references missing or unsupported BLIP data and remains preserve-only.";
            }
            return $"The {background.Kind} OfficeArt background fill remains preserve-only.";
        }
    }
}
