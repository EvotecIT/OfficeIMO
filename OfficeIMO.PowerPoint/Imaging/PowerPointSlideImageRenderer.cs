using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static A.ColorScheme? GetSlideColorScheme(PowerPointSlide slide) =>
            slide.ResolvedColorScheme;

        private static A.FormatScheme? GetSlideFormatScheme(PowerPointSlide slide) =>
            slide.ResolvedFormatScheme;

        /// <summary>Projects the existing shared shape appearance for preset and text-frame export.</summary>
        internal static void ApplyShapeAppearanceForExport(OfficeShape target, PowerPointShape source,
            List<OfficeImageExportDiagnostic> diagnostics) {
            bool transformedAncestor = source.Element.Ancestors<DocumentFormat.OpenXml.Presentation.GroupShape>()
                .Any(group => group.GroupShapeProperties?.TransformGroup is { } transform
                    && ((transform.Rotation?.Value ?? 0) != 0 || transform.HorizontalFlip?.Value == true
                        || transform.VerticalFlip?.Value == true));
            ApplyShapeStyle(target, source, source.OwnerSlide?.ResolvedColorScheme,
                PowerPointShapeBoundsMapping.Identity.WithTransformedAncestor(transformedAncestor), diagnostics);
            if (source.UnsupportedFillKindForExport is { } fillKind) {
                AddUnsupportedShapeDiagnostic(diagnostics, source,
                    "Skipped a PowerPoint " + fillKind + " fill that is not supported by the shared Drawing renderer.");
            }
        }

        private static void ApplyShapeStyle(OfficeShape target, PowerPointShape source,
            A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping,
            List<OfficeImageExportDiagnostic> diagnostics) {
            ShapeFillGradientProjection gradientProjection = ApplyShapeFillGradient(
                target, source, colorScheme, mapping.HasTransformedAncestor);
            if (gradientProjection == ShapeFillGradientProjection.Unsupported) {
                AddUnsupportedShapeDiagnostic(diagnostics, source,
                    "Skipped a PowerPoint shape gradient that cannot be represented faithfully by the shared Drawing renderer.");
            } else if (gradientProjection == ShapeFillGradientProjection.None
                && TryResolveShapeFillColor(source, colorScheme, out OfficeColor fill)) {
                target.FillColor = fill;
            }

            if (TryResolveShapeOutlineColor(source, colorScheme, out OfficeColor stroke)) {
                A.Outline? localOutline = GetLocalShapeOutline(source);
                A.Outline? themeOutline = GetThemeShapeOutline(source);
                target.StrokeColor = stroke;
                long? widthEmus = localOutline?.Width?.Value
                    ?? themeOutline?.Width?.Value;
                double widthPoints = widthEmus.HasValue
                    ? widthEmus.Value / 12700D
                    : 1D;
                target.StrokeWidth = mapping.MapStrokeWidth(widthPoints);
                A.PresetLineDashValues? dash = localOutline?
                    .GetFirstChild<A.PresetDash>()?.Val?.Value;
                if (!dash.HasValue
                    && localOutline?.GetFirstChild<A.CustomDash>() == null) {
                    dash = themeOutline?.GetFirstChild<A.PresetDash>()?
                        .Val?.Value;
                }
                target.StrokeDashStyle = MapDash(dash);
                ApplyStrokeLineCapAndJoin(target, localOutline, themeOutline);
                ApplyLineMarkers(target, localOutline, themeOutline);
            } else {
                target.StrokeWidth = 0D;
            }

            ApplyShapeEffects(target, source, colorScheme, mapping);
        }

        private static A.Outline? GetLocalShapeOutline(PowerPointShape source) =>
            GetOpenXmlShapeProperties(source)?.GetFirstChild<A.Outline>();

        private static A.Outline? GetThemeShapeOutline(PowerPointShape source) {
            A.LineReference? lineReference =
                GetOpenXmlShapeStyle(source)?.LineReference;
            A.FormatScheme? formatScheme = source.OwnerSlide == null
                ? null
                : GetSlideFormatScheme(source.OwnerSlide);
            OpenXmlElement? themeLine = ResolveThemeShapeLine(formatScheme,
                lineReference?.Index?.Value);
            return themeLine as A.Outline
                ?? themeLine?.GetFirstChild<A.Outline>();
        }

        private static bool TryResolveShapeFillColor(PowerPointShape source, A.ColorScheme? colorScheme, out OfficeColor color) {
            OfficeColor? resolved = source.ResolveFillColorForExport();
            color = resolved.GetValueOrDefault();
            return resolved.HasValue;
        }

        private static bool TryResolveShapeOutlineColor(PowerPointShape source, A.ColorScheme? colorScheme, out OfficeColor color) {
            OfficeColor? resolved = source.ResolveOutlineColorForExport();
            color = resolved.GetValueOrDefault();
            return resolved.HasValue;
        }

        private static void ApplyStrokeLineCapAndJoin(OfficeShape target,
            A.Outline? localOutline, A.Outline? themeOutline) {
            target.StrokeLineCap = MapLineCap(localOutline?.CapType?.Value)
                ?? MapLineCap(themeOutline?.CapType?.Value);
            target.StrokeLineJoin = localOutline == null
                ? null
                : MapLineJoin(localOutline);
            target.StrokeLineJoin ??= themeOutline == null
                ? null
                : MapLineJoin(themeOutline);
        }

        private static OfficeStrokeLineCap? MapLineCap(A.LineCapValues? cap) {
            if (!cap.HasValue) {
                return null;
            }

            if (cap.Value == A.LineCapValues.Round) {
                return OfficeStrokeLineCap.Round;
            }

            if (cap.Value == A.LineCapValues.Square) {
                return OfficeStrokeLineCap.Square;
            }

            if (cap.Value == A.LineCapValues.Flat) {
                return OfficeStrokeLineCap.Butt;
            }

            return null;
        }

        private static OfficeStrokeLineJoin? MapLineJoin(A.Outline outline) {
            if (outline.GetFirstChild<A.Round>() != null) {
                return OfficeStrokeLineJoin.Round;
            }

            if (outline.GetFirstChild<A.LineJoinBevel>() != null) {
                return OfficeStrokeLineJoin.Bevel;
            }

            if (outline.GetFirstChild<A.Miter>() != null) {
                return OfficeStrokeLineJoin.Miter;
            }

            return null;
        }

        private static void ApplyLineMarkers(OfficeShape target,
            A.Outline? localOutline, A.Outline? themeOutline) {
            if (target.Kind != OfficeShapeKind.Line && target.Kind != OfficeShapeKind.Path) {
                return;
            }

            A.HeadEnd? headEnd = localOutline?.GetFirstChild<A.HeadEnd>()
                ?? themeOutline?.GetFirstChild<A.HeadEnd>();
            A.TailEnd? tailEnd = localOutline?.GetFirstChild<A.TailEnd>()
                ?? themeOutline?.GetFirstChild<A.TailEnd>();
            target.StrokeStartMarker = MapLineMarker(headEnd?.Type?.Value, headEnd?.Width?.Value, headEnd?.Length?.Value, target.StrokeWidth);
            target.StrokeEndMarker = MapLineMarker(tailEnd?.Type?.Value, tailEnd?.Width?.Value, tailEnd?.Length?.Value, target.StrokeWidth);
        }

        private static OfficeLineMarker? MapLineMarker(A.LineEndValues? type, A.LineEndWidthValues? width, A.LineEndLengthValues? length, double strokeWidth) {
            OfficeLineMarkerKind kind = MapLineMarkerKind(type);
            if (kind == OfficeLineMarkerKind.None) {
                return null;
            }

            double markerWidth = Math.Max(1D, strokeWidth * MapLineMarkerWidthFactor(width));
            double markerLength = Math.Max(1D, strokeWidth * MapLineMarkerLengthFactor(length));
            return new OfficeLineMarker(kind, markerWidth, markerLength);
        }

        private static DocumentFormat.OpenXml.Presentation.ShapeProperties? GetOpenXmlShapeProperties(PowerPointShape shape) {
            return shape.Element switch {
                DocumentFormat.OpenXml.Presentation.Shape openXmlShape => openXmlShape.ShapeProperties,
                DocumentFormat.OpenXml.Presentation.ConnectionShape connectionShape => connectionShape.ShapeProperties,
                _ => null
            };
        }

        private static OfficeLineMarkerKind MapLineMarkerKind(A.LineEndValues? type) {
            if (type == A.LineEndValues.Triangle) {
                return OfficeLineMarkerKind.Triangle;
            }

            if (type == A.LineEndValues.Stealth) {
                return OfficeLineMarkerKind.Stealth;
            }

            if (type == A.LineEndValues.Diamond) {
                return OfficeLineMarkerKind.Diamond;
            }

            if (type == A.LineEndValues.Oval) {
                return OfficeLineMarkerKind.Oval;
            }

            if (type == A.LineEndValues.Arrow) {
                return OfficeLineMarkerKind.Arrow;
            }

            return OfficeLineMarkerKind.None;
        }

        private static double MapLineMarkerWidthFactor(A.LineEndWidthValues? width) {
            if (width == A.LineEndWidthValues.Small) {
                return 3D;
            }

            if (width == A.LineEndWidthValues.Large) {
                return 6D;
            }

            return 4.5D;
        }

        private static double MapLineMarkerLengthFactor(A.LineEndLengthValues? length) {
            if (length == A.LineEndLengthValues.Small) {
                return 4D;
            }

            if (length == A.LineEndLengthValues.Large) {
                return 8D;
            }

            return 6D;
        }

        private static void ApplyShapeTransform(OfficeShape target, PowerPointShape source, double width, double height) {
            double rotation = source.Rotation ?? 0D;
            bool flipHorizontal = source.HorizontalFlip == true;
            bool flipVertical = source.VerticalFlip == true;
            var transform = new OfficeImageFrameTransform(
                rotation,
                width / 2D,
                height / 2D,
                flipHorizontal,
                flipVertical);
            if (transform.HasTransform) {
                target.Transform = transform.CreateDestinationTransform();
            }
        }

        private static bool ShouldRenderRichText(IReadOnlyList<OfficeRichTextRun> richRuns) =>
            richRuns.Count > 1 ||
            (richRuns.Count == 1 && (richRuns[0].Underline || richRuns[0].Strikethrough ||
                richRuns[0].LinkDestinationName != null || richRuns[0].LinkUri != null ||
                richRuns[0].BackgroundColor.HasValue || richRuns[0].Baseline != OfficeTextBaseline.Normal));

        private static OfficeTextParagraphIndent CreateParagraphIndent(PowerPointParagraph? paragraph, PowerPointShapeBoundsMapping mapping) {
            if (paragraph == null) {
                return OfficeTextParagraphIndent.Empty;
            }

            double leftMargin = Math.Max(0D, mapping.MapHorizontalLength(paragraph.LeftMarginPoints ?? 0D));
            double firstLine = Math.Max(0D, leftMargin + mapping.MapHorizontalLength(paragraph.IndentPoints ?? 0D));
            return firstLine > 0D || leftMargin > 0D
                ? new OfficeTextParagraphIndent(firstLine, leftMargin)
                : OfficeTextParagraphIndent.Empty;
        }

        private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTextBox textBox, PowerPointParagraph? paragraph, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping, bool markerRun = false, string? defaultFontFamily = null) {
            OfficeColor color = ResolveTextRunColor(run, textBox, colorScheme, paragraph);
            OfficeColor? backgroundColor = ResolveTextRunBackgroundColor(run, textBox, colorScheme);
            return new OfficeRichTextRun(
                ResolvePowerPointDisplayText(text, run, paragraph, markerRun),
                mapping.MapFontSize(markerRun ? paragraph?.BulletSizePoints ?? run?.FontSize ?? textBox.FontSize ?? 18 : run?.FontSize ?? textBox.FontSize ?? 18),
                color,
                run?.Bold == true,
                run?.Italic == true,
                run?.Underline == true,
                (markerRun ? paragraph?.BulletFontName : null) ?? run?.FontName ?? textBox.FontName ?? ResolveTextBoxFallbackFont(textBox, defaultFontFamily),
                run?.Strikethrough == true,
                backgroundColor,
                MapUnderlineStyle(run?.UnderlineStyle),
                MapStrikeStyle(run?.StrikeStyle),
                MapBaseline(ResolvePowerPointBaselinePercent(run, paragraph)));
        }

        private static string ResolvePowerPointDisplayText(string text, PowerPointTextRun? run, PowerPointParagraph? paragraph, bool markerRun = false) {
            PowerPointCapitalization? capitalization = ResolvePowerPointCapitalization(run, paragraph);
            if (markerRun || capitalization is not (PowerPointCapitalization.AllCaps or PowerPointCapitalization.SmallCaps)) {
                return text;
            }

            return OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Uppercase,
                ResolvePowerPointRunCulture(ResolvePowerPointLanguage(run, paragraph)));
        }

        private static string ResolvePowerPointDisplayText(PowerPointTextBox textBox) {
            if (!textBox.Paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => node.Run != null &&
                    ResolvePowerPointCapitalization(node.Run, paragraph) is PowerPointCapitalization.AllCaps or PowerPointCapitalization.SmallCaps))) {
                return textBox.Text;
            }

            return string.Join(Environment.NewLine, textBox.Paragraphs.Select(paragraph =>
                string.Concat(paragraph.InlineNodes.Select(node =>
                    node.Run == null ? node.Text : ResolvePowerPointDisplayText(node.Text, node.Run, paragraph)))));
        }

        private static PowerPointCapitalization? ResolvePowerPointCapitalization(PowerPointTextRun? run, PowerPointParagraph? paragraph) {
            if (run?.Capitalization.HasValue == true) return run.Capitalization;
            A.TextCapsValues? inherited = paragraph?.Paragraph.ParagraphProperties?
                .GetFirstChild<A.DefaultRunProperties>()?.Capital?.Value;
            if (!inherited.HasValue) return null;
            if (inherited.Value == A.TextCapsValues.None) return PowerPointCapitalization.None;
            if (inherited.Value == A.TextCapsValues.Small) return PowerPointCapitalization.SmallCaps;
            if (inherited.Value == A.TextCapsValues.All) return PowerPointCapitalization.AllCaps;
            return null;
        }

        private static string? ResolvePowerPointLanguage(PowerPointTextRun? run, PowerPointParagraph? paragraph) =>
            !string.IsNullOrWhiteSpace(run?.Language)
                ? run!.Language
                : paragraph?.Paragraph.ParagraphProperties?
                    .GetFirstChild<A.DefaultRunProperties>()?.Language?.Value;

        private static double? ResolvePowerPointBaselinePercent(PowerPointTextRun? run, PowerPointParagraph? paragraph) {
            if (run?.BaselinePercent.HasValue == true) return run.BaselinePercent;
            int? inherited = paragraph?.Paragraph.ParagraphProperties?
                .GetFirstChild<A.DefaultRunProperties>()?.Baseline?.Value;
            return inherited.HasValue ? inherited.Value / 1000D : (double?)null;
        }

        private static CultureInfo ResolvePowerPointRunCulture(string? language) {
            if (!string.IsNullOrWhiteSpace(language)) {
                try {
                    return CultureInfo.GetCultureInfo(language!);
                } catch (CultureNotFoundException) {
                    // Invalid or unsupported tags fall back deterministically for image export.
                }
            }

            return CultureInfo.InvariantCulture;
        }

        private static void AddSmallCapsApproximationDiagnosticIfNeeded(
            PowerPointTextBox textBox,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (!textBox.Paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => node.Run != null &&
                    ResolvePowerPointCapitalization(node.Run, paragraph) == PowerPointCapitalization.SmallCaps && !string.IsNullOrEmpty(node.Text)))) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                PowerPointImageExportDiagnosticCodes.SmallCapsApproximated,
                "Rendered native PowerPoint small caps as uppercase glyphs; reduced small-cap glyph sizing is approximated.",
                DescribeShape(textBox),
                OfficeConversionLossKind.Approximation));
        }

        private static void AddBaselineApproximationDiagnosticIfNeeded(
            PowerPointTextBox textBox,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (!textBox.Paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => node.Run != null &&
                    IsPowerPointBaselineApproximated(ResolvePowerPointBaselinePercent(node.Run, paragraph))))) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                PowerPointImageExportDiagnosticCodes.BaselinePercentApproximated,
                "Rendered a native PowerPoint baseline percentage using the shared superscript or subscript geometry; the exact authored percentage is not representable.",
                DescribeShape(textBox),
                OfficeConversionLossKind.Approximation));
        }

        private static bool IsPowerPointBaselineApproximated(double? baselinePercent) =>
            baselinePercent.HasValue && baselinePercent.Value != 0D &&
            Math.Abs(baselinePercent.Value - 30D) > 0.0001D &&
            Math.Abs(baselinePercent.Value + 25D) > 0.0001D;

        private static void AddWavyDoubleUnderlineApproximationDiagnosticIfNeeded(
            PowerPointTextBox textBox,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (!textBox.Paragraphs.SelectMany(paragraph => paragraph.InlineNodes)
                .Any(node => node.Run?.UnderlineStyle == PowerPointUnderlineStyle.WavyDouble && !string.IsNullOrEmpty(node.Text))) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                PowerPointImageExportDiagnosticCodes.WavyDoubleUnderlineApproximated,
                "Rendered native PowerPoint double-wavy underline as a double solid underline because the shared Drawing model cannot represent the compound pattern.",
                DescribeShape(textBox),
                OfficeConversionLossKind.Approximation));
        }

        private static OfficeTextDecorationStyle MapUnderlineStyle(PowerPointUnderlineStyle? style) => style switch {
            null or PowerPointUnderlineStyle.None => OfficeTextDecorationStyle.None,
            PowerPointUnderlineStyle.Double or PowerPointUnderlineStyle.WavyDouble => OfficeTextDecorationStyle.Double,
            PowerPointUnderlineStyle.Dotted or PowerPointUnderlineStyle.HeavyDotted => OfficeTextDecorationStyle.Dotted,
            PowerPointUnderlineStyle.Dash or PowerPointUnderlineStyle.DashHeavy or
                PowerPointUnderlineStyle.DashLong or PowerPointUnderlineStyle.DashLongHeavy or
                PowerPointUnderlineStyle.DotDash or PowerPointUnderlineStyle.DotDashHeavy or
                PowerPointUnderlineStyle.DotDotDash or PowerPointUnderlineStyle.DotDotDashHeavy => OfficeTextDecorationStyle.Dashed,
            PowerPointUnderlineStyle.Wavy or PowerPointUnderlineStyle.WavyHeavy => OfficeTextDecorationStyle.Wavy,
            _ => OfficeTextDecorationStyle.Single
        };

        private static OfficeTextDecorationStyle MapStrikeStyle(PowerPointStrikeStyle? style) => style switch {
            PowerPointStrikeStyle.Double => OfficeTextDecorationStyle.Double,
            PowerPointStrikeStyle.Single => OfficeTextDecorationStyle.Single,
            _ => OfficeTextDecorationStyle.None
        };

        private static OfficeTextBaseline MapBaseline(double? baselinePercent) => baselinePercent switch {
            > 0D => OfficeTextBaseline.Superscript,
            < 0D => OfficeTextBaseline.Subscript,
            _ => OfficeTextBaseline.Normal
        };

        private static OfficeColor ResolveTextRunColor(PowerPointTextRun? run, PowerPointTextBox textBox, A.ColorScheme? colorScheme,
            PowerPointParagraph? paragraph = null) {
            if (run == null) return ResolveTextBoxColor(textBox, colorScheme);
            paragraph ??= textBox.Paragraphs.FirstOrDefault(item =>
                ReferenceEquals(item.Paragraph, run.Run.Parent ?? run.RunProperties?.Parent?.Parent));
            return PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(run, paragraph, textBox)
                ?? OfficeColor.Black;
        }

        private static OfficeColor ResolveTextBoxColor(PowerPointTextBox textBox, A.ColorScheme? colorScheme) {
            PowerPointParagraph? paragraph = textBox.Paragraphs.FirstOrDefault(item =>
                item.InlineNodes.Any(node => node.Run != null && !string.IsNullOrEmpty(node.Text)))
                ?? textBox.Paragraphs.FirstOrDefault();
            PowerPointTextRun? run = paragraph?.InlineNodes
                .FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run;
            return PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(run, paragraph, textBox)
                ?? OfficeColor.Black;
        }

        private static OfficeColor? ResolveTextRunBackgroundColor(PowerPointTextRun? run, PowerPointTextBox textBox, A.ColorScheme? colorScheme) =>
            PowerPointEffectiveRunStyleResolver.ResolveTextColorElementForExport(
                run?.RunProperties?.GetFirstChild<A.Highlight>(), textBox);

        private static OfficeTextAlignment MapTextAlignment(PowerPointTextAlignment? alignment) {
            if (alignment == PowerPointTextAlignment.Center) {
                return OfficeTextAlignment.Center;
            }

            if (alignment == PowerPointTextAlignment.Right) {
                return OfficeTextAlignment.Right;
            }

            if (alignment == PowerPointTextAlignment.Justified ||
                alignment == PowerPointTextAlignment.Distributed ||
                alignment == PowerPointTextAlignment.ThaiDistributed ||
                alignment == PowerPointTextAlignment.JustifiedLow) {
                return OfficeTextAlignment.Justify;
            }

            return OfficeTextAlignment.Left;
        }

        private static OfficeStrokeDashStyle MapDash(A.PresetLineDashValues? dash) {
            if (dash == A.PresetLineDashValues.LargeDashDotDot ||
                dash == A.PresetLineDashValues.SystemDashDotDot) {
                return OfficeStrokeDashStyle.DashDotDot;
            }

            if (dash == A.PresetLineDashValues.DashDot ||
                dash == A.PresetLineDashValues.LargeDashDot ||
                dash == A.PresetLineDashValues.SystemDashDot) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (dash == A.PresetLineDashValues.Dot ||
                dash == A.PresetLineDashValues.SystemDot) {
                return OfficeStrokeDashStyle.Dot;
            }

            if (dash == A.PresetLineDashValues.Dash ||
                dash == A.PresetLineDashValues.LargeDash ||
                dash == A.PresetLineDashValues.SystemDash) {
                return OfficeStrokeDashStyle.Dash;
            }

            return OfficeStrokeDashStyle.Solid;
        }

        private static bool TryParseOfficeColor(string? value, out OfficeColor color) {
            color = default;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string normalized = value!.Trim();
            if (normalized.StartsWith("#", StringComparison.Ordinal)) {
                normalized = normalized.Substring(1);
            }

            return OfficeColor.TryParseHex(normalized, out color);
        }

        private static string DescribeShape(PowerPointShape shape) =>
            string.IsNullOrWhiteSpace(shape.Name) ? "PowerPoint shape" : "PowerPoint shape '" + shape.Name + "'";

        private static void AddUnsupportedShapeDiagnostic(
            List<OfficeImageExportDiagnostic> diagnostics,
            PowerPointShape shape,
            string message,
            OfficeConversionLossKind lossKind = OfficeConversionLossKind.Omission) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                PowerPointImageExportDiagnosticCodes.UnsupportedShape,
                message,
                DescribeShape(shape),
                lossKind));
        }

        private readonly struct PowerPointShapeBoundsMapping {
            internal static PowerPointShapeBoundsMapping Identity { get; } = new PowerPointShapeBoundsMapping(0D, 0D, 1D, 1D);

            private readonly double _offsetX;
            private readonly double _offsetY;
            private readonly double _scaleX;
            private readonly double _scaleY;
            private readonly bool _hasTransformedAncestor;

            internal PowerPointShapeBoundsMapping(double offsetX, double offsetY,
                double scaleX, double scaleY,
                bool hasTransformedAncestor = false) {
                _offsetX = offsetX;
                _offsetY = offsetY;
                _scaleX = scaleX;
                _scaleY = scaleY;
                _hasTransformedAncestor = hasTransformedAncestor;
            }

            internal bool HasTransformedAncestor => _hasTransformedAncestor;

            internal double MapX(double value) => _offsetX + (value * _scaleX);

            internal double MapY(double value) => _offsetY + (value * _scaleY);

            internal double MapWidth(double value) => value * _scaleX;

            internal double MapHeight(double value) => value * _scaleY;

            internal double MapHorizontalLength(double value) => value * Math.Abs(_scaleX);

            internal double MapVerticalLength(double value) => value * Math.Abs(_scaleY);

            internal double MapFontSize(double value) => value * Math.Abs(_scaleY);

            internal double MapStrokeWidth(double value) => value * Math.Sqrt(Math.Abs(_scaleX * _scaleY));

            internal PowerPointShapeBoundsMapping WithTransformedAncestor(bool value) =>
                new PowerPointShapeBoundsMapping(_offsetX, _offsetY, _scaleX,
                    _scaleY, _hasTransformedAncestor || value);

            internal PowerPointShapeBoundsMapping Compose(PowerPointShapeBoundsMapping child) =>
                new PowerPointShapeBoundsMapping(
                    _offsetX + (child._offsetX * _scaleX),
                    _offsetY + (child._offsetY * _scaleY),
                    _scaleX * child._scaleX,
                    _scaleY * child._scaleY,
                    _hasTransformedAncestor || child._hasTransformedAncestor);
        }
    }
}
