using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private const int D = OfficeXmlNamespaces.DrawingML;
        private const int Pml = OfficeXmlNamespaces.PresentationML;

        private static OfficeXmlNode GetSlideColorScheme(PowerPointSlide slide) =>
            PowerPointThemeScope.For(slide.PartRef).ColorScheme;

        private static OfficeXmlNode GetSlideFormatScheme(PowerPointSlide slide) =>
            PowerPointThemeScope.For(slide.PartRef).FormatScheme;

        /// <summary>The color map the slide's scheme colors go through, or none without a slide.</summary>
        private static OfficeXmlNode GetSlideColorMap(PowerPointSlide? slide) =>
            slide == null ? default : PowerPointThemeScope.For(slide.PartRef).ColorMap;

        /// <summary>The color map of the slide a part ref names, or none for no slide.</summary>
        private static OfficeXmlNode GetSlideColorMap(PowerPointPartRef slide) =>
            slide.IsNone ? default : PowerPointThemeScope.For(slide).ColorMap;

        /// <summary>Projects the existing shared shape appearance for preset and text-frame export.</summary>
        internal static void ApplyShapeAppearanceForExport(OfficeShape target, PowerPointShape source,
            List<OfficeImageExportDiagnostic> diagnostics) {
            bool transformedAncestor = false;
            for (OfficeXmlNode ancestor = source.XmlElement.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
                if (!ancestor.Is(Pml, "grpSp")) continue;
                OfficeXmlNode transform = ancestor.Child(Pml, "grpSpPr").Child(D, "xfrm");
                if (!transform.IsNone && ((transform.Int32Attribute("rot") ?? 0) != 0 || transform.BooleanAttribute("flipH") == true
                        || transform.BooleanAttribute("flipV") == true)) {
                    transformedAncestor = true;
                    break;
                }
            }
            ApplyShapeStyle(target, source, source.OwnerSlide == null ? default : GetSlideColorScheme(source.OwnerSlide),
                PowerPointShapeBoundsMapping.Identity.WithTransformedAncestor(transformedAncestor), diagnostics);
            if (source.UnsupportedFillKindForExport is { } fillKind) {
                AddUnsupportedShapeDiagnostic(diagnostics, source,
                    "Skipped a PowerPoint " + fillKind + " fill that is not supported by the shared Drawing renderer.");
            }
        }

        private static void ApplyShapeStyle(OfficeShape target, PowerPointShape source,
            OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping,
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
                OfficeXmlNode localOutline = GetLocalShapeOutline(source);
                OfficeXmlNode themeOutline = GetThemeShapeOutline(source);
                target.StrokeColor = stroke;
                long? widthEmus = localOutline.Int32Attribute("w")
                    ?? themeOutline.Int32Attribute("w");
                double widthPoints = widthEmus.HasValue
                    ? widthEmus.Value / 12700D
                    : 1D;
                // Group extents scale shape coordinates; the existing slide export contract
                // retains authored outline widths for shapes and connectors.
                target.StrokeWidth = widthPoints;
                ApplyStrokeDash(target, source, localOutline, themeOutline, diagnostics);
                ApplyStrokeLineCapAndJoin(target, localOutline, themeOutline);
                ApplyLineMarkers(target, localOutline, themeOutline);
            } else {
                target.StrokeWidth = 0D;
            }

            ApplyShapeEffects(target, source, colorScheme, mapping);
        }

        private static OfficeXmlNode GetLocalShapeOutline(PowerPointShape source) =>
            GetShapePropertiesElement(source).Child(D, "ln");

        private static OfficeXmlNode GetThemeShapeOutline(PowerPointShape source) {
            OfficeXmlNode lineReference = GetShapeStyleElement(source).Child(D, "lnRef");
            OfficeXmlNode formatScheme = source.OwnerSlide == null
                ? default
                : GetSlideFormatScheme(source.OwnerSlide);
            OfficeXmlNode themeLine = ResolveThemeShapeLine(formatScheme,
                lineReference.UInt32Attribute("idx"));
            return themeLine.Is(D, "ln") ? themeLine : themeLine.Child(D, "ln");
        }

        private static bool TryResolveShapeFillColor(PowerPointShape source, OfficeXmlNode colorScheme, out OfficeColor color) {
            OfficeColor? resolved = source.ResolveFillColorForExport();
            color = resolved.GetValueOrDefault();
            return resolved.HasValue;
        }

        private static bool TryResolveShapeOutlineColor(PowerPointShape source, OfficeXmlNode colorScheme, out OfficeColor color) {
            OfficeColor? resolved = source.ResolveOutlineColorForExport();
            color = resolved.GetValueOrDefault();
            return resolved.HasValue;
        }

        /// <summary>Converts DrawingML dash percentages to exact lengths in the authored stroke width.</summary>
        private static void ApplyStrokeDash(OfficeShape target, PowerPointShape source,
            OfficeXmlNode localOutline, OfficeXmlNode themeOutline, List<OfficeImageExportDiagnostic> diagnostics) {
            OfficeXmlNode outline = !localOutline.Child(D, "prstDash").IsNone
                || !localOutline.Child(D, "custDash").IsNone ? localOutline : themeOutline;
            target.StrokeDashStyle = MapDash(outline.Child(D, "prstDash").Attribute("val"));
            OfficeXmlNode custom = outline.Child(D, "custDash");
            if (custom.IsNone || target.StrokeWidth <= 0D) return;

            var pattern = new List<double>();
            foreach (OfficeXmlNode stop in custom.Children(D, "ds")) {
                pattern.Add(target.StrokeWidth * (stop.Int32Attribute("d") ?? 0) / 100000D);
                pattern.Add(target.StrokeWidth * (stop.Int32Attribute("sp") ?? 0) / 100000D);
            }
            if (pattern.Any(value => value < 0D) || !pattern.Any(value => value > 0D)) {
                AddUnsupportedShapeDiagnostic(diagnostics, source,
                    "Rendered a PowerPoint custom dash pattern without valid positive lengths as a solid outline.",
                    OfficeConversionLossKind.Approximation);
                return;
            }
            target.SetStrokeDashArray(pattern);
        }

        private static void ApplyStrokeLineCapAndJoin(OfficeShape target,
            OfficeXmlNode localOutline, OfficeXmlNode themeOutline) {
            target.StrokeLineCap = MapLineCap(localOutline.Attribute("cap"))
                ?? MapLineCap(themeOutline.Attribute("cap"));
            target.StrokeLineJoin = localOutline.IsNone
                ? null
                : MapLineJoin(localOutline);
            target.StrokeLineJoin ??= themeOutline.IsNone
                ? null
                : MapLineJoin(themeOutline);
        }

        // DrawingML ST_LineCap tokens.
        private static OfficeStrokeLineCap? MapLineCap(string? cap) => cap switch {
            "rnd" => OfficeStrokeLineCap.Round,
            "sq" => OfficeStrokeLineCap.Square,
            "flat" => OfficeStrokeLineCap.Butt,
            _ => null
        };

        private static OfficeStrokeLineJoin? MapLineJoin(OfficeXmlNode outline) {
            if (!outline.Child(D, "round").IsNone) {
                return OfficeStrokeLineJoin.Round;
            }

            if (!outline.Child(D, "bevel").IsNone) {
                return OfficeStrokeLineJoin.Bevel;
            }

            if (!outline.Child(D, "miter").IsNone) {
                return OfficeStrokeLineJoin.Miter;
            }

            return null;
        }

        private static void ApplyLineMarkers(OfficeShape target,
            OfficeXmlNode localOutline, OfficeXmlNode themeOutline) {
            if (target.Kind != OfficeShapeKind.Line && target.Kind != OfficeShapeKind.Path) {
                return;
            }

            OfficeXmlNode headEnd = localOutline.Child(D, "headEnd");
            if (headEnd.IsNone) headEnd = themeOutline.Child(D, "headEnd");
            OfficeXmlNode tailEnd = localOutline.Child(D, "tailEnd");
            if (tailEnd.IsNone) tailEnd = themeOutline.Child(D, "tailEnd");
            target.StrokeStartMarker = MapLineMarker(headEnd.Attribute("type"), headEnd.Attribute("w"), headEnd.Attribute("len"), target.StrokeWidth);
            target.StrokeEndMarker = MapLineMarker(tailEnd.Attribute("type"), tailEnd.Attribute("w"), tailEnd.Attribute("len"), target.StrokeWidth);
        }

        private static OfficeLineMarker? MapLineMarker(string? type, string? width, string? length, double strokeWidth) {
            OfficeLineMarkerKind kind = MapLineMarkerKind(type);
            if (kind == OfficeLineMarkerKind.None) {
                return null;
            }

            double markerWidth = Math.Max(1D, strokeWidth * MapLineMarkerWidthFactor(width));
            double markerLength = Math.Max(1D, strokeWidth * MapLineMarkerLengthFactor(length));
            return new OfficeLineMarker(kind, markerWidth, markerLength);
        }

        /// <summary>The <c>p:spPr</c> of a shape or connector, or none for other elements.</summary>
        private static OfficeXmlNode GetShapePropertiesElement(PowerPointShape shape) {
            OfficeXmlNode element = shape.XmlElement;
            return element.Is(Pml, "sp") || element.Is(Pml, "cxnSp") ? element.Child(Pml, "spPr") : default;
        }

        // DrawingML ST_LineEndType, ST_LineEndWidth and ST_LineEndLength tokens.
        private static OfficeLineMarkerKind MapLineMarkerKind(string? type) => type switch {
            "triangle" => OfficeLineMarkerKind.Triangle,
            "stealth" => OfficeLineMarkerKind.Stealth,
            "diamond" => OfficeLineMarkerKind.Diamond,
            "oval" => OfficeLineMarkerKind.Oval,
            "arrow" => OfficeLineMarkerKind.Arrow,
            _ => OfficeLineMarkerKind.None
        };

        private static double MapLineMarkerWidthFactor(string? width) => width switch {
            "sm" => 3D,
            "lg" => 6D,
            _ => 4.5D
        };

        private static double MapLineMarkerLengthFactor(string? length) => length switch {
            "sm" => 4D,
            "lg" => 8D,
            _ => 6D
        };

        /// <summary>Applies the native shape frame to the composed geometry and paint.</summary>
        internal static void ApplyShapeTransformForExport(OfficeShape target, PowerPointShape source, double width, double height) {
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

            double leftMargin = Math.Max(0D, mapping.MapHorizontalLength(paragraph.LeftMargin?.ToPoints() ?? 0D));
            double firstLine = Math.Max(0D, leftMargin + mapping.MapHorizontalLength(paragraph.Indent?.ToPoints() ?? 0D));
            return firstLine > 0D || leftMargin > 0D
                ? new OfficeTextParagraphIndent(firstLine, leftMargin)
                : OfficeTextParagraphIndent.Empty;
        }

        private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTextBox textBox, PowerPointParagraph? paragraph, OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping, bool markerRun = false, string? defaultFontFamily = null) {
            OfficeColor color = ResolveTextRunColor(run, textBox, colorScheme, paragraph);
            if (markerRun && paragraph != null)
                color = PowerPointEffectiveRunStyleResolver.ResolveBulletColorForExport(paragraph, textBox, color);
            OfficeColor? backgroundColor = ResolveTextRunBackgroundColor(run, textBox, colorScheme, paragraph);
            return new OfficeRichTextRun(
                ResolvePowerPointDisplayText(text, run, paragraph, markerRun),
                mapping.MapFontSize(markerRun ? paragraph?.BulletSize?.ToPoints() ?? run?.FontSize?.ToPoints() ?? textBox.FontSize?.ToPoints() ?? 18 : run?.FontSize?.ToPoints() ?? textBox.FontSize?.ToPoints() ?? 18),
                color,
                run?.Bold == true,
                run?.Italic == true,
                run?.Underline == true,
                ResolveThemeFontReference(
                    (markerRun ? paragraph?.BulletFontFamily : null) ?? run?.FontFamily ?? textBox.FontFamily,
                    run, paragraph, textBox)
                    ?? ResolveTextBoxFallbackFont(textBox, defaultFontFamily),
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
            if (paragraph == null) return null;
            // DrawingML ST_TextCapsType tokens.
            return ParagraphDefaultRunProperties(paragraph).Attribute("cap") switch {
                "none" => PowerPointCapitalization.None,
                "small" => PowerPointCapitalization.SmallCaps,
                "all" => PowerPointCapitalization.AllCaps,
                _ => null
            };
        }

        /// <summary>The paragraph's <c>a:pPr/a:defRPr</c>, or none.</summary>
        private static OfficeXmlNode ParagraphDefaultRunProperties(PowerPointParagraph paragraph) =>
            paragraph.Element.Child(D, "pPr").Child(D, "defRPr");

        private static string? ResolvePowerPointLanguage(PowerPointTextRun? run, PowerPointParagraph? paragraph) =>
            !string.IsNullOrWhiteSpace(run?.Language)
                ? run!.Language
                : paragraph == null ? null : ParagraphDefaultRunProperties(paragraph).Attribute("lang");

        private static double? ResolvePowerPointBaselinePercent(PowerPointTextRun? run, PowerPointParagraph? paragraph) {
            if (run?.BaselinePercent.HasValue == true) return run.BaselinePercent;
            int? inherited = paragraph == null ? null : ParagraphDefaultRunProperties(paragraph).Int32Attribute("baseline");
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

        private static OfficeColor ResolveTextRunColor(PowerPointTextRun? run, PowerPointTextBox textBox, OfficeXmlNode colorScheme,
            PowerPointParagraph? paragraph = null) {
            if (run == null) return ResolveTextBoxColor(textBox, colorScheme);
            paragraph ??= FindRunParagraph(textBox.Paragraphs, run);
            return PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(run, paragraph, textBox)
                ?? OfficeColor.Black;
        }

        /// <summary>
        /// The paragraph of <paramref name="paragraphs"/> that holds the run: the parent of an <c>a:r</c>, or of an
        /// <c>a:fld</c> that has run properties; a field without them, or a run that is not attached, has none.
        /// </summary>
        private static PowerPointParagraph? FindRunParagraph(IEnumerable<PowerPointParagraph> paragraphs, PowerPointTextRun run) {
            OfficeXmlNode element = run.Element;
            OfficeXmlNode parent = element.Is(D, "r") || element.Is(D, "fld") && !element.Child(D, "rPr").IsNone
                ? element.Parent
                : default;
            if (parent.IsNone) return null;
            return paragraphs.FirstOrDefault(item => item.Element == parent);
        }

        private static OfficeColor ResolveTextBoxColor(PowerPointTextBox textBox, OfficeXmlNode colorScheme) {
            PowerPointParagraph? paragraph = textBox.Paragraphs.FirstOrDefault(item =>
                item.InlineNodes.Any(node => node.Run != null && !string.IsNullOrEmpty(node.Text)))
                ?? textBox.Paragraphs.FirstOrDefault();
            PowerPointTextRun? run = paragraph?.InlineNodes
                .FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run;
            return PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(run, paragraph, textBox)
                ?? OfficeColor.Black;
        }

        private static OfficeColor? ResolveTextRunBackgroundColor(PowerPointTextRun? run, PowerPointTextBox textBox,
            OfficeXmlNode colorScheme, PowerPointParagraph? paragraph = null) {
            if (run != null) paragraph ??= FindRunParagraph(textBox.Paragraphs, run);
            return ResolveTextBoxHighlightColor(run, paragraph, textBox);
        }

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

        // DrawingML ST_PresetLineDashVal tokens.
        private static OfficeStrokeDashStyle MapDash(string? dash) => dash switch {
            "lgDashDotDot" or "sysDashDotDot" => OfficeStrokeDashStyle.DashDotDot,
            "dashDot" or "lgDashDot" or "sysDashDot" => OfficeStrokeDashStyle.DashDot,
            "dot" or "sysDot" => OfficeStrokeDashStyle.Dot,
            "dash" or "lgDash" or "sysDash" => OfficeStrokeDashStyle.Dash,
            _ => OfficeStrokeDashStyle.Solid
        };

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
            OfficeConversionLossKind lossKind = OfficeConversionLossKind.Omission, string? diagnosticCode = null) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                diagnosticCode ?? PowerPointImageExportDiagnosticCodes.UnsupportedShape,
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
