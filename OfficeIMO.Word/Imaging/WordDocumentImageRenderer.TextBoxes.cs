using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double DefaultTextBoxWidthPoints = 180D;
        private const double DefaultTextBoxHeightPoints = 72D;

        private static bool AddTextBox(WordTextBox textBox, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics, A.ColorScheme? colorScheme) {
            List<WordParagraph> paragraphs = textBox.Paragraphs;
            string text = GetTextBoxText(textBox, paragraphs, context);
            if (string.IsNullOrWhiteSpace(text)) {
                text = string.Empty;
            }

            WordParagraph? firstParagraph = paragraphs.FirstOrDefault(paragraph => !string.IsNullOrEmpty(paragraph.Text));
            OfficeFontInfo font = firstParagraph == null ? OfficeFontInfo.Default : CreateFont(firstParagraph);
            double lineHeight = Math.Max(font.Size * 1.25D, 12D);
            OfficeTextPadding padding = GetTextBoxPadding(textBox);
            if (!TryGetTextBoxSize(textBox, text, font.Size, lineHeight, padding, out double width, out double height)) {
                AddDiagnostic(diagnostics, "unsupported-word-textbox", "Skipped a Word text box because its size could not be resolved.", "Word text box");
                return false;
            }

            Anchor? anchor = textBox.Anchor;
            if (anchor != null) {
                return AddAnchoredTextBox(textBox, anchor, text, firstParagraph, font, lineHeight, padding, width, height, context, diagnostics, colorScheme);
            }

            width = Math.Min(width, context.ContentWidth);
            if (!EnsureVerticalSpace(context, height, diagnostics)) {
                return false;
            }

            if (context.IsTargetPage) {
                AddTextBoxDrawing(textBox, text, firstParagraph, font, lineHeight, padding, context.Left, context.Y, width, height, context, colorScheme);
            }

            context.Y += height + ParagraphGapPoints;
            return true;
        }

        private static bool AddAnchoredTextBox(
            WordTextBox textBox,
            Anchor anchor,
            string text,
            WordParagraph? firstParagraph,
            OfficeFontInfo font,
            double lineHeight,
            OfficeTextPadding padding,
            double width,
            double height,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            A.ColorScheme? colorScheme) {
            width = Math.Min(width, context.ContentWidth);
            double left = ResolveHorizontalAnchorPosition(anchor.HorizontalPosition, context, width);
            double top = ResolveVerticalAnchorPosition(anchor.VerticalPosition, context, height);
            if (!IsFinite(left) || !IsFinite(top)) {
                if (context.IsTargetPage) {
                    AddDiagnostic(diagnostics, "unsupported-word-textbox", "Skipped a Word text box because its anchor position could not be resolved.", "Word text box");
                }

                return false;
            }

            if (anchor.GetFirstChild<WrapTopBottom>() != null) {
                return AddTopAndBottomAnchoredTextBox(textBox, anchor, text, firstParagraph, font, lineHeight, padding, width, height, context, diagnostics, colorScheme);
            }

            double right = left + width;
            double bottom = top + height;
            if (left < 0D || top < 0D || right > context.Drawing.Width || bottom > context.Drawing.Height) {
                if (context.IsTargetPage) {
                    AddDiagnostic(diagnostics, "unsupported-word-textbox", "Skipped a Word text box because its anchor projects outside the current page preview.", "Word text box");
                }

                return false;
            }

            WordTextBoxFrameTransform transform = GetTextBoxFrameTransform(textBox);
            if (context.IsTargetPage) {
                AddTextBoxDrawing(textBox, text, firstParagraph, font, lineHeight, padding, left, top, width, height, context, colorScheme);
            }

            bool hasSquareWrap = anchor.GetFirstChild<WrapSquare>() != null;
            bool hasTightWrap = anchor.GetFirstChild<WrapTight>() != null;
            bool hasThroughWrap = anchor.GetFirstChild<WrapThrough>() != null;
            bool usedAuthoredPolygon = false;
            bool usedFramePolygon = false;
            if (hasSquareWrap || hasTightWrap || hasThroughWrap) {
                double exclusionLeft = Math.Max(context.Left, left - GetAnchorDistancePoints(anchor.DistanceFromLeft));
                double exclusionTop = Math.Max(0D, top - GetAnchorDistancePoints(anchor.DistanceFromTop));
                double exclusionRight = Math.Min(context.Left + context.ContentWidth, right + GetAnchorDistancePoints(anchor.DistanceFromRight));
                double exclusionBottom = Math.Min(context.ContentBottom, bottom + GetAnchorDistancePoints(anchor.DistanceFromBottom));
                WordTextWrapSide wrapSide = GetTextBoxWrapSide(anchor);
                IReadOnlyList<OfficePoint> polygon = Array.Empty<OfficePoint>();
                usedAuthoredPolygon = (hasTightWrap || hasThroughWrap) &&
                    TryCreateAuthoredWrapPolygonTextExclusion(anchor, exclusionLeft, exclusionTop, exclusionRight, exclusionBottom, out polygon);
                if (usedAuthoredPolygon) {
                    context.AddTextExclusion(polygon, wrapSide);
                } else if ((hasTightWrap || hasThroughWrap) &&
                    !transform.HasTransform &&
                    TryCreateTextBoxFrameTextExclusion(exclusionLeft, exclusionTop, exclusionRight, exclusionBottom, out polygon)) {
                    context.AddTextExclusion(polygon, wrapSide);
                    usedFramePolygon = true;
                } else {
                    context.AddTextExclusion(exclusionLeft, exclusionTop, exclusionRight, exclusionBottom, wrapSide);
                }
            }

            if (context.IsTargetPage && (hasTightWrap || hasThroughWrap) && !usedAuthoredPolygon && !usedFramePolygon) {
                AddDiagnostic(
                    diagnostics,
                    "limited-word-floating-textbox-wrap",
                    "Rendered a Word text box with a rectangular text exclusion because dependency-free export does not yet implement polygon wrapping.",
                    "Word text box");
            }

            if (hasSquareWrap || hasTightWrap || hasThroughWrap) {
                AdvanceFlowToAnchoredWrapTop(context, top);
            }

            return true;
        }

        private static bool AddTopAndBottomAnchoredTextBox(
            WordTextBox textBox,
            Anchor anchor,
            string text,
            WordParagraph? firstParagraph,
            OfficeFontInfo font,
            double lineHeight,
            OfficeTextPadding padding,
            double width,
            double height,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            A.ColorScheme? colorScheme) {
            double left = ResolveHorizontalAnchorPosition(anchor.HorizontalPosition, context, width);
            double top = ResolveVerticalAnchorPosition(anchor.VerticalPosition, context, height);
            if (!IsFinite(left) || !IsFinite(top)) {
                if (context.IsTargetPage) {
                    AddDiagnostic(diagnostics, "unsupported-word-textbox", "Skipped a Word top-and-bottom text box because its anchor position could not be resolved.", "Word text box");
                }

                return false;
            }

            top = Math.Max(top, context.Y);
            double distanceFromBottom = GetAnchorDistancePoints(anchor.DistanceFromBottom);
            double requiredHeight = top + height + distanceFromBottom - context.Y;
            if (!EnsureVerticalSpace(context, requiredHeight, diagnostics)) {
                return false;
            }

            left = ResolveHorizontalAnchorPosition(anchor.HorizontalPosition, context, width);
            top = Math.Max(ResolveVerticalAnchorPosition(anchor.VerticalPosition, context, height), context.Y);
            if (!IsFinite(left) || !IsFinite(top)) {
                if (context.IsTargetPage) {
                    AddDiagnostic(diagnostics, "unsupported-word-textbox", "Skipped a Word top-and-bottom text box because its anchor position could not be resolved.", "Word text box");
                }

                return false;
            }

            double right = left + width;
            double bottom = top + height;
            if (left < 0D || top < 0D || right > context.Drawing.Width || bottom > context.Drawing.Height) {
                if (context.IsTargetPage) {
                    AddDiagnostic(diagnostics, "unsupported-word-textbox", "Skipped a Word top-and-bottom text box because its anchor projects outside the current page preview.", "Word text box");
                }

                return false;
            }

            if (context.IsTargetPage) {
                AddTextBoxDrawing(textBox, text, firstParagraph, font, lineHeight, padding, left, top, width, height, context, colorScheme);
            }

            context.Y = bottom + distanceFromBottom + ParagraphGapPoints;
            return true;
        }

        private static void AddTextBoxDrawing(
            WordTextBox textBox,
            string text,
            WordParagraph? firstParagraph,
            OfficeFontInfo font,
            double lineHeight,
            OfficeTextPadding padding,
            double left,
            double top,
            double width,
            double height,
            WordImageFlowContext context,
            A.ColorScheme? colorScheme) {
            OfficeShape frame = OfficeShape.Rectangle(width, height);
            ApplyTextBoxStyle(frame, textBox);
            WordTextBoxFrameTransform transform = GetTextBoxFrameTransform(textBox);
            if (transform.HasTransform) {
                frame.Transform = CreateLocalTextBoxFrameTransform(width, height, transform);
            }

            bool drawBehindContent = textBox.WrapText == WrapTextImage.BehindText;
            if (drawBehindContent) {
                context.Drawing.AddShapeBehindContent(frame, left, top);
            } else {
                context.Drawing.AddShape(frame, left, top);
            }

            double rotationCenterX = left + (width / 2D);
            double rotationCenterY = top + (height / 2D);

            List<OfficeRichTextRun> richRuns = CreateTextBoxRichTextRuns(textBox, colorScheme, context);
            if (ShouldRenderTextBoxAsRichText(textBox, richRuns)) {
                double maxFontSize = richRuns.Max(run => run.FontSize);
                double richLineHeight = Math.Max(maxFontSize * 1.25D, 12D);
                if (drawBehindContent) {
                    context.Drawing.AddRichTextBehindContent(
                        richRuns,
                        left,
                        top,
                        width,
                        height,
                        MapTextAlignment(firstParagraph?.ParagraphAlignment),
                        richLineHeight,
                        GetTextBoxVerticalAlignment(textBox),
                        rotationDegrees: transform.RotationDegrees,
                        rotationCenterX: rotationCenterX,
                        rotationCenterY: rotationCenterY,
                        wrapText: true,
                        flipHorizontal: transform.FlipHorizontal,
                        flipVertical: transform.FlipVertical,
                        padding: padding);
                    return;
                }

                context.Drawing.AddRichText(
                    richRuns,
                    left,
                    top,
                    width,
                    height,
                    MapTextAlignment(firstParagraph?.ParagraphAlignment),
                    richLineHeight,
                    GetTextBoxVerticalAlignment(textBox),
                    rotationDegrees: transform.RotationDegrees,
                    rotationCenterX: rotationCenterX,
                    rotationCenterY: rotationCenterY,
                    wrapText: true,
                    flipHorizontal: transform.FlipHorizontal,
                    flipVertical: transform.FlipVertical,
                    padding: padding);
                return;
            }

            if (drawBehindContent) {
                context.Drawing.AddTextBehindContent(
                    text,
                    left,
                    top,
                    width,
                    height,
                    font,
                    ResolveParagraphTextColor(firstParagraph, colorScheme),
                    MapTextAlignment(firstParagraph?.ParagraphAlignment),
                    lineHeight,
                    verticalAlignment: GetTextBoxVerticalAlignment(textBox),
                    rotationDegrees: transform.RotationDegrees,
                    rotationCenterX: rotationCenterX,
                    rotationCenterY: rotationCenterY,
                    wrapText: true,
                    flipHorizontal: transform.FlipHorizontal,
                    flipVertical: transform.FlipVertical,
                    padding: padding);
                return;
            }

            context.Drawing.AddText(
                text,
                left,
                top,
                width,
                height,
                font,
                ResolveParagraphTextColor(firstParagraph, colorScheme),
                MapTextAlignment(firstParagraph?.ParagraphAlignment),
                lineHeight,
                verticalAlignment: GetTextBoxVerticalAlignment(textBox),
                rotationDegrees: transform.RotationDegrees,
                rotationCenterX: rotationCenterX,
                rotationCenterY: rotationCenterY,
                wrapText: true,
                flipHorizontal: transform.FlipHorizontal,
                flipVertical: transform.FlipVertical,
                padding: padding);
        }

        private static bool TryGetTextBoxSize(
            WordTextBox textBox,
            string text,
            double fontSize,
            double lineHeight,
            OfficeTextPadding padding,
            out double width,
            out double height) {
            width = 0D;
            height = 0D;

            if (textBox.Anchor?.Extent?.Cx?.Value is long anchorCx &&
                textBox.Anchor.Extent.Cy?.Value is long anchorCy &&
                anchorCx > 0L &&
                anchorCy > 0L) {
                width = Helpers.ConvertEmusToPoints(anchorCx);
                height = Helpers.ConvertEmusToPoints(anchorCy);
                return width > 0D && height > 0D;
            }

            if (textBox.Inline?.Extent?.Cx?.Value is long inlineCx &&
                textBox.Inline.Extent.Cy?.Value is long inlineCy &&
                inlineCx > 0L &&
                inlineCy > 0L) {
                width = Helpers.ConvertEmusToPoints(inlineCx);
                height = Helpers.ConvertEmusToPoints(inlineCy);
                return width > 0D && height > 0D;
            }

            A.Extents? transformExtents = textBox.DrawingShapeProperties?
                .GetFirstChild<A.Transform2D>()?
                .GetFirstChild<A.Extents>();
            if (transformExtents?.Cx?.Value is long transformCx &&
                transformExtents.Cy?.Value is long transformCy &&
                transformCx > 0L &&
                transformCy > 0L) {
                width = Helpers.ConvertEmusToPoints(transformCx);
                height = Helpers.ConvertEmusToPoints(transformCy);
                return width > 0D && height > 0D;
            }

            if (TryGetVmlTextBoxSize(textBox.VmlShape, text, fontSize, lineHeight, padding, out width, out height)) {
                return true;
            }

            width = DefaultTextBoxWidthPoints;
            height = DefaultTextBoxHeightPoints;
            return true;
        }

        private static bool TryGetVmlTextBoxSize(V.Shape? shape, string text, double fontSize, double lineHeight, OfficeTextPadding padding, out double width, out double height) {
            width = 0D;
            height = 0D;
            string? style = shape?.Style?.Value;
            if (string.IsNullOrWhiteSpace(style)) {
                return false;
            }

            bool hasWidth = TryGetVmlStylePoints(style!, "width", out width);
            bool hasHeight = TryGetVmlStylePoints(style!, "height", out height);
            if (!hasWidth) {
                width = DefaultTextBoxWidthPoints;
            }

            if (!hasHeight && HasVmlStyleFlag(style!, "mso-fit-shape-to-text", "t")) {
                double contentWidth = Math.Max(1D, width - padding.Horizontal);
                height = EstimateTextHeight(text, fontSize, contentWidth, lineHeight) + padding.Vertical;
                hasHeight = true;
            }

            return width > 0D && hasHeight && height > 0D;
        }

        private static bool TryGetVmlStylePoints(string style, string name, out double points) {
            points = 0D;
            string prefix = name + ":";
            string[] parts = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts) {
                string text = part.Trim();
                if (!text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                string value = text.Substring(prefix.Length).Trim();
                if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) {
                    value = value.Substring(0, value.Length - 2);
                }

                return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out points) && IsFinite(points);
            }

            return false;
        }

        private static bool HasVmlStyleFlag(string style, string name, string expectedValue) {
            return TryGetVmlStyleText(style, name, out string value) &&
                string.Equals(value, expectedValue, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetTextBoxText(WordTextBox textBox, IEnumerable<WordParagraph> fallbackRuns, WordImageFlowContext? context = null) {
            DocumentFormat.OpenXml.Wordprocessing.TextBoxContent? content = textBox.Content;
            if (content != null) {
                List<string> paragraphText = content.ChildElements
                    .OfType<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                    .Select(paragraph => {
                        if (context?.ResolveDynamicPageFields != true) {
                            return NormalizeTextBoxParagraphText(paragraph.InnerText);
                        }

                        string resolvedText = string.Concat(
                            WordSection.ConvertParagraphToWordParagraphs(textBox.Document, paragraph, splitPaginationMarkers: true)
                                .Select(run => ResolveImageExportText(run, context)));
                        return string.IsNullOrEmpty(resolvedText)
                            ? NormalizeTextBoxParagraphText(paragraph.InnerText)
                            : NormalizeTextBoxParagraphText(resolvedText);
                    })
                    .Where(text => !string.IsNullOrEmpty(text))
                    .ToList();
                if (paragraphText.Count > 0) {
                    return string.Join(Environment.NewLine, paragraphText);
                }
            }

            List<string> parts = fallbackRuns
                .Select(paragraph => NormalizeTextBoxParagraphText(ResolveImageExportText(paragraph, context)))
                .Where(text => !string.IsNullOrEmpty(text))
                .ToList();
            return string.Join(Environment.NewLine, parts);
        }

        private static string NormalizeTextBoxParagraphText(string text) =>
            text.TrimEnd('\r', '\n');

        private static void AdvanceFlowToAnchoredWrapTop(WordImageFlowContext context, double top) {
            if (IsFinite(top) && IsFinite(context.ContentBottom) && context.ContentBottom < double.MaxValue / 2D && top > context.Y && top < context.ContentBottom) {
                context.Y = top;
            }
        }

        private static bool ShouldRenderTextBoxAsRichText(WordTextBox textBox, IReadOnlyList<OfficeRichTextRun> richRuns) {
            if (richRuns.Any(run => run.BackgroundColor.HasValue)) {
                return true;
            }

            if (ContainsTextBoxField(textBox)) {
                return richRuns.Count > 0;
            }

            OfficeRichTextRun? firstRun = richRuns.FirstOrDefault();
            return firstRun != null && richRuns.Any(run => !HasSameTextBoxRunStyle(firstRun, run));
        }

        private static bool ContainsTextBoxField(WordTextBox textBox) =>
            textBox.Content?.Descendants<W.FieldCode>().Any() == true ||
            textBox.Content?.Descendants<W.SimpleField>().Any() == true;

        private static bool HasSameTextBoxRunStyle(OfficeRichTextRun left, OfficeRichTextRun right) =>
            left.FontSize.Equals(right.FontSize) &&
            left.Color == right.Color &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.Underline == right.Underline &&
            left.Strikethrough == right.Strikethrough &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.Ordinal);

        private static List<OfficeRichTextRun> CreateTextBoxRichTextRuns(WordTextBox textBox, A.ColorScheme? colorScheme, WordImageFlowContext? context = null) {
            var richRuns = new List<OfficeRichTextRun>();
            DocumentFormat.OpenXml.Wordprocessing.TextBoxContent? content = textBox.Content;
            if (content == null) {
                return richRuns;
            }

            foreach (DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph in content.ChildElements.OfType<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()) {
                List<(WordParagraph Run, string Text)> paragraphRuns = WordSection.ConvertParagraphToWordParagraphs(textBox.Document, paragraph, splitPaginationMarkers: true)
                    .Select(run => (Run: run, Text: ResolveImageExportText(run, context)))
                    .Where(run => !string.IsNullOrEmpty(run.Text))
                    .ToList();
                if (paragraphRuns.Count == 0) {
                    continue;
                }

                if (richRuns.Count > 0) {
                    richRuns.Add(CreateRichTextRun(paragraphRuns[0].Run, colorScheme, Environment.NewLine));
                }

                for (int runIndex = 0; runIndex < paragraphRuns.Count; runIndex++) {
                    (WordParagraph run, string text) = paragraphRuns[runIndex];
                    richRuns.Add(CreateRichTextRun(run, colorScheme, text));
                }
            }

            return richRuns;
        }

        private static OfficeTextPadding GetTextBoxPadding(WordTextBox textBox) {
            var properties = textBox.DrawingTextBodyProperties;
            if (properties == null) {
                return TryGetVmlTextBoxPadding(textBox.VmlTextBox, out OfficeTextPadding padding)
                    ? padding
                    : new OfficeTextPadding(6D, 3D, 6D, 3D);
            }

            return new OfficeTextPadding(
                GetInsetPoints(properties.LeftInset, 6D),
                GetInsetPoints(properties.TopInset, 3D),
                GetInsetPoints(properties.RightInset, 6D),
                GetInsetPoints(properties.BottomInset, 3D));
        }

        private static OfficeTextVerticalAlignment GetTextBoxVerticalAlignment(WordTextBox textBox) {
            A.TextAnchoringTypeValues? drawingAnchor = textBox.DrawingTextBodyProperties?.Anchor?.Value;
            if (drawingAnchor == A.TextAnchoringTypeValues.Center) {
                return OfficeTextVerticalAlignment.Center;
            }

            if (drawingAnchor == A.TextAnchoringTypeValues.Bottom) {
                return OfficeTextVerticalAlignment.Bottom;
            }

            if (TryGetVmlStyleText(textBox.VmlShape?.Style?.Value, "v-text-anchor", out string vmlAnchor)) {
                if (string.Equals(vmlAnchor, "middle", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(vmlAnchor, "center", StringComparison.OrdinalIgnoreCase)) {
                    return OfficeTextVerticalAlignment.Center;
                }

                if (string.Equals(vmlAnchor, "bottom", StringComparison.OrdinalIgnoreCase)) {
                    return OfficeTextVerticalAlignment.Bottom;
                }
            }

            return OfficeTextVerticalAlignment.Top;
        }

        private static WordTextBoxFrameTransform GetTextBoxFrameTransform(WordTextBox textBox) {
            A.Transform2D? drawingTransform = textBox.DrawingShapeProperties?.GetFirstChild<A.Transform2D>();
            double rotation = drawingTransform?.Rotation?.Value is int rotationValue ? rotationValue / 60000D : 0D;
            bool flipHorizontal = drawingTransform?.HorizontalFlip?.Value == true;
            bool flipVertical = drawingTransform?.VerticalFlip?.Value == true;

            string? vmlStyle = textBox.VmlShape?.Style?.Value;
            if (!string.IsNullOrWhiteSpace(vmlStyle)) {
                if (TryGetVmlStyleText(vmlStyle, "rotation", out string vmlRotation) &&
                    double.TryParse(vmlRotation, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedRotation) &&
                    IsFinite(parsedRotation)) {
                    rotation = parsedRotation;
                }

                if (TryGetVmlStyleText(vmlStyle, "flip", out string vmlFlip)) {
                    flipHorizontal = vmlFlip.IndexOf('x') >= 0 || vmlFlip.IndexOf('X') >= 0;
                    flipVertical = vmlFlip.IndexOf('y') >= 0 || vmlFlip.IndexOf('Y') >= 0;
                }
            }

            return new WordTextBoxFrameTransform(rotation, flipHorizontal, flipVertical);
        }

        private static OfficeTransform CreateLocalTextBoxFrameTransform(double width, double height, WordTextBoxFrameTransform transform) {
            double centerX = width / 2D;
            double centerY = height / 2D;
            return OfficeTransform.Translate(-centerX, -centerY)
                .Then(OfficeTransform.Scale(transform.FlipHorizontal ? -1D : 1D, transform.FlipVertical ? -1D : 1D))
                .Then(OfficeTransform.RotateDegrees(transform.RotationDegrees))
                .Then(OfficeTransform.Translate(centerX, centerY));
        }

        private static double GetInsetPoints(DocumentFormat.OpenXml.Int32Value? inset, double fallback) =>
            inset?.Value != null ? Helpers.ConvertEmusToPoints(inset.Value) : fallback;

        private static bool TryGetVmlStyleText(string? style, string name, out string value) {
            value = string.Empty;
            if (string.IsNullOrWhiteSpace(style)) {
                return false;
            }

            string prefix = name + ":";
            string[] parts = style!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts) {
                string text = part.Trim();
                if (!text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                value = text.Substring(prefix.Length).Trim();
                return value.Length > 0;
            }

            return false;
        }


        private static bool TryGetVmlTextBoxPadding(V.TextBox? textBox, out OfficeTextPadding padding) {
            padding = OfficeTextPadding.Empty;
            string? inset = textBox?.Inset?.Value;
            if (string.IsNullOrWhiteSpace(inset)) {
                return false;
            }

            string[] parts = inset!.Split(',');
            double left = ParseVmlInsetPart(parts, 0, 6D);
            double top = ParseVmlInsetPart(parts, 1, 3D);
            double right = ParseVmlInsetPart(parts, 2, 6D);
            double bottom = ParseVmlInsetPart(parts, 3, 3D);
            padding = new OfficeTextPadding(left, top, right, bottom);
            return true;
        }

        private static double ParseVmlInsetPart(string[] parts, int index, double fallback) {
            if (index >= parts.Length) {
                return fallback;
            }

            string value = parts[index].Trim();
            if (string.IsNullOrWhiteSpace(value)) {
                return fallback;
            }

            return TryParseVmlLengthPoints(value, out double points) ? points : fallback;
        }

        private static bool TryParseVmlLengthPoints(string value, out double points) {
            points = 0D;
            string text = value.Trim();
            if (text.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) {
                text = text.Substring(0, text.Length - 2);
                return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out points) && IsFinite(points);
            }

            if (text.EndsWith("in", StringComparison.OrdinalIgnoreCase)) {
                text = text.Substring(0, text.Length - 2);
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double inches) && IsFinite(inches)) {
                    points = inches * 72D;
                    return true;
                }

                return false;
            }

            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out points) && IsFinite(points);
        }

        private static void ApplyTextBoxStyle(OfficeShape frame, WordTextBox textBox) {
            frame.FillColor = TryGetDrawingFillColor(textBox.DrawingShapeProperties, out OfficeColor fillColor)
                ? fillColor
                : TryGetVmlFillColor(textBox.VmlShape, out fillColor)
                    ? fillColor
                    : OfficeColor.White;
            frame.StrokeColor = TryGetDrawingOutlineColor(textBox.DrawingShapeProperties, out OfficeColor strokeColor)
                ? strokeColor
                : TryGetVmlStrokeColor(textBox.VmlShape, out strokeColor)
                    ? strokeColor
                    : OfficeColor.Black;
            frame.StrokeWidth = TryGetDrawingOutlineWidth(textBox.DrawingShapeProperties, out double strokeWidth)
                ? strokeWidth
                : TryGetVmlStrokeWidth(textBox.VmlShape, out strokeWidth)
                    ? strokeWidth
                    : 1D;

            if (textBox.VmlShape?.Stroked?.Value == false) {
                frame.StrokeColor = null;
                frame.StrokeWidth = 0D;
            }
        }

        private static bool TryGetDrawingFillColor(DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties? shapeProperties, out OfficeColor color) {
            color = OfficeColor.White;
            string? value = shapeProperties?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            return TryParseOfficeColor(value, out color);
        }

        private static bool TryGetDrawingOutlineColor(DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties? shapeProperties, out OfficeColor color) {
            color = OfficeColor.Black;
            string? value = shapeProperties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            return TryParseOfficeColor(value, out color);
        }

        private static bool TryGetDrawingOutlineWidth(DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties? shapeProperties, out double width) {
            width = 0D;
            if (shapeProperties?.GetFirstChild<A.Outline>()?.Width?.Value is int emus && emus > 0) {
                width = Helpers.ConvertEmusToPoints(emus);
                return width > 0D;
            }

            return false;
        }

        private static bool TryGetVmlFillColor(V.Shape? shape, out OfficeColor color) {
            color = OfficeColor.White;
            return TryParseOfficeColor(shape?.FillColor?.Value, out color);
        }

        private static bool TryGetVmlStrokeColor(V.Shape? shape, out OfficeColor color) {
            color = OfficeColor.Black;
            return TryParseOfficeColor(shape?.StrokeColor?.Value, out color);
        }

        private static bool TryGetVmlStrokeWidth(V.Shape? shape, out double width) {
            width = 0D;
            return TryParseVmlLengthPoints(shape?.StrokeWeight?.Value ?? string.Empty, out width) && width > 0D;
        }

        private static WordTextWrapSide GetTextBoxWrapSide(Anchor anchor) {
            WrapTextValues? wrapValue =
                anchor.Elements<WrapSquare>().FirstOrDefault()?.WrapText?.Value ??
                anchor.Elements<WrapTight>().FirstOrDefault()?.WrapText?.Value ??
                anchor.Elements<WrapThrough>().FirstOrDefault()?.WrapText?.Value;
            if (wrapValue == WrapTextValues.Left) {
                return WordTextWrapSide.Left;
            }

            if (wrapValue == WrapTextValues.Right) {
                return WordTextWrapSide.Right;
            }

            return WordTextWrapSide.Largest;
        }

        private static bool TryCreateTextBoxFrameTextExclusion(
            double left,
            double top,
            double right,
            double bottom,
            out IReadOnlyList<OfficePoint> polygon) {
            polygon = Array.Empty<OfficePoint>();
            if (right <= left || bottom <= top) {
                return false;
            }

            polygon = new[] {
                new OfficePoint(left, top),
                new OfficePoint(right, top),
                new OfficePoint(right, bottom),
                new OfficePoint(left, bottom)
            };
            return true;
        }

        private readonly struct WordTextBoxFrameTransform {
            internal WordTextBoxFrameTransform(double rotationDegrees, bool flipHorizontal, bool flipVertical) {
                RotationDegrees = rotationDegrees;
                FlipHorizontal = flipHorizontal;
                FlipVertical = flipVertical;
            }

            internal double RotationDegrees { get; }

            internal bool FlipHorizontal { get; }

            internal bool FlipVertical { get; }

            internal bool HasTransform => Math.Abs(RotationDegrees) > 0.000001D || FlipHorizontal || FlipVertical;
        }
    }
}
