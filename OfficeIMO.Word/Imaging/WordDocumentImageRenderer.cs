using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double TwipsPerPoint = 20D;
        private const double DefaultPageWidthPoints = 595.3D;
        private const double DefaultPageHeightPoints = 841.9D;
        private const double DefaultMarginPoints = 72D;
        private const double ParagraphGapPoints = 6D;
        private const double DefaultCellMarginPoints = 5.4D;
        private const double MinimumTableRowHeightPoints = 22D;

        internal static OfficeImageExportResult Render(WordDocument document, OfficeImageExportFormat format, WordImageExportOptions options) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            WordDocumentVisualSnapshot snapshot = CreateSnapshot(document, options);
            OfficeDrawing drawing = snapshot.Drawing;

            if (format == OfficeImageExportFormat.Svg) {
                List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
                AddSvgImageDiagnostics(drawing, diagnostics);
                byte[] svg = OfficeDrawingSvgExporter.ToSvgBytes(drawing);
                return new OfficeImageExportResult(format, UnscaledWidth(drawing), UnscaledHeight(drawing), svg, "Page " + (options.PageIndex + 1), "Word document", diagnostics);
            }

            if (format == OfficeImageExportFormat.Png) {
                List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
                AddRasterImageDiagnostics(drawing, diagnostics);
                OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, options.Scale);
                return new OfficeImageExportResult(format, image.Width, image.Height, OfficePngWriter.Encode(image), "Page " + (options.PageIndex + 1), "Word document", diagnostics);
            }

            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
        }

        internal static WordDocumentVisualSnapshot CreateSnapshot(WordDocument document, WordImageExportOptions options) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeDrawing drawing = CreateDrawing(document, options, diagnostics);
            return new WordDocumentVisualSnapshot(drawing, options.PageIndex, diagnostics.AsReadOnly());
        }

        private static OfficeDrawing CreateDrawing(WordDocument document, WordImageExportOptions options, List<OfficeImageExportDiagnostic> diagnostics) {
            (double width, double height) = GetPageSizePoints(document);
            OfficeDrawing drawing = new OfficeDrawing(width, height);
            AddBackgroundRectangle(drawing, options.BackgroundColor);

            if (options.PageIndex > 0) {
                AddDiagnostic(diagnostics, "unsupported-word-page-index", "Rendered a blank page because dependency-free Word image export currently supports only the first page.");
                return drawing;
            }

            if (options.IncludeDocumentContent) {
                AddSupportedHeaderFooterContent(document, drawing, diagnostics);
                AddSupportedBodyContent(document, drawing, diagnostics);
            }

            return drawing;
        }

        private static void AddSupportedBodyContent(WordDocument document, OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics) {
            WordImageFlowContext context = CreateFlowContext(document, drawing);
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            foreach (OpenXmlElement element in document.BodyRoot.ChildElements) {
                bool added = false;
                if (element is Paragraph paragraph) {
                    added = AddParagraphContent(document, paragraph, context, diagnostics, listMarkers);
                } else if (element is Table table) {
                    context.ClearParagraphSpacingState();
                    added = AddTable(new WordTable(document, table), context, diagnostics, listMarkers);
                    context.ClearParagraphSpacingState();
                } else if (element is SectionProperties) {
                    continue;
                } else {
                    context.ClearParagraphSpacingState();
                    AddDiagnostic(diagnostics, "unsupported-word-body-element", "Skipped a Word body element that is not yet projected through OfficeIMO.Drawing.", element.GetType().Name);
                }

                if (context.StoppedForPagination) {
                    break;
                }

                if (!added && element is Paragraph) {
                    context.ClearParagraphSpacingState();
                    context.Y += ParagraphGapPoints;
                }
            }
        }

        private static WordImageFlowContext CreateFlowContext(WordDocument document, OfficeDrawing drawing) {
            WordMargins margins = document.Margins;
            double left = ToPoints(margins.Left?.Value, DefaultMarginPoints);
            double right = ToPoints(margins.Right?.Value, DefaultMarginPoints);
            double top = ToPoints(margins.Top, DefaultMarginPoints);
            double bottom = ToPoints(margins.Bottom, DefaultMarginPoints);
            double contentWidth = Math.Max(1D, drawing.Width - left - right);
            double contentBottom = Math.Max(top, drawing.Height - bottom);
            return new WordImageFlowContext(drawing, left, top, contentWidth, contentBottom);
        }

        private static WordImageFlowContext CreateFlowContext(
            OfficeDrawing drawing,
            double left,
            double y,
            double contentWidth,
            double contentBottom,
            string overflowDiagnosticCode,
            string overflowDiagnosticMessage) =>
            new WordImageFlowContext(drawing, left, y, contentWidth, contentBottom, overflowDiagnosticCode, overflowDiagnosticMessage);

        private static bool AddParagraphContent(
            WordDocument document,
            Paragraph paragraph,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            bool added = false;
            var colorScheme = GetDocumentColorScheme(document);
            WordImageListMarker? listMarker = CreateListMarker(document, paragraph, listMarkers);
            bool markerRendered = false;
            var textRuns = new List<WordParagraph>();

            bool FlushTextRuns() {
                if (textRuns.Count == 0) {
                    return false;
                }

                WordImageListMarker? currentMarker = markerRendered ? null : listMarker;
                bool runAdded = textRuns.Count == 1 && !HasRunHighlight(textRuns[0])
                    ? AddTextRun(textRuns[0], context, diagnostics, currentMarker, colorScheme)
                    : AddRichTextRuns(textRuns, context, diagnostics, currentMarker, colorScheme);
                if (runAdded && currentMarker.HasValue) {
                    markerRendered = true;
                }

                textRuns.Clear();
                return runAdded;
            }

            foreach (WordParagraph run in WordSection.ConvertParagraphToWordParagraphs(document, paragraph)) {
                WordImage? image = run.Image;
                if (image != null) {
                    added |= FlushTextRuns();
                    context.ClearParagraphSpacingState();
                    added |= AddImage(image, context, diagnostics);
                    context.ClearParagraphSpacingState();
                    continue;
                }

                WordShape? shape = run.Shape;
                if (shape != null) {
                    added |= FlushTextRuns();
                    context.ClearParagraphSpacingState();
                    added |= AddShape(shape, context, diagnostics);
                    context.ClearParagraphSpacingState();
                    continue;
                }

                WordTextBox? textBox = run.TextBox;
                if (textBox != null) {
                    added |= FlushTextRuns();
                    context.ClearParagraphSpacingState();
                    added |= AddTextBox(textBox, context, diagnostics, colorScheme);
                    context.ClearParagraphSpacingState();
                    continue;
                }

                string text = run.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                textRuns.Add(run);
            }

            added |= FlushTextRuns();
            return added;
        }

        private static bool AddRichTextRuns(IReadOnlyList<WordParagraph> paragraphs, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics, WordImageListMarker? listMarker, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme) {
            List<OfficeRichTextRun> richRuns = CreateRichTextRuns(paragraphs, colorScheme);
            if (richRuns.Count == 0) {
                return false;
            }

            double maxFontSize = 10D;
            for (int i = 0; i < richRuns.Count; i++) {
                maxFontSize = Math.Max(maxFontSize, richRuns[i].FontSize);
            }

            double lineHeight = Math.Max(maxFontSize * 1.25D, 12D);
            WordImageTextLayout textLayout = ResolveTextLayout(context, listMarker, paragraphs[0]);
            double height = EstimateRichTextHeight(richRuns, maxFontSize, textLayout.ContentWidth, lineHeight, textLayout.ParagraphIndent);
            WordParagraphSpacing spacing = ResolveParagraphSpacing(paragraphs[0], maxFontSize, lineHeight, context, out WordParagraphSpacingState spacingState);
            if (!EnsureVerticalSpace(context, spacing.Before + height, diagnostics)) {
                return false;
            }

            context.Y += spacing.Before;
            AddParagraphFrame(paragraphs[0], context, textLayout, height, colorScheme);
            if (listMarker.HasValue) {
                WordImageListMarker marker = listMarker.Value;
                context.Drawing.AddText(
                    marker.Marker,
                    textLayout.MarkerLeft,
                    context.Y,
                    textLayout.MarkerWidth,
                    lineHeight,
                    marker.Font,
                    marker.Color,
                    marker.Alignment,
                    lineHeight,
                    wrapText: false);
            }

            context.Drawing.AddRichText(
                richRuns,
                textLayout.TextLeft,
                context.Y,
                textLayout.TextWidth,
                height,
                MapTextAlignment(paragraphs[0].ParagraphAlignment),
                lineHeight,
                wrapText: true,
                padding: textLayout.Padding,
                paragraphIndent: textLayout.ParagraphIndent);
            context.Y += height + spacing.After;
            context.SetParagraphSpacingState(spacingState);
            return true;
        }

        private static bool AddTextRun(WordParagraph paragraph, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics, WordImageListMarker? listMarker, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme) {
            OfficeFontInfo font = CreateFont(paragraph);
            double lineHeight = Math.Max(font.Size * 1.25D, 12D);
            WordImageTextLayout textLayout = ResolveTextLayout(context, listMarker, paragraph);
            double height = EstimateTextHeight(paragraph.Text, font.Size, textLayout.LayoutWidth, lineHeight);
            WordParagraphSpacing spacing = ResolveParagraphSpacing(paragraph, font.Size, lineHeight, context, out WordParagraphSpacingState spacingState);
            if (!EnsureVerticalSpace(context, spacing.Before + height, diagnostics)) {
                return false;
            }

            context.Y += spacing.Before;
            AddParagraphFrame(paragraph, context, textLayout, height, colorScheme);
            if (listMarker.HasValue) {
                WordImageListMarker marker = listMarker.Value;
                context.Drawing.AddText(
                    marker.Marker,
                    textLayout.MarkerLeft,
                    context.Y,
                    textLayout.MarkerWidth,
                    lineHeight,
                    marker.Font,
                    marker.Color,
                    marker.Alignment,
                    lineHeight,
                    wrapText: false);
            }

            context.Drawing.AddText(
                paragraph.Text,
                textLayout.TextLeft,
                context.Y,
                textLayout.TextWidth,
                height,
                font,
                ResolveParagraphTextColor(paragraph, colorScheme),
                MapTextAlignment(paragraph.ParagraphAlignment),
                lineHeight,
                wrapText: true,
                padding: textLayout.Padding,
                paragraphIndent: textLayout.ParagraphIndent);
            context.Y += height + spacing.After;
            context.SetParagraphSpacingState(spacingState);
            return true;
        }

        private static List<OfficeRichTextRun> CreateRichTextRuns(IReadOnlyList<WordParagraph> paragraphs, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme) {
            var richRuns = new List<OfficeRichTextRun>(paragraphs.Count);
            for (int i = 0; i < paragraphs.Count; i++) {
                WordParagraph paragraph = paragraphs[i];
                string text = paragraph.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                richRuns.Add(CreateRichTextRun(paragraph, colorScheme, text));
            }

            return richRuns;
        }

        private static OfficeRichTextRun CreateRichTextRun(WordParagraph paragraph, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme, string? text = null) {
            double fontSize = paragraph.FontSize ?? 11;
            return new OfficeRichTextRun(
                text ?? paragraph.Text,
                fontSize,
                ResolveParagraphTextColor(paragraph, colorScheme),
                paragraph.Bold,
                paragraph.Italic,
                paragraph.Underline.HasValue && paragraph.Underline.Value != UnderlineValues.None,
                paragraph.FontFamily ?? "Calibri",
                paragraph.Strike || paragraph.DoubleStrike,
                ResolveRunHighlightColor(ResolveRunHighlight(paragraph)));
        }

        private static OfficeFontInfo CreateFont(WordParagraph paragraph) {
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (paragraph.Bold) {
                style |= OfficeFontStyle.Bold;
            }

            if (paragraph.Italic) {
                style |= OfficeFontStyle.Italic;
            }

            if (paragraph.Underline.HasValue && paragraph.Underline.Value != UnderlineValues.None) {
                style |= OfficeFontStyle.Underline;
            }

            if (paragraph.Strike || paragraph.DoubleStrike) {
                style |= OfficeFontStyle.Strikethrough;
            }

            return new OfficeFontInfo(paragraph.FontFamily ?? "Calibri", paragraph.FontSize ?? 11, style);
        }

        private static OfficeTextAlignment MapTextAlignment(JustificationValues? alignment) {
            if (alignment == JustificationValues.Center) {
                return OfficeTextAlignment.Center;
            }

            if (alignment == JustificationValues.Right || alignment == JustificationValues.End) {
                return OfficeTextAlignment.Right;
            }

            if (alignment == JustificationValues.Both ||
                alignment == JustificationValues.Distribute ||
                alignment == JustificationValues.ThaiDistribute ||
                alignment == JustificationValues.HighKashida ||
                alignment == JustificationValues.MediumKashida ||
                alignment == JustificationValues.LowKashida) {
                return OfficeTextAlignment.Justify;
            }

            return OfficeTextAlignment.Left;
        }

        private static double EstimateTextHeight(string text, double fontSize, double contentWidth, double lineHeight) {
            double averageCharacterWidth = Math.Max(1D, fontSize * 0.52D);
            int charactersPerLine = Math.Max(1, (int)Math.Floor(contentWidth / averageCharacterWidth));
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            string[] explicitLines = normalized.Split('\n');
            int lineCount = 0;
            foreach (string line in explicitLines) {
                lineCount += Math.Max(1, (int)Math.Ceiling(line.Length / (double)charactersPerLine));
            }

            return Math.Max(lineHeight, lineCount * lineHeight);
        }

        private static double EstimateRichTextHeight(IReadOnlyList<OfficeRichTextRun> runs, double maxFontSize, double contentWidth, double lineHeight, OfficeTextParagraphIndent paragraphIndent) {
            double lineHeightFactor = Math.Max(1D, lineHeight / Math.Max(1D, maxFontSize));
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create();
            Func<string?, double, string?, double> measure = (value, size, family) => {
                    OfficeTextMeasurementStyle measuredStyle = measurer.CreateStyle(new OfficeFontInfo(family, size));
                    return measurer.MeasureWidth(value, measuredStyle);
                };
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                runs,
                contentWidth,
                double.MaxValue,
                lineHeightFactor,
                measure,
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: Math.Min(6D, maxFontSize),
                overflowBehavior: OfficeTextOverflowBehavior.Clip,
                paragraphIndent: paragraphIndent);
            return Math.Max(lineHeight, layout.Height);
        }

        private static void AddBackgroundRectangle(OfficeDrawing drawing, OfficeColor fillColor) {
            OfficeShape shape = OfficeShape.Rectangle(drawing.Width, drawing.Height);
            shape.FillColor = fillColor;
            shape.StrokeWidth = 0D;
            drawing.AddShape(shape, 0D, 0D);
        }

        private static (double Width, double Height) GetPageSizePoints(WordDocument document) {
            WordPageSizes pageSettings = document.PageSettings;
            double width = ToPoints(pageSettings.Width?.Value, DefaultPageWidthPoints);
            double height = ToPoints(pageSettings.Height?.Value, DefaultPageHeightPoints);
            return (Math.Max(1D, width), Math.Max(1D, height));
        }

        private static double ToPoints(uint? twips, double fallbackPoints) =>
            twips.HasValue ? twips.Value / TwipsPerPoint : fallbackPoints;

        private static double ToPoints(int? twips, double fallbackPoints) =>
            twips.HasValue ? twips.Value / TwipsPerPoint : fallbackPoints;

        private static double ToPoints(short? twips, double fallbackPoints) =>
            twips.HasValue ? twips.Value / TwipsPerPoint : fallbackPoints;

        private static bool EnsureVerticalSpace(WordImageFlowContext context, double height, List<OfficeImageExportDiagnostic> diagnostics) {
            if (context.Y + height <= context.ContentBottom) {
                return true;
            }

            if (!context.StoppedForPagination) {
                AddDiagnostic(diagnostics, context.OverflowDiagnosticCode, context.OverflowDiagnosticMessage);
                context.StoppedForPagination = true;
            }

            return false;
        }

        private static int ScaledWidth(OfficeDrawing drawing, WordImageExportOptions options) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Width * options.Scale));

        private static int ScaledHeight(OfficeDrawing drawing, WordImageExportOptions options) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Height * options.Scale));

        private static int UnscaledWidth(OfficeDrawing drawing) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Width));

        private static int UnscaledHeight(OfficeDrawing drawing) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Height));

        private static void AddDiagnostic(List<OfficeImageExportDiagnostic> diagnostics, string code, string message, string? source = null) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                code,
                message,
                string.IsNullOrWhiteSpace(source) ? "Word document" : source));
        }

        private sealed class WordImageFlowContext {
            internal WordImageFlowContext(
                OfficeDrawing drawing,
                double left,
                double top,
                double contentWidth,
                double contentBottom,
                string overflowDiagnosticCode = "unsupported-word-pagination",
                string overflowDiagnosticMessage = "Stopped rendering Word body content after the first page because dependency-free pagination is not implemented yet.") {
                Drawing = drawing;
                Left = left;
                Y = top;
                ContentWidth = contentWidth;
                ContentBottom = contentBottom;
                OverflowDiagnosticCode = overflowDiagnosticCode;
                OverflowDiagnosticMessage = overflowDiagnosticMessage;
            }

            internal OfficeDrawing Drawing { get; }

            internal double Left { get; }

            internal double Y { get; set; }

            internal double ContentWidth { get; }

            internal double ContentBottom { get; }

            internal bool StoppedForPagination { get; set; }

            internal string OverflowDiagnosticCode { get; }

            internal string OverflowDiagnosticMessage { get; }

            internal WordParagraphSpacingState? PreviousParagraphSpacingState { get; private set; }

            private List<WordTextExclusion>? TextExclusions { get; set; }

            internal void AddTextExclusion(double left, double top, double right, double bottom) =>
                AddTextExclusion(left, top, right, bottom, WordTextWrapSide.Largest);

            internal void AddTextExclusion(double left, double top, double right, double bottom, WordTextWrapSide wrapSide) {
                if (right <= left || bottom <= top) {
                    return;
                }

                TextExclusions ??= new List<WordTextExclusion>();
                TextExclusions.Add(new WordTextExclusion(left, top, right, bottom, wrapSide));
            }

            internal WordTextFlowFrame ResolveTextFlowFrame() {
                double contentRight = Left + ContentWidth;
                if (TextExclusions == null || TextExclusions.Count == 0) {
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                for (int index = TextExclusions.Count - 1; index >= 0; index--) {
                    if (Y >= TextExclusions[index].Bottom) {
                        TextExclusions.RemoveAt(index);
                    }
                }

                if (TextExclusions.Count == 0) {
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                WordTextExclusion? active = null;
                for (int index = 0; index < TextExclusions.Count; index++) {
                    WordTextExclusion exclusion = TextExclusions[index];
                    if (Y >= exclusion.Top && Y < exclusion.Bottom) {
                        active = exclusion;
                        break;
                    }
                }

                if (!active.HasValue) {
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                WordTextExclusion current = active.Value;
                double leftWidth = Math.Max(0D, current.Left - Left);
                double rightLeft = Math.Min(contentRight, current.Right);
                double rightWidth = Math.Max(0D, contentRight - rightLeft);

                if (current.WrapSide == WordTextWrapSide.Left) {
                    if (leftWidth >= 1D) {
                        return new WordTextFlowFrame(Left, Math.Max(1D, leftWidth));
                    }

                    Y = current.Bottom + ParagraphGapPoints;
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                if (current.WrapSide == WordTextWrapSide.Right) {
                    if (rightWidth >= 1D) {
                        return new WordTextFlowFrame(rightLeft, Math.Max(1D, rightWidth));
                    }

                    Y = current.Bottom + ParagraphGapPoints;
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                if (leftWidth < 1D && rightWidth < 1D) {
                    Y = current.Bottom + ParagraphGapPoints;
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                return rightWidth >= leftWidth
                    ? new WordTextFlowFrame(rightLeft, Math.Max(1D, rightWidth))
                    : new WordTextFlowFrame(Left, Math.Max(1D, leftWidth));
            }

            internal void SetParagraphSpacingState(WordParagraphSpacingState state) =>
                PreviousParagraphSpacingState = state;

            internal void ClearParagraphSpacingState() =>
                PreviousParagraphSpacingState = null;
        }

        private readonly struct WordTextExclusion {
            internal WordTextExclusion(double left, double top, double right, double bottom, WordTextWrapSide wrapSide) {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
                WrapSide = wrapSide;
            }

            internal double Left { get; }

            internal double Top { get; }

            internal double Right { get; }

            internal double Bottom { get; }

            internal WordTextWrapSide WrapSide { get; }
        }

        private enum WordTextWrapSide {
            Largest,
            Left,
            Right
        }

        private readonly struct WordTextFlowFrame {
            internal WordTextFlowFrame(double left, double width) {
                Left = left;
                Width = width;
            }

            internal double Left { get; }

            internal double Width { get; }
        }
    }
}
