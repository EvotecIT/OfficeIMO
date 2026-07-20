using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void WriteLinesInternal(string fontRes, double fontSize, double lineHeight, double x, double widthUsed, double startY, System.Collections.Generic.List<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false, string? structureType = null, int? markedContentId = null, PdfNamedFontFace? namedFont = null) {
            EnsurePage();
            pageDirty = true;
            AppendMarkedContentBegin(sb, structureType, markedContentId);
            var content = new ContentStreamBuilder(sb)
                .BeginText()
                .Font(fontRes, fontSize)
                .TextLeading(lineHeight);
            var lineFont = ResolveFontFromResourceName(fontRes, ChooseNormal(currentOpts.DefaultFont));
            double yStart2 = startY;
            if (applyBaselineTweak) {
                yStart2 -= GetDescenderForOptions(lineFont, namedFont, fontSize, currentOpts) * 0.0;
            }
            content.TextMatrix(x, yStart2);
            var effectiveColor = color ?? currentOpts.DefaultTextColor ?? PdfColor.Black;
            content.FillColor(effectiveColor);
            for (int i = 0; i < lines.Count; i++) {
                string line = lines[i];
                double dx = 0;
                double estWidth = EstimateSimpleTextWidthForOptions(line, lineFont, namedFont, fontSize, currentOpts);
                if (align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - estWidth) / 2);
                else if (align == PdfAlign.Right) dx = Math.Max(0, (widthUsed - estWidth));
                if (Math.Abs(dx) > 0.0001) content.MoveText(dx, 0);
                content.ShowText(EncodeTextShowCommand(line, lineFont, namedFont, currentOpts), fontSize);
                if (Math.Abs(dx) > 0.0001) content.MoveText(-dx, 0);
                if (i != lines.Count - 1) content.NextTextLine();
            }
            content.EndText();
            AppendMarkedContentEnd(sb, markedContentId);
        }

        private void WriteLines(string fontRes, double fontSize, double lineHeight, double x, double startY, System.Collections.Generic.List<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false, string? structureType = null, int? markedContentId = null)
            => WriteLinesInternal(fontRes, fontSize, lineHeight, x, width, startY, lines, align, color, applyBaselineTweak, structureType, markedContentId);

        private void AddHeadingLinkAnnotations(HeadingBlock heading, System.Collections.Generic.List<string> lines, PdfStandardFont font, double fontSize, double lineHeight, double x, double widthUsed, double startBaselineY, int? structElementIndex = null) {
            if (string.IsNullOrEmpty(heading.LinkUri) && string.IsNullOrEmpty(heading.LinkDestinationName)) {
                return;
            }

            double asc = GetAscenderForOptions(font, fontSize, currentOpts);
            double desc = GetDescenderForOptions(font, fontSize, currentOpts);
            for (int i = 0; i < lines.Count; i++) {
                string line = lines[i];
                double lineWidth = EstimateSimpleTextWidthForOptions(line, font, fontSize, currentOpts);
                if (lineWidth <= 0.001D) {
                    continue;
                }

                double dx = 0D;
                if (heading.Align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - lineWidth) / 2);
                else if (heading.Align == PdfAlign.Right) dx = Math.Max(0, widthUsed - lineWidth);
                double baselineY = startBaselineY - i * lineHeight;
                double x1 = x + dx;
                double x2 = x1 + Math.Min(widthUsed, lineWidth);
                double y1 = baselineY - desc;
                double y2 = baselineY + asc;
                currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = heading.LinkUri, DestinationName = heading.LinkDestinationName, Contents = heading.LinkContents, StructElementIndex = structElementIndex });
            }
        }

        private void AddHeadingLinkAnnotations(HeadingBlock heading, System.Collections.Generic.IReadOnlyList<System.Collections.Generic.List<RichSeg>> lines, PdfStandardFont font, double fontSize, double lineHeight, double x, double widthUsed, double startBaselineY, int? structElementIndex = null) {
            if (string.IsNullOrEmpty(heading.LinkUri) && string.IsNullOrEmpty(heading.LinkDestinationName)) {
                return;
            }

            double asc = GetAscenderForOptions(font, fontSize, currentOpts);
            double desc = GetDescenderForOptions(font, fontSize, currentOpts);
            for (int i = 0; i < lines.Count; i++) {
                double lineWidth = MeasureRichLineWidth(lines[i], currentOpts);
                if (lineWidth <= 0.001D) {
                    continue;
                }

                double dx = 0D;
                if (heading.Align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - lineWidth) / 2);
                else if (heading.Align == PdfAlign.Right) dx = Math.Max(0, widthUsed - lineWidth);
                double baselineY = startBaselineY - i * lineHeight;
                double x1 = x + dx;
                double x2 = x1 + Math.Min(widthUsed, lineWidth);
                double y1 = baselineY - desc;
                double y2 = baselineY + asc;
                currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = heading.LinkUri, DestinationName = heading.LinkDestinationName, Contents = heading.LinkContents, StructElementIndex = structElementIndex });
            }
        }

        private void AddImageLinkAnnotation(ImageBlock image, PdfImageStyle style, PageImage pageImage, double targetX, double targetBottomY, double targetWidth, double targetHeight) {
            if (string.IsNullOrEmpty(image.LinkUri)) {
                return;
            }

            GetImageAnnotationBounds(style, pageImage, targetX, targetBottomY, targetWidth, targetHeight, out double x1, out double y1, out double x2, out double y2);

            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = image.LinkUri!, Contents = image.LinkContents, LinkedImage = pageImage });
        }

        private void AddNamedDestination(BookmarkBlock bookmark, double topY) {
            AddNamedDestinationName(bookmark.Name, topY);
        }

        private void AddNamedDestinationName(string name, double topY) {
            EnsurePage();
            currentPage!.NamedDestinations.Add(new PageNamedDestination { Name = name, Y = topY });
        }

        private void AddTableCellNamedDestinationName(string? name, double topY) {
            if (string.IsNullOrWhiteSpace(name) || !emittedTableCellNamedDestinations.Add(name!)) {
                return;
            }

            AddNamedDestinationName(name!, topY);
        }

        private double FirstTextBaselineFromTop(PdfStandardFont font, double fontSize, double topY) =>
            topY - GetAscenderForOptions(font, fontSize, currentOpts);

        private void MarkRichFonts(System.Collections.Generic.IEnumerable<TextRun> runs) {
            System.Collections.Generic.IReadOnlyList<TextRun> effectiveRuns = NormalizeFallbackRuns(runs, ChooseNormal(currentOpts.DefaultFont), currentOpts);
            foreach (TextRun run in effectiveRuns) {
                if (run.InlineElement != null) {
                    continue;
                }

                PdfStandardFont runBaseFont = ChooseNormal(run.Font ?? currentOpts.DefaultFont);
                PdfStandardFont runFont = run.Bold && run.Italic
                    ? ChooseBoldItalic(runBaseFont)
                    : run.Bold
                        ? ChooseBold(runBaseFont)
                        : run.Italic
                            ? ChooseItalic(runBaseFont)
                            : runBaseFont;
                if (currentOpts.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace namedFont)) {
                    currentPage!.UsedNamedFonts.Add(namedFont);
                } else {
                    currentPage!.UsedFonts.Add(runFont);
                }
            }

            if (effectiveRuns.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
            if (effectiveRuns.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
            if (effectiveRuns.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
        }

        private void MarkSimpleFont(PdfStandardFont font) {
            EnsurePage();
            currentPage!.UsedFonts.Add(font);
            PdfStandardFont normalFont = ChooseNormal(currentOpts.DefaultFont);
            if (font == ChooseBold(normalFont)) {
                currentPage.UsedBold = true;
                usedBold = true;
            } else if (font == ChooseItalic(normalFont)) {
                currentPage.UsedItalic = true;
                usedItalic = true;
            } else if (font == ChooseBoldItalic(normalFont)) {
                currentPage.UsedBoldItalic = true;
                usedBoldItalic = true;
            }
        }

    }
}
