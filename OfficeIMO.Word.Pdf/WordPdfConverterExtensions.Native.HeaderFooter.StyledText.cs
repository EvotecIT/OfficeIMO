using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private enum NativeHeaderFooterZone {
            Left,
            Center,
            Right
        }

        private sealed class NativeHeaderFooterStyledReplacement {
            public NativeHeaderFooterStyledReplacement(string serializedText, PdfCore.PdfTextRun styledRun, bool isFieldToken = false) {
                SerializedText = serializedText;
                StyledRun = styledRun;
                IsFieldToken = isFieldToken;
            }

            public string SerializedText { get; }
            public PdfCore.PdfTextRun StyledRun { get; }
            public bool IsFieldToken { get; }
        }

        private sealed class NativeHeaderFooterText {
            public string? Left { get; private set; }
            public string? Center { get; private set; }
            public string? Right { get; private set; }
            public bool HasPageTokens { get; private set; }
            public PdfCore.PdfPageNumberStyle? PageNumberStyle { get; private set; }
            public List<PdfCore.PdfFooterSegment> LeftSegments { get; } = new List<PdfCore.PdfFooterSegment>();
            public List<PdfCore.PdfFooterSegment> CenterSegments { get; } = new List<PdfCore.PdfFooterSegment>();
            public List<PdfCore.PdfFooterSegment> RightSegments { get; } = new List<PdfCore.PdfFooterSegment>();
            public bool HasStyledZones { get; private set; }
            private bool _hasConflictingPageNumberStyles;
            private int _leftParagraphCount;
            private int _centerParagraphCount;
            private int _rightParagraphCount;
            private bool _leftPreviousHasText;
            private bool _centerPreviousHasText;
            private bool _rightPreviousHasText;
            public bool HasContent =>
                !string.IsNullOrWhiteSpace(Left) ||
                !string.IsNullOrWhiteSpace(Center) ||
                !string.IsNullOrWhiteSpace(Right);

            public Action<PdfCore.PdfHeaderTextBuilder>? CreateHeaderZone(NativeHeaderFooterZone zone) {
                IReadOnlyList<PdfCore.PdfFooterSegment> segments = GetSegments(zone);
                return segments.Count == 0 ? null : builder => AppendSegments(builder, segments);
            }

            public Action<PdfCore.PdfFooterTextBuilder>? CreateFooterZone(NativeHeaderFooterZone zone) {
                IReadOnlyList<PdfCore.PdfFooterSegment> segments = GetSegments(zone);
                return segments.Count == 0 ? null : builder => AppendSegments(builder, segments);
            }

            public void AppendLeft(string text) => Left = Append(Left, text, null, LeftSegments, null, null, null, ref _leftParagraphCount, ref _leftPreviousHasText);
            public void AppendCenter(string text) => Center = Append(Center, text, null, CenterSegments, null, null, null, ref _centerParagraphCount, ref _centerPreviousHasText);
            public void AppendRight(string text) => Right = Append(Right, text, null, RightSegments, null, null, null, ref _rightParagraphCount, ref _rightPreviousHasText);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Left = Append(Left, text, pageNumberStyle, LeftSegments, null, null, null, ref _leftParagraphCount, ref _leftPreviousHasText);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Center = Append(Center, text, pageNumberStyle, CenterSegments, null, null, null, ref _centerParagraphCount, ref _centerPreviousHasText);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Right = Append(Right, text, pageNumberStyle, RightSegments, null, null, null, ref _rightParagraphCount, ref _rightPreviousHasText);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) => AppendLeft(text, pageNumberStyle, markerRun, null);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) => AppendCenter(text, pageNumberStyle, markerRun, null);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) => AppendRight(text, pageNumberStyle, markerRun, null);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) => AppendLeft(text, pageNumberStyle, markerRun, contentStyleRun, null);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) => AppendCenter(text, pageNumberStyle, markerRun, contentStyleRun, null);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) => AppendRight(text, pageNumberStyle, markerRun, contentStyleRun, null);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) => Left = Append(Left, text, pageNumberStyle, LeftSegments, markerRun, contentStyleRun, replacements, ref _leftParagraphCount, ref _leftPreviousHasText);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) => Center = Append(Center, text, pageNumberStyle, CenterSegments, markerRun, contentStyleRun, replacements, ref _centerParagraphCount, ref _centerPreviousHasText);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) => Right = Append(Right, text, pageNumberStyle, RightSegments, markerRun, contentStyleRun, replacements, ref _rightParagraphCount, ref _rightPreviousHasText);

            public void Append(NativeHeaderFooterZone zone, string text) => Append(zone, text, null);

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) {
                switch (zone) {
                    case NativeHeaderFooterZone.Center:
                        AppendCenter(text, pageNumberStyle);
                        break;
                    case NativeHeaderFooterZone.Right:
                        AppendRight(text, pageNumberStyle);
                        break;
                    default:
                        AppendLeft(text, pageNumberStyle);
                        break;
                }
            }

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) {
                Append(zone, text, pageNumberStyle, markerRun, null);
            }

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) {
                Append(zone, text, pageNumberStyle, markerRun, contentStyleRun, null);
            }

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) {
                switch (zone) {
                    case NativeHeaderFooterZone.Center:
                        AppendCenter(text, pageNumberStyle, markerRun, contentStyleRun, replacements);
                        break;
                    case NativeHeaderFooterZone.Right:
                        AppendRight(text, pageNumberStyle, markerRun, contentStyleRun, replacements);
                        break;
                    default:
                        AppendLeft(text, pageNumberStyle, markerRun, contentStyleRun, replacements);
                        break;
                }
            }

            public NativeHeaderFooterText Clone() {
                NativeHeaderFooterText clone = new NativeHeaderFooterText {
                    Left = Left,
                    Center = Center,
                    Right = Right,
                    HasPageTokens = HasPageTokens,
                    PageNumberStyle = PageNumberStyle,
                    _hasConflictingPageNumberStyles = _hasConflictingPageNumberStyles,
                    HasStyledZones = HasStyledZones,
                    _leftParagraphCount = _leftParagraphCount,
                    _centerParagraphCount = _centerParagraphCount,
                    _rightParagraphCount = _rightParagraphCount,
                    _leftPreviousHasText = _leftPreviousHasText,
                    _centerPreviousHasText = _centerPreviousHasText,
                    _rightPreviousHasText = _rightPreviousHasText
                };
                clone.LeftSegments.AddRange(LeftSegments);
                clone.CenterSegments.AddRange(CenterSegments);
                clone.RightSegments.AddRange(RightSegments);
                return clone;
            }

            private string Append(string? current, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, List<PdfCore.PdfFooterSegment> segments, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements, ref int paragraphCount, ref bool previousHasText) {
                text = NormalizeNativeHeaderFooterText(text);
                string plainText = (markerRun?.Text ?? string.Empty) + text;
                bool currentHasText = !string.IsNullOrWhiteSpace(plainText);
                if (text.IndexOf("{page}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("{pages}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("{sectionpages}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("{documentpages}", StringComparison.OrdinalIgnoreCase) >= 0) {
                    HasPageTokens = true;
                }

                RecordPageNumberStyle(pageNumberStyle);
                string separator = previousHasText && currentHasText
                    ? Environment.NewLine + Environment.NewLine
                    : Environment.NewLine;
                if (paragraphCount > 0) {
                    segments.Add(new PdfCore.PdfFooterSegment(PdfCore.PdfFooterSegmentKind.Text, separator));
                }
                if (markerRun != null) {
                    segments.Add(PdfCore.PdfFooterSegment.RichText(markerRun));
                    HasStyledZones = true;
                }
                AppendNativeHeaderFooterSegments(segments, text, contentStyleRun, replacements);
                if (contentStyleRun != null || replacements?.Count > 0) {
                    HasStyledZones = true;
                }
                string combined = paragraphCount == 0 ? plainText : (current ?? string.Empty) + separator + plainText;
                paragraphCount++;
                previousHasText = currentHasText;
                return combined;
            }

            private static void AppendNativeHeaderFooterSegments(List<PdfCore.PdfFooterSegment> segments, string text, PdfCore.PdfTextRun? styleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) {
                int index = 0;
                if (replacements != null) {
                    foreach (NativeHeaderFooterStyledReplacement replacement in replacements) {
                        // Match the same newline/tab representation as the zone text.
                        string serializedText = NormalizeNativeHeaderFooterText(replacement.SerializedText);
                        if (serializedText.Length == 0) continue;
                        int replacementIndex = text.IndexOf(serializedText, index, StringComparison.Ordinal);
                        if (replacementIndex < 0) {
                            continue;
                        }

                        AppendTokenizedText(text.Substring(index, replacementIndex - index));
                        bool beginsVisualLine = replacementIndex == 0 || text[replacementIndex - 1] == '\r' || text[replacementIndex - 1] == '\n';
                        PdfCore.PdfTextRun styledRun = beginsVisualLine
                            ? replacement.StyledRun
                            : replacement.StyledRun.WithHorizontalOffset(0D);
                        if (replacement.IsFieldToken && replacement.SerializedText == "{page}")
                            segments.Add(PdfCore.PdfFooterSegment.PageNumber(CloneNativeHeaderFooterTextRun(styledRun, string.Empty)));
                        else if (replacement.IsFieldToken && replacement.SerializedText == "{pages}")
                            segments.Add(PdfCore.PdfFooterSegment.TotalPages(CloneNativeHeaderFooterTextRun(styledRun, string.Empty)));
                        else if (replacement.IsFieldToken && replacement.SerializedText == "{documentpages}")
                            segments.Add(PdfCore.PdfFooterSegment.DocumentPages(CloneNativeHeaderFooterTextRun(styledRun, string.Empty)));
                        else segments.Add(PdfCore.PdfFooterSegment.RichText(styledRun));
                        index = replacementIndex + serializedText.Length;
                    }
                }
                AppendTokenizedText(text.Substring(index));

                void AppendTokenizedText(string value) {
                    int tokenCursor = 0;
                    while (tokenCursor < value.Length) {
                        int pageIndex = value.IndexOf("{page}", tokenCursor, StringComparison.OrdinalIgnoreCase);
                        int pagesIndex = value.IndexOf("{pages}", tokenCursor, StringComparison.OrdinalIgnoreCase);
                        int documentPagesIndex = value.IndexOf("{documentpages}", tokenCursor, StringComparison.OrdinalIgnoreCase);
                        int tokenIndex = pageIndex < 0 ? pagesIndex : pagesIndex < 0 ? pageIndex : Math.Min(pageIndex, pagesIndex);
                        if (documentPagesIndex >= 0) tokenIndex = tokenIndex < 0 ? documentPagesIndex : Math.Min(tokenIndex, documentPagesIndex);
                        if (tokenIndex < 0) {
                            if (tokenCursor < value.Length) AddText(value.Substring(tokenCursor));
                            break;
                        }
                        if (tokenIndex > tokenCursor) AddText(value.Substring(tokenCursor, tokenIndex - tokenCursor));
                        bool documentTotal = tokenIndex == documentPagesIndex;
                        bool totalPages = tokenIndex == pagesIndex;
                        if (styleRun != null) {
                            PdfCore.PdfTextRun tokenStyle = CloneNativeHeaderFooterTextRun(styleRun, string.Empty);
                            segments.Add(documentTotal ? PdfCore.PdfFooterSegment.DocumentPages(tokenStyle) : totalPages ? PdfCore.PdfFooterSegment.TotalPages(tokenStyle) : PdfCore.PdfFooterSegment.PageNumber(tokenStyle));
                        } else {
                            segments.Add(new PdfCore.PdfFooterSegment(documentTotal ? PdfCore.PdfFooterSegmentKind.DocumentPages : totalPages ? PdfCore.PdfFooterSegmentKind.TotalPages : PdfCore.PdfFooterSegmentKind.PageNumber));
                        }
                        tokenCursor = tokenIndex + (documentTotal ? 15 : totalPages ? 7 : 6);
                    }
                }

                void AddText(string value) {
                    if (styleRun != null) {
                        segments.Add(PdfCore.PdfFooterSegment.RichText(CloneNativeHeaderFooterTextRun(styleRun, value)));
                    } else {
                        segments.Add(new PdfCore.PdfFooterSegment(PdfCore.PdfFooterSegmentKind.Text, value));
                    }
                }
            }

            private IReadOnlyList<PdfCore.PdfFooterSegment> GetSegments(NativeHeaderFooterZone zone) => zone switch {
                NativeHeaderFooterZone.Center => CenterSegments,
                NativeHeaderFooterZone.Right => RightSegments,
                _ => LeftSegments
            };

            private static void AppendSegments(PdfCore.PdfHeaderTextBuilder builder, IReadOnlyList<PdfCore.PdfFooterSegment> segments) {
                foreach (PdfCore.PdfFooterSegment segment in segments) {
                    if (segment.Kind == PdfCore.PdfFooterSegmentKind.PageNumber) {
                        if (segment.StyledRun != null) builder.CurrentPage(segment.StyledRun); else builder.CurrentPage();
                    } else if (segment.Kind == PdfCore.PdfFooterSegmentKind.TotalPages) {
                        if (segment.StyledRun != null) builder.TotalPages(segment.StyledRun); else builder.TotalPages();
                    } else if (segment.Kind == PdfCore.PdfFooterSegmentKind.DocumentPages) {
                        if (segment.StyledRun != null) builder.DocumentPages(segment.StyledRun); else builder.DocumentPages();
                    } else if (segment.StyledRun != null) builder.Run(segment.StyledRun); else builder.Text(segment.Text ?? string.Empty);
                }
            }

            private static void AppendSegments(PdfCore.PdfFooterTextBuilder builder, IReadOnlyList<PdfCore.PdfFooterSegment> segments) {
                foreach (PdfCore.PdfFooterSegment segment in segments) {
                    if (segment.Kind == PdfCore.PdfFooterSegmentKind.PageNumber) {
                        if (segment.StyledRun != null) builder.CurrentPage(segment.StyledRun); else builder.CurrentPage();
                    } else if (segment.Kind == PdfCore.PdfFooterSegmentKind.TotalPages) {
                        if (segment.StyledRun != null) builder.TotalPages(segment.StyledRun); else builder.TotalPages();
                    } else if (segment.Kind == PdfCore.PdfFooterSegmentKind.DocumentPages) {
                        if (segment.StyledRun != null) builder.DocumentPages(segment.StyledRun); else builder.DocumentPages();
                    } else if (segment.StyledRun != null) builder.Run(segment.StyledRun); else builder.Text(segment.Text ?? string.Empty);
                }
            }

            private static string NormalizeNativeHeaderFooterText(string? text) {
                if (string.IsNullOrEmpty(text)) {
                    return string.Empty;
                }

                string normalized = text!
                    .Replace("\r\n", "\n")
                    .Replace('\r', '\n');
                string[] lines = normalized.Split('\n');
                for (int i = 0; i < lines.Length; i++) {
                    lines[i] = NormalizeNativeDirectText(lines[i]);
                }

                return string.Join(Environment.NewLine, lines);
            }

            private void RecordPageNumberStyle(PdfCore.PdfPageNumberStyle? style) {
                if (!style.HasValue || _hasConflictingPageNumberStyles) {
                    return;
                }

                if (PageNumberStyle.HasValue && PageNumberStyle.Value != style.Value) {
                    PageNumberStyle = null;
                    _hasConflictingPageNumberStyles = true;
                    return;
                }

                PageNumberStyle = style.Value;
            }
        }

    }
}
