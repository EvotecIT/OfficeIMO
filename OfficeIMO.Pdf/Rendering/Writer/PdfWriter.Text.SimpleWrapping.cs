using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static System.Collections.Generic.List<string> WrapMonospace(string text, double widthPts, double fontSize, double glyphWidthEm) {
        double glyphWidth = fontSize * glyphWidthEm;
        int maxChars = Math.Max(8, (int)Math.Floor(widthPts / glyphWidth));
        int maxUnspacedTokenChars = Math.Max(1, (int)Math.Floor(widthPts / (fontSize * Math.Max(glyphWidthEm, 0.9))));
        var hardLines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split(HardLineSplitChars, StringSplitOptions.None);
        var lines = new System.Collections.Generic.List<string>();
        var line = new StringBuilder();
        void AddWrappedWord(string word) {
            if (word.Length <= maxChars) {
                line.Append(word);
                return;
            }

            for (int i = 0; i < word.Length; i += maxUnspacedTokenChars) {
                var chunk = word.Substring(i, Math.Min(maxUnspacedTokenChars, word.Length - i));
                if (chunk.Length == maxUnspacedTokenChars) {
                    lines.Add(chunk);
                } else {
                    line.Append(chunk);
                }
            }
        }

        void AddSoftWrappedLine(string hardLine) {
            int startingLineCount = lines.Count;
            var words = hardLine.Split(SoftLineSplitChars, StringSplitOptions.None);
            foreach (var w in words) {
                if (line.Length == 0) {
                    AddWrappedWord(w);
                } else {
                    if (line.Length + 1 + w.Length <= maxChars) {
                        line.Append(' ').Append(w);
                    } else {
                        lines.Add(line.ToString());
                        line.Clear();
                        AddWrappedWord(w);
                    }
                }
            }

            if (line.Length > 0) {
                lines.Add(line.ToString());
                line.Clear();
            } else if (hardLine.Length == 0 && lines.Count == startingLineCount) {
                lines.Add(string.Empty);
            }
        }

        for (int i = 0; i < hardLines.Length; i++) {
            AddSoftWrappedLine(hardLines[i]);
        }

        if (lines.Count == 0) lines.Add(string.Empty);
        return lines;
    }

    private static System.Collections.Generic.List<string> WrapSimpleText(string text, double widthPts, PdfStandardFont font, double fontSize) =>
        WrapSimpleTextForOptions(text, widthPts, font, fontSize, options: null);

    private static System.Collections.Generic.List<string> WrapSimpleTextForOptions(string text, double widthPts, PdfStandardFont font, double fontSize, PdfOptions? options) {
        var hardLines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split(HardLineSplitChars, StringSplitOptions.None);
        var lines = new System.Collections.Generic.List<string>();
        double maxWidth = Math.Max(1D, widthPts);
        double spaceWidth = EstimateSimpleTextWidthForOptions(" ", font, fontSize, options);

        void FlushLine(StringBuilder current, ref double currentWidth) {
            if (current.Length > 0) {
                lines.Add(current.ToString());
                current.Clear();
                currentWidth = 0D;
            }
        }

        void AppendLongToken(string token, StringBuilder current, ref double currentWidth) {
            FlushLine(current, ref currentWidth);
            var softLineBreakChunks = TryBuildSoftLineBreakTokenChunks(
                token,
                options,
                part => EstimateSimpleTextWidthForOptions(part, font, fontSize, options),
                maxWidth,
                maxWidth);
            if (softLineBreakChunks != null) {
                AppendTokenChunks(softLineBreakChunks, current, ref currentWidth);
                return;
            }

            var multilingualChunks = TryBuildMultilingualTokenChunks(
                token,
                part => EstimateSimpleTextWidthForOptions(part, font, fontSize, options),
                maxWidth,
                maxWidth);
            if (multilingualChunks != null) {
                AppendTokenChunks(multilingualChunks, current, ref currentWidth);
                return;
            }

            for (int i = 0; i < token.Length; i++) {
                int scalarLength = GetScalarUtf16Length(token, i);
                string scalar = token.Substring(i, scalarLength);
                double characterWidth = EstimateSimpleTextWidthForOptions(scalar, font, fontSize, options);
                if (current.Length > 0 && currentWidth + characterWidth > maxWidth) {
                    FlushLine(current, ref currentWidth);
                }

                current.Append(scalar);
                currentWidth += characterWidth;
                i += scalarLength - 1;
            }
        }

        void AppendTokenChunks(System.Collections.Generic.IReadOnlyList<PdfTextTokenChunk> chunks, StringBuilder current, ref double currentWidth) {
            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                current.Append(chunk.Text);
                currentWidth += chunk.Width;
                if (chunkIndex + 1 < chunks.Count) {
                    FlushLine(current, ref currentWidth);
                }
            }
        }

        for (int hardLineIndex = 0; hardLineIndex < hardLines.Length; hardLineIndex++) {
            string hardLine = hardLines[hardLineIndex];
            int startingLineCount = lines.Count;
            var current = new StringBuilder();
            double currentWidth = 0D;
            bool pendingSpace = false;
            int index = 0;

            while (index < hardLine.Length) {
                int nextWhitespace = hardLine.IndexOfAny(SoftLineSplitChars, index);
                string token;
                if (nextWhitespace == -1) {
                    token = hardLine.Substring(index);
                    index = hardLine.Length;
                } else {
                    token = hardLine.Substring(index, nextWhitespace - index);
                    index = nextWhitespace + 1;
                }

                if (token.Length > 0) {
                    double tokenWidth = EstimateSimpleTextWidthForOptions(token, font, fontSize, options);
                    if (tokenWidth > maxWidth) {
                        AppendLongToken(token, current, ref currentWidth);
                    } else {
                        double neededWidth = current.Length == 0 ? tokenWidth : (pendingSpace ? spaceWidth : 0D) + tokenWidth;
                        if (current.Length > 0 && currentWidth + neededWidth > maxWidth) {
                            FlushLine(current, ref currentWidth);
                        }

                        if (current.Length > 0 && pendingSpace) {
                            current.Append(' ');
                            currentWidth += spaceWidth;
                        }

                        current.Append(token);
                        currentWidth += tokenWidth;
                    }

                    pendingSpace = false;
                }

                if (nextWhitespace != -1) {
                    pendingSpace = true;
                }
            }

            FlushLine(current, ref currentWidth);
            if (hardLine.Length == 0 && lines.Count == startingLineCount) {
                lines.Add(string.Empty);
            }
        }

        if (lines.Count == 0) lines.Add(string.Empty);
        return lines;
    }
}
