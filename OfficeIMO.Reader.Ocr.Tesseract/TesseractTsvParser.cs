using OfficeIMO.Reader;

namespace OfficeIMO.Reader.Ocr.Tesseract;

internal static class TesseractTsvParser {
    internal static OfficeOcrEngineResult Parse(string tsv, string? language) {
        if (tsv == null) throw new ArgumentNullException(nameof(tsv));
        string[] lines = tsv.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        int headerIndex = Array.FindIndex(lines, static line => !string.IsNullOrWhiteSpace(line));
        if (headerIndex < 0 || !lines[headerIndex].TrimStart('\uFEFF').StartsWith("level\tpage_num\tblock_num", StringComparison.Ordinal)) {
            throw new InvalidDataException("Tesseract TSV output is missing the expected header.");
        }

        var words = new List<WordRow>();
        int invalidRowCount = 0;
        for (int lineIndex = headerIndex + 1; lineIndex < lines.Length; lineIndex++) {
            if (string.IsNullOrWhiteSpace(lines[lineIndex])) continue;
            string[] columns = lines[lineIndex].Split(new[] { '\t' }, 12, StringSplitOptions.None);
            if (columns.Length < 11 || !TryParseWord(columns, lineIndex, out WordRow? word)) {
                invalidRowCount++;
                continue;
            }
            if (word != null) words.Add(word);
        }

        var spans = new List<OfficeOcrTextSpan>();
        var recognizedLines = new List<string>();
        int sequence = 0;
        IEnumerable<IGrouping<(int Page, int Block, int Paragraph, int Line), WordRow>> groups = words
            .GroupBy(static word => (word.Page, word.Block, word.Paragraph, word.Line))
            .OrderBy(static group => group.Min(word => word.SourceIndex));
        foreach (IGrouping<(int Page, int Block, int Paragraph, int Line), WordRow> group in groups) {
            WordRow[] lineWords = group.OrderBy(static word => word.SourceIndex).ToArray();
            string lineText = string.Join(" ", lineWords.Select(static word => word.Text));
            recognizedLines.Add(lineText);
            spans.Add(new OfficeOcrTextSpan {
                Sequence = sequence++,
                Level = OfficeOcrTextSpanLevel.Line,
                Text = lineText,
                Confidence = AverageConfidence(lineWords),
                Language = language,
                PageNumber = group.Key.Page,
                Region = Union(lineWords),
                CoordinateUnit = OfficeOcrCoordinateUnit.Pixels
            });
            foreach (WordRow word in lineWords) {
                spans.Add(new OfficeOcrTextSpan {
                    Sequence = sequence++,
                    Level = OfficeOcrTextSpanLevel.Word,
                    Text = word.Text,
                    Confidence = word.Confidence,
                    Language = language,
                    PageNumber = word.Page,
                    Region = new OfficeDocumentRegion { X = word.Left, Y = word.Top, Width = word.Width, Height = word.Height },
                    CoordinateUnit = OfficeOcrCoordinateUnit.Pixels
                });
            }
        }

        IReadOnlyList<OfficeDocumentDiagnostic> diagnostics = invalidRowCount == 0
            ? Array.Empty<OfficeDocumentDiagnostic>()
            : new[] {
                new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Category = OfficeDocumentDiagnosticCategory.Ocr,
                    Code = "tesseract-tsv-row-invalid",
                    Message = "One or more malformed Tesseract TSV rows were ignored.",
                    Source = "tesseract-cli",
                    IsRecoverable = true,
                    Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                        ["invalidRowCount"] = invalidRowCount.ToString(CultureInfo.InvariantCulture)
                    }
                }
            };
        return new OfficeOcrEngineResult {
            Text = string.Join(Environment.NewLine, recognizedLines),
            Confidence = AverageConfidence(words),
            Language = language,
            Provider = "tesseract-cli",
            Model = string.IsNullOrWhiteSpace(language) ? null : "tessdata:" + language,
            Spans = spans,
            Diagnostics = diagnostics
        };
    }

    private static bool TryParseWord(string[] columns, int sourceIndex, out WordRow? word) {
        word = null;
        if (!int.TryParse(columns[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int level)) return false;
        if (level != 5) return true;
        if (!int.TryParse(columns[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int page)
            || !int.TryParse(columns[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int block)
            || !int.TryParse(columns[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out int paragraph)
            || !int.TryParse(columns[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out int line)
            || !int.TryParse(columns[6], NumberStyles.Integer, CultureInfo.InvariantCulture, out int left)
            || !int.TryParse(columns[7], NumberStyles.Integer, CultureInfo.InvariantCulture, out int top)
            || !int.TryParse(columns[8], NumberStyles.Integer, CultureInfo.InvariantCulture, out int width)
            || !int.TryParse(columns[9], NumberStyles.Integer, CultureInfo.InvariantCulture, out int height)
            || !double.TryParse(columns[10], NumberStyles.Float, CultureInfo.InvariantCulture, out double confidence)) return false;
        string text = columns.Length > 11 ? columns[11].Trim() : string.Empty;
        if (text.Length == 0) return true;
        word = new WordRow(sourceIndex, page, block, paragraph, line, left, top, width, height, confidence < 0D ? null : confidence / 100D, text);
        return true;
    }

    private static double? AverageConfidence(IEnumerable<WordRow> words) {
        double[] values = words.Where(static word => word.Confidence.HasValue).Select(static word => word.Confidence!.Value).ToArray();
        return values.Length == 0 ? null : values.Average();
    }

    private static OfficeDocumentRegion Union(IReadOnlyList<WordRow> words) {
        int left = words.Min(static word => word.Left);
        int top = words.Min(static word => word.Top);
        int right = words.Max(static word => word.Left + word.Width);
        int bottom = words.Max(static word => word.Top + word.Height);
        return new OfficeDocumentRegion { X = left, Y = top, Width = right - left, Height = bottom - top };
    }

    private sealed class WordRow {
        internal WordRow(int sourceIndex, int page, int block, int paragraph, int line, int left, int top, int width, int height, double? confidence, string text) {
            SourceIndex = sourceIndex;
            Page = page;
            Block = block;
            Paragraph = paragraph;
            Line = line;
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            Confidence = confidence;
            Text = text;
        }

        internal int SourceIndex { get; }
        internal int Page { get; }
        internal int Block { get; }
        internal int Paragraph { get; }
        internal int Line { get; }
        internal int Left { get; }
        internal int Top { get; }
        internal int Width { get; }
        internal int Height { get; }
        internal double? Confidence { get; }
        internal string Text { get; }
    }
}
