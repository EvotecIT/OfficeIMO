namespace OfficeIMO.Pdf;

internal sealed partial class PdfUnderstandingReadingFrame {
    internal IReadOnlyList<PdfUnderstandingWord> ProjectWords(IReadOnlyList<PdfUnderstandingWord> words) {
        var projected = new PdfUnderstandingWord[words.Count];
        for (int index = 0; index < words.Count; index++) {
            _context.ConsumeWork();
            PdfUnderstandingWord word = words[index];
            double radians = word.RotationDegrees * Math.PI / 180D;
            double anchorX = Math.Cos(radians) >= 0D ? word.XStart : word.XEnd;
            (double x, double y) = ToFrame(anchorX, word.BaselineY);
            double advance = word.Advance ?? Math.Max(0D, word.XEnd - word.XStart);
            (double endX, _) = ToFrame(anchorX + Math.Cos(radians) * advance,
                word.BaselineY + Math.Sin(radians) * advance);
            double angle = PdfAdvancedUnderstandingStages.NormalizeAngle(word.RotationDegrees - Angle);
            PdfLogicalVisualBounds? visual = ProjectVisualBounds(word.VisualBounds);
            projected[index] = new PdfUnderstandingWord(word.Text, Math.Min(x, endX), Math.Max(x, endX),
                y, word.FontSize, angle, word.SourceRuns.Select(run => _projectedRuns[run]).ToArray(),
                word.Confidence, word.Evidence, advance, visual, sourceSequence: word.SourceSequence) {
                IsSelectionBox = word.IsSelectionBox
            };
            _originalWords.Add(projected[index], word);
            _projectedWords.Add(word, projected[index]);
        }
        return Array.AsReadOnly(projected);
    }

    internal IReadOnlyList<PdfUnderstandingLine> ProjectLines(IReadOnlyList<PdfUnderstandingLine> lines) {
        var projected = new PdfUnderstandingLine[lines.Count];
        for (int index = 0; index < lines.Count; index++) {
            _context.ConsumeWork();
            PdfUnderstandingLine line = lines[index];
            projected[index] = new PdfUnderstandingLine(line.Words.Select(word => _projectedWords[word]).ToArray(),
                line.Text, line.Confidence, line.Evidence, line.SourceKind, line.SourceSequence,
                line.BlockId, line.ParagraphId, line.LineId, ProjectVisualBounds(line.VisualBounds));
        }
        return Array.AsReadOnly(projected);
    }

    private PdfLogicalVisualBounds? ProjectVisualBounds(PdfLogicalVisualBounds? source) {
        if (source is null) return null;
        PdfVisualBounds bounds = ImageBounds(new PdfVisualBounds(source.Left, source.Top, source.Right, source.Bottom));
        return new PdfLogicalVisualBounds(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
    }
}
