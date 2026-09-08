namespace OfficeIMO.Pdf;

internal sealed partial class PdfUnderstandingReadingFrame {
    internal PdfUnderstandingPageResult Restore(PdfUnderstandingPageResult result) {
        var words = new Dictionary<PdfUnderstandingWord, PdfUnderstandingWord>();
        var lines = new Dictionary<PdfUnderstandingLine, PdfUnderstandingLine>();
        var regions = new Dictionary<PdfUnderstandingRegion, PdfUnderstandingRegion>();
        var elements = new Dictionary<PdfUnderstandingSemanticElement, PdfUnderstandingSemanticElement>();
        return new PdfUnderstandingPageResult(result.PageNumber,
            result.DecodedRuns.Select(RestoreRun).ToArray(), result.Words.Select(RestoreWord).ToArray(),
            result.Lines.Select(RestoreLine).ToArray(), result.Regions.Select(RestoreRegion).ToArray(),
            result.ReadingOrder.Select(RestoreRegion).ToArray(),
            result.ReadingOrderEvidence.Select(item => new PdfReadingOrderEvidence(item.Index, RestoreRegion(item.Region), item.Confidence, item.Evidence)).ToArray(),
            result.Elements.Select(RestoreElement).ToArray(), result.Trace,
            result.ConsumeWork, result.CancellationCheck, result.CompleteOperation,
            result.LogicalProjectionLines.Select(RestoreLine).ToArray(), result.RestrictLogicalProjectionToReadingOrder,
            result.TableCandidates.Select(RestoreTable).ToArray(), result.ImagePlacements,
            result.ImageRegions.Select(item => new PdfUnderstandingImageRegion(item.Placement,
                item.Caption is null ? null : RestoreElement(item.Caption), item.Confidence, item.Evidence,
                item.IsFigure, item.AlternativeText)).ToArray());

        PdfTextSpan RestoreRun(PdfTextSpan run) {
            _context.ConsumeWork();
            return _originalRuns[run];
        }

        PdfUnderstandingWord RestoreWord(PdfUnderstandingWord word) {
            _context.ConsumeWork();
            if (words.TryGetValue(word, out PdfUnderstandingWord? restored)) return restored;
            if (_originalWords.TryGetValue(word, out restored)) {
                words.Add(word, restored);
                return restored;
            }
            double radians = word.RotationDegrees * Math.PI / 180D;
            double anchorX = Math.Cos(radians) >= 0D ? word.XStart : word.XEnd;
            (double x, double y) = ToSource(anchorX, word.BaselineY);
            double advance = word.Advance ?? Math.Max(0D, word.XEnd - word.XStart);
            (double endX, _) = ToSource(anchorX + Math.Cos(radians) * advance, word.BaselineY + Math.Sin(radians) * advance);
            var region = new PdfUnderstandingRegion(new[] { new PdfUnderstandingLine(new[] { word }) });
            var bounds = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(region);
            restored = new PdfUnderstandingWord(word.Text, Math.Min(x, endX), Math.Max(x, endX), y, word.FontSize,
                PdfAdvancedUnderstandingStages.NormalizeAngle(word.RotationDegrees + Angle),
                word.SourceRuns.Select(RestoreRun).ToArray(), word.Confidence, word.Evidence, word.Advance,
                word.VisualBounds is not null ? RestoreVisualBounds(word.VisualBounds)
                    : SourceVisualBounds(bounds.Left, bounds.Bottom, bounds.Right, bounds.Top), word.SourceSequence) {
                IsSelectionBox = word.IsSelectionBox
            };
            words.Add(word, restored);
            return restored;
        }

        PdfUnderstandingLine RestoreLine(PdfUnderstandingLine line) {
            _context.ConsumeWork();
            if (lines.TryGetValue(line, out PdfUnderstandingLine? restored)) return restored;
            var bounds = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(new PdfUnderstandingRegion(new[] { line }));
            restored = new PdfUnderstandingLine(line.Words.Select(RestoreWord).ToArray(), line.Text, line.Confidence,
                line.Evidence.Concat(new[] { new PdfInferenceEvidence("line.corrected-reading-frame",
                    "Layout was inferred along the dominant quarter-turn baseline and mapped back to the source page.", 0.8D) }),
                line.SourceKind, line.SourceSequence, line.BlockId, line.ParagraphId, line.LineId,
                line.VisualBounds is not null ? RestoreVisualBounds(line.VisualBounds)
                    : SourceVisualBounds(bounds.Left, bounds.Bottom, bounds.Right, bounds.Top));
            lines.Add(line, restored);
            return restored;
        }

        PdfUnderstandingRegion RestoreRegion(PdfUnderstandingRegion region) {
            _context.ConsumeWork();
            if (regions.TryGetValue(region, out PdfUnderstandingRegion? restored)) return restored;
            restored = new PdfUnderstandingRegion(region.Lines.Select(RestoreLine).ToArray(), region.Confidence, region.Evidence);
            regions.Add(region, restored);
            return restored;
        }

        PdfUnderstandingSemanticElement RestoreElement(PdfUnderstandingSemanticElement element) {
            _context.ConsumeWork();
            if (elements.TryGetValue(element, out PdfUnderstandingSemanticElement? restored)) return restored;
            restored = new PdfUnderstandingSemanticElement(RestoreRegion(element.Region), element.Kind,
                element.Confidence, element.Evidence, element.Level);
            elements.Add(element, restored);
            return restored;
        }

        PdfUnderstandingTableCandidate RestoreTable(PdfUnderstandingTableCandidate table) {
            _context.ConsumeWork();
            double left = table.Columns.Min(static column => column.From), right = table.Columns.Max(static column => column.To);
            double bottom = table.CoordinateSpace == PdfTableCoordinateSpace.VisualTopLeft ? Height - table.YBottom : table.YBottom;
            double top = table.CoordinateSpace == PdfTableCoordinateSpace.VisualTopLeft ? Height - table.YTop : table.YTop;
            PdfLogicalVisualBounds visual = SourceVisualBounds(left, bottom, right, top);
            var columns = new PdfUnderstandingTableColumn[table.Columns.Count];
            for (int index = 0; index < columns.Length; index++) {
                _context.ConsumeWork();
                PdfLogicalVisualBounds column = SourceVisualBounds(table.Columns[index].From, bottom, table.Columns[index].To, top);
                columns[index] = new PdfUnderstandingTableColumn(column.Left, column.Right);
            }
            return table.WithSourceGeometry(visual, columns, table.SourceLines.Select(RestoreLine).ToArray(),
                table.NativeSourceRuns.Select(RestoreRun).ToArray(), _context.ConsumeWork, _context.ThrowIfCancellationRequested);
        }
    }

    private PdfLogicalVisualBounds RestoreVisualBounds(PdfLogicalVisualBounds bounds) =>
        SourceVisualBounds(bounds.Left, Height - bounds.Bottom, bounds.Right, Height - bounds.Top);
}
