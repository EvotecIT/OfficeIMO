using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static bool RemoveIntersectingPathObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, PdfRedactionArea[] areas, PdfReadLimits limits, HashSet<PdfStream> sourceStreamIdentities, ref int nextObjectNumber) {
        if (areas.Length == 0 || !page.Items.TryGetValue("Contents", out PdfObject? contentsObject)) return false;
        Dictionary<string, double> graphicsStateLineWidths = GetPathGraphicsStateLineWidths(objects, page);
        Dictionary<int, int> referenceCounts = CountIndirectReferenceUsage(objects); PdfObject currentContents = contentsObject; bool changed = false;
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject).ToArray()) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) || indirect.Value is not PdfStream stream || stream.DecodingFailed) continue;
            string content = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, GetMutationDecodeLimit(stream, limits, sourceStreamIdentities))); string scrubbed = ScrubIntersectingPaths(content, areas, graphicsStateLineWidths);
            if (string.Equals(content, scrubbed, StringComparison.Ordinal)) continue;
            PdfReference target = reference;
            if (IsSharedReference(referenceCounts, reference)) { target = CloneIndirectObject(objects, reference, indirect, ref nextObjectNumber); ReplacePageContentReference(objects, page, currentContents, reference, target); currentContents = page.Items.TryGetValue("Contents", out PdfObject? updated) ? updated : currentContents; }
            objects[target.ObjectNumber] = new PdfIndirectObject(target.ObjectNumber, target.Generation, new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed))); changed = true;
        }
        return changed;
    }

    private static string ScrubIntersectingPaths(string content, PdfRedactionArea[] areas, Dictionary<string, double> graphicsStateLineWidths) {
        var ranges = new List<RemovalRange>(); var args = new List<ImageContentOperand>(8); var stack = new Stack<(Matrix2D Transform, double LineWidth)>(); Matrix2D ctm = Matrix2D.Identity; double lineWidth = 1D;
        int pathStart = -1; double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue; int index = 0;
        while (index < content.Length) {
            SkipWhiteSpace(content, ref index); if (index >= content.Length) break; char current = content[index];
            if (current == '%') { SkipComment(content, ref index); continue; }
            if (current == '/') { args.Add(ReadNameOperand(content, ref index)); continue; }
            if (current == '(') { SkipLiteralString(content, ref index); continue; }
            if (current == '<') { if (index + 1 < content.Length && content[index + 1] == '<') SkipDictionary(content, ref index); else SkipHexString(content, ref index); continue; }
            if (current == '[') { SkipArray(content, ref index); continue; }
            if (IsNumberStart(current)) { args.Add(ReadNumberOperand(content, ref index)); continue; }
            string op = ReadOperator(content, ref index); int opEnd = index; if (op.Length == 0) { index++; continue; }
            if (op == "q") stack.Push((ctm, lineWidth)); else if (op == "Q") { (ctm, lineWidth) = stack.Count > 0 ? stack.Pop() : (Matrix2D.Identity, 1D); } else if (op == "cm" && args.Count >= 6) { int start = args.Count - 6; ctm = Matrix2D.Multiply(ctm, new Matrix2D(args[start].Number, args[start + 1].Number, args[start + 2].Number, args[start + 3].Number, args[start + 4].Number, args[start + 5].Number)); }
            else if (op == "w" && args.Count > 0) lineWidth = Math.Max(0D, args[args.Count - 1].Number);
            else if (op == "gs" && args.Count > 0 && args[args.Count - 1].Name is string graphicsStateName && graphicsStateLineWidths.TryGetValue(graphicsStateName, out double graphicsStateLineWidth)) lineWidth = graphicsStateLineWidth;
            else if (op == "m" || op == "l") { StartPath(args, ref pathStart); if (args.Count >= 2) AddPoint(ctm, args[args.Count - 2].Number, args[args.Count - 1].Number, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "c") { StartPath(args, ref pathStart); for (int i = Math.Max(0, args.Count - 6); i + 1 < args.Count; i += 2) AddPoint(ctm, args[i].Number, args[i + 1].Number, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "v" || op == "y") { StartPath(args, ref pathStart); for (int i = Math.Max(0, args.Count - 4); i + 1 < args.Count; i += 2) AddPoint(ctm, args[i].Number, args[i + 1].Number, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "re" && args.Count >= 4) { StartPath(args, ref pathStart); int start = args.Count - 4; double x = args[start].Number, y = args[start + 1].Number, width = args[start + 2].Number, height = args[start + 3].Number; AddPoint(ctm, x, y, ref minX, ref minY, ref maxX, ref maxY); AddPoint(ctm, x + width, y, ref minX, ref minY, ref maxX, ref maxY); AddPoint(ctm, x, y + height, ref minX, ref minY, ref maxX, ref maxY); AddPoint(ctm, x + width, y + height, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "n") { pathStart = -1; minX = minY = double.MaxValue; maxX = maxY = double.MinValue; }
            else if (IsPathPaintOperator(op)) {
                double strokePadding = IsPathStrokePaintOperator(op) ? GetRenderedStrokeWidth(lineWidth, ctm) / 2D : 0D;
                double pathX = minX - strokePadding;
                double pathY = minY - strokePadding;
                double pathWidth = maxX - minX + strokePadding * 2D;
                double pathHeight = maxY - minY + strokePadding * 2D;
                if (pathStart >= 0 && maxX >= minX && maxY >= minY) {
                    PdfRedactionArea[] intersections = areas.Where(area => area.IntersectsRectangle(pathX, pathY, pathWidth, pathHeight)).ToArray();
                    if (intersections.Any(area => area.ExactGeometry is not null && !area.ContainsRectangle(pathX, pathY, pathWidth, pathHeight))) {
                        throw new NotSupportedException("An exact non-rectangular redaction intersects only part of a vector path. The engine refuses to remove content outside the reviewed geometry.");
                    }
                    if (intersections.Length > 0) ranges.Add(new RemovalRange(pathStart, opEnd));
                }
                pathStart = -1; minX = minY = double.MaxValue; maxX = maxY = double.MinValue;
            }
            args.Clear();
        }
        return RemoveRanges(content, ranges);
    }

    private static void StartPath(List<ImageContentOperand> args, ref int pathStart) { if (pathStart < 0 && args.Count > 0) pathStart = args[0].Start; }
    private static void AddPoint(Matrix2D transform, double x, double y, ref double minX, ref double minY, ref double maxX, ref double maxY) { var point = transform.Transform(x, y); minX = Math.Min(minX, point.X); minY = Math.Min(minY, point.Y); maxX = Math.Max(maxX, point.X); maxY = Math.Max(maxY, point.Y); }
    private static bool IsPathPaintOperator(string value) => value == "S" || value == "s" || value == "f" || value == "F" || value == "f*" || value == "B" || value == "B*" || value == "b" || value == "b*";

    private static bool IsPathStrokePaintOperator(string value) => value == "S" || value == "s" || value == "B" || value == "B*" || value == "b" || value == "b*";

    private static Dictionary<string, double> GetPathGraphicsStateLineWidths(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        var lineWidths = new Dictionary<string, double>(StringComparer.Ordinal);
        PdfDictionary? resources = GetInheritedDictionary(objects, page, "Resources");
        if (resources is null ||
            ResolveDictionary(objects, resources.Items.TryGetValue("ExtGState", out PdfObject? value) ? value : null) is not PdfDictionary graphicsStates) {
            return lineWidths;
        }
        foreach (KeyValuePair<string, PdfObject> entry in graphicsStates.Items) {
            if (ResolveDictionary(objects, entry.Value) is PdfDictionary graphicsState &&
                PdfObjectLookup.Resolve(objects, graphicsState.Items.TryGetValue("LW", out PdfObject? width) ? width : null) is PdfNumber number &&
                !double.IsNaN(number.Value) &&
                !double.IsInfinity(number.Value) &&
                number.Value >= 0D) {
                lineWidths[entry.Key] = number.Value;
            }
        }
        return lineWidths;
    }

    private static double GetRenderedStrokeWidth(double lineWidth, Matrix2D transform) =>
        lineWidth == 0D ? 0.25D : lineWidth * GetMaximumScale(transform);

    private static double GetMaximumScale(Matrix2D transform) {
        double first = transform.A * transform.A + transform.B * transform.B;
        double second = transform.C * transform.C + transform.D * transform.D;
        double cross = transform.A * transform.C + transform.B * transform.D;
        double discriminant = Math.Sqrt(Math.Max(0D, (first - second) * (first - second) + 4D * cross * cross));
        return Math.Sqrt(Math.Max(0D, (first + second + discriminant) / 2D));
    }
}
