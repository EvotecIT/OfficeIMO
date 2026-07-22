using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static bool RemoveIntersectingPathObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, PdfRedactionArea[] areas, int maximumDecodedStreamBytes, ref int nextObjectNumber) {
        if (areas.Length == 0 || !page.Items.TryGetValue("Contents", out PdfObject? contentsObject)) return false;
        Dictionary<int, int> referenceCounts = CountIndirectReferenceUsage(objects); PdfObject currentContents = contentsObject; bool changed = false;
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject).ToArray()) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) || indirect.Value is not PdfStream stream || stream.DecodingFailed) continue;
            string content = PdfEncoding.Latin1GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects, maximumDecodedStreamBytes)); string scrubbed = ScrubIntersectingPaths(content, areas);
            if (string.Equals(content, scrubbed, StringComparison.Ordinal)) continue;
            PdfReference target = reference;
            if (IsSharedReference(referenceCounts, reference)) { target = CloneIndirectObject(objects, reference, indirect, ref nextObjectNumber); ReplacePageContentReference(objects, page, currentContents, reference, target); currentContents = page.Items.TryGetValue("Contents", out PdfObject? updated) ? updated : currentContents; }
            objects[target.ObjectNumber] = new PdfIndirectObject(target.ObjectNumber, target.Generation, new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed))); changed = true;
        }
        return changed;
    }

    private static string ScrubIntersectingPaths(string content, PdfRedactionArea[] areas) {
        var ranges = new List<RemovalRange>(); var args = new List<ImageContentOperand>(8); var stack = new Stack<Matrix2D>(); Matrix2D ctm = Matrix2D.Identity;
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
            if (op == "q") stack.Push(ctm); else if (op == "Q") ctm = stack.Count > 0 ? stack.Pop() : Matrix2D.Identity; else if (op == "cm" && args.Count >= 6) { int start = args.Count - 6; ctm = Matrix2D.Multiply(ctm, new Matrix2D(args[start].Number, args[start + 1].Number, args[start + 2].Number, args[start + 3].Number, args[start + 4].Number, args[start + 5].Number)); }
            else if (op == "m" || op == "l") { StartPath(args, ref pathStart); if (args.Count >= 2) AddPoint(ctm, args[args.Count - 2].Number, args[args.Count - 1].Number, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "c") { StartPath(args, ref pathStart); for (int i = Math.Max(0, args.Count - 6); i + 1 < args.Count; i += 2) AddPoint(ctm, args[i].Number, args[i + 1].Number, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "v" || op == "y") { StartPath(args, ref pathStart); for (int i = Math.Max(0, args.Count - 4); i + 1 < args.Count; i += 2) AddPoint(ctm, args[i].Number, args[i + 1].Number, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "re" && args.Count >= 4) { StartPath(args, ref pathStart); int start = args.Count - 4; double x = args[start].Number, y = args[start + 1].Number, width = args[start + 2].Number, height = args[start + 3].Number; AddPoint(ctm, x, y, ref minX, ref minY, ref maxX, ref maxY); AddPoint(ctm, x + width, y, ref minX, ref minY, ref maxX, ref maxY); AddPoint(ctm, x, y + height, ref minX, ref minY, ref maxX, ref maxY); AddPoint(ctm, x + width, y + height, ref minX, ref minY, ref maxX, ref maxY); }
            else if (op == "n") { pathStart = -1; minX = minY = double.MaxValue; maxX = maxY = double.MinValue; }
            else if (IsPathPaintOperator(op)) { if (pathStart >= 0 && maxX > minX && maxY > minY && areas.Any(area => Intersects(area.X, area.Y, area.Width, area.Height, minX, minY, maxX - minX, maxY - minY))) ranges.Add(new RemovalRange(pathStart, opEnd)); pathStart = -1; minX = minY = double.MaxValue; maxX = maxY = double.MinValue; }
            args.Clear();
        }
        return RemoveRanges(content, ranges);
    }

    private static void StartPath(List<ImageContentOperand> args, ref int pathStart) { if (pathStart < 0 && args.Count > 0) pathStart = args[0].Start; }
    private static void AddPoint(Matrix2D transform, double x, double y, ref double minX, ref double minY, ref double maxX, ref double maxY) { var point = transform.Transform(x, y); minX = Math.Min(minX, point.X); minY = Math.Min(minY, point.Y); maxX = Math.Max(maxX, point.X); maxY = Math.Max(maxY, point.Y); }
    private static bool IsPathPaintOperator(string value) => value == "S" || value == "s" || value == "f" || value == "F" || value == "f*" || value == "B" || value == "B*" || value == "b" || value == "b*";
}
