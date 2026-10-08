namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private const int CidCodeCount = 65536;
    private const int ExpandedCidWidthRangeLimit = 4096;

    private readonly struct CidWidthEntry {
        internal readonly double Width;
        internal readonly int Order;
        internal CidWidthEntry(double width, int order) { Width = width; Order = order; }
    }

    private readonly struct CidWidthRange {
        internal readonly int First;
        internal readonly int Last;
        internal readonly double Width;
        internal readonly int Order;
        internal CidWidthRange(int first, int last, double width, int order) {
            First = first; Last = last; Width = width; Order = order;
        }
    }

    private sealed class CidWidthMap {
        private readonly double _defaultWidth;
        private readonly Dictionary<int, CidWidthEntry> _entries;
        private readonly double[]? _indexedWidths;

        internal CidWidthMap(double defaultWidth, Dictionary<int, CidWidthEntry> entries, List<CidWidthRange> ranges) {
            _defaultWidth = defaultWidth;
            _entries = entries;
            if (ranges.Count == 0) return;
            if (ranges.Any(range => range.Last - range.First + 1 > ExpandedCidWidthRangeLimit)) {
                _indexedWidths = new double[CidCodeCount];
                for (int cid = 0; cid < CidCodeCount; cid++) _indexedWidths[cid] = defaultWidth;
                foreach (KeyValuePair<int, CidWidthEntry> entry in entries) _indexedWidths[entry.Key] = entry.Value.Width;
            }

            // Resolve newest declarations first. The successor index skips CIDs already
            // resolved, so overlapping ranges visit each 16-bit CID at most once rather
            // than expanding every range or scanning declarations for every glyph.
            int[] next = new int[CidCodeCount + 1];
            for (int cid = 0; cid <= CidCodeCount; cid++) next[cid] = cid;
            for (int index = ranges.Count - 1; index >= 0; index--) {
                CidWidthRange range = ranges[index];
                for (int cid = FindNextCid(next, range.First); cid <= range.Last; cid = FindNextCid(next, cid)) {
                    if (!entries.TryGetValue(cid, out CidWidthEntry entry) || range.Order > entry.Order) {
                        if (_indexedWidths != null) _indexedWidths[cid] = range.Width;
                        else entries[cid] = new CidWidthEntry(range.Width, range.Order);
                    }
                    next[cid] = FindNextCid(next, cid + 1);
                }
            }
        }

        internal double GetWidth(int cid) {
            if (_indexedWidths != null) return _indexedWidths[cid];
            return _entries.TryGetValue(cid, out CidWidthEntry entry) ? entry.Width : _defaultWidth;
        }

        private static int FindNextCid(int[] next, int cid) {
            int root = cid;
            while (next[root] != root) root = next[root];
            while (next[cid] != cid) {
                int successor = next[cid];
                next[cid] = root;
                cid = successor;
            }
            return root;
        }
    }

    private static bool TryBuildCidWidthMap(PdfDictionary type0Font, Dictionary<int, PdfIndirectObject> objects, out CidWidthMap? map) {
        map = null;
        if (!type0Font.Items.TryGetValue("DescendantFonts", out PdfObject? descendants)) return false;
        PdfArray? descendantsArray = ResolveArray(descendants, objects);
        if (descendantsArray is null || descendantsArray.Items.Count == 0) return false;
        PdfDictionary? descendant = ResolveDict(descendantsArray.Items[0], objects);
        if (descendant is null) return false;
        double defaultWidth = descendant.Get<PdfNumber>("DW")?.Value ?? 1000D;
        PdfArray? widths = ResolveArray(descendant.Items.TryGetValue("W", out PdfObject? value) ? value : null, objects);
        Dictionary<int, CidWidthEntry> entries = new Dictionary<int, CidWidthEntry>();
        List<CidWidthRange> ranges = new List<CidWidthRange>();
        if (widths != null) {
            // W contains either start [widths...] or start end constant-width. Only the
            // 16-bit CID domain can be addressed; bounded expansion never truncates a range.
            for (int index = 0; index < widths.Items.Count; index++) {
                int order = index;
                if (widths.Items[index] is not PdfNumber first || first.Value < 0 || first.Value >= CidCodeCount) break;
                int startCid = (int)first.Value;
                if (++index >= widths.Items.Count) break;
                if (widths.Items[index] is PdfArray list) {
                    int count = Math.Min(list.Items.Count, CidCodeCount - startCid);
                    for (int offset = 0; offset < count; offset++) {
                        double width = (list.Items[offset] as PdfNumber)?.Value ?? defaultWidth;
                        entries[startCid + offset] = new CidWidthEntry(width, order);
                    }
                } else if (widths.Items[index] is PdfNumber last) {
                    if (++index >= widths.Items.Count) break;
                    if (last.Value < startCid) continue;
                    int endCid = (int)Math.Min(last.Value, CidCodeCount - 1);
                    double width = (widths.Items[index] as PdfNumber)?.Value ?? defaultWidth;
                    ranges.Add(new CidWidthRange(startCid, endCid, width, order));
                }
            }
        }
        map = new CidWidthMap(defaultWidth, entries, ranges);
        return true;
    }

    private static double SumWidthsCid(byte[] bytes, CidWidthMap map) {
        if (bytes == null || bytes.Length == 0) return 0D;
        double sum = 0D;
        // Identity-H uses two-byte big-endian CIDs; preserve the existing prefix handling.
        for (int index = 0; index + 1 < bytes.Length; index += 2) sum += map.GetWidth((bytes[index] << 8) | bytes[index + 1]);
        return sum;
    }
}
