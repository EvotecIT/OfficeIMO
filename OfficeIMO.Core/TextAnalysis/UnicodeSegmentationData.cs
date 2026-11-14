#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.TextAnalysis;
#else
namespace OfficeIMO.TextAnalysis;
#endif

/// <summary>Lookup over the generated UAX #29 property ranges (see Build/Generate-UnicodeSegmentationData.py).</summary>
internal static partial class UnicodeSegmentationData {
    /// <summary>Returns the packed Word_Break, Grapheme_Cluster_Break, InCB and Extended_Pictographic value.</summary>
    internal static ushort Lookup(int codePoint) {
        if ((uint)codePoint < (uint)Latin1Values.Length) return Latin1Values[codePoint];
        int low = 0;
        int high = RangeStarts.Length - 1;
        while (low < high) {
            int middle = (low + high + 1) >> 1;
            if (RangeStarts[middle] <= codePoint) low = middle;
            else high = middle - 1;
        }
        return RangeValues[low];
    }
}
