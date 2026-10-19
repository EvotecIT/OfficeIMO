namespace OfficeIMO.Drawing;

internal static partial class OfficeFaxDecoder {
    /// <summary>Inverts the normative decoder table into sentinel-prefixed wire codes indexed by run length.</summary>
    internal static int[] CreateRunEncodingTable(bool black) {
        var result = new int[104]; // Terminating runs 0..63 and makeup runs 64..2560.
        foreach (var entry in black ? BlackCodes : WhiteCodes) {
            int index = entry.Value < 64 ? entry.Value : 63 + entry.Value / 64;
            result[index] = entry.Key;
        }
        return result;
    }
}
