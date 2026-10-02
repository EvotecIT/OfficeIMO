using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OfficeIMO.Drawing;

// One installed-font policy for Drawing measurement and document writers. These aliases
// select installed UI or mathematical families; browser fallback order can differ.
internal static class OfficeSystemFontFamilyAliases {
    internal static bool IsSystemUi(string family) =>
        family.Equals("system-ui", StringComparison.OrdinalIgnoreCase) ||
        family.Equals("-apple-system", StringComparison.OrdinalIgnoreCase) ||
        family.Equals("BlinkMacSystemFont", StringComparison.OrdinalIgnoreCase);

    internal static bool IsMath(string family) => family.Equals("math", StringComparison.OrdinalIgnoreCase);

    internal static int MathFamilyRank(string family) {
        int rank = 0;
        foreach (string candidate in Expand("math")) {
            if (string.Equals(family, candidate, StringComparison.OrdinalIgnoreCase)) return rank;
            rank++;
        }
        return int.MaxValue;
    }

    internal static IEnumerable<string> Expand(string family) {
        yield return family;
        if (IsMath(family)) {
            yield return "STIX Two Math";
            yield return "Cambria Math";
            yield return "Latin Modern Math";
            yield return "DejaVu Math TeX Gyre";
            yield break;
        }
        if (!IsSystemUi(family)) yield break;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) yield return ".SF NS";
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) yield return "Segoe UI";
        yield return "DejaVu Sans";
        yield return "Liberation Sans";
    }
}
