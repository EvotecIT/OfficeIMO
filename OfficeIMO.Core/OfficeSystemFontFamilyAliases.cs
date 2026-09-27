using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OfficeIMO.Drawing;

// One installed-font policy for Drawing measurement and document writers. These aliases
// select a platform UI family; they do not assert that every browser uses the same fallback.
internal static class OfficeSystemFontFamilyAliases {
    internal static bool IsSystemUi(string family) =>
        family.Equals("system-ui", StringComparison.OrdinalIgnoreCase) ||
        family.Equals("-apple-system", StringComparison.OrdinalIgnoreCase) ||
        family.Equals("BlinkMacSystemFont", StringComparison.OrdinalIgnoreCase);

    internal static IEnumerable<string> Expand(string family) {
        yield return family;
        if (!IsSystemUi(family)) yield break;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) yield return ".SF NS";
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) yield return "Segoe UI";
        yield return "DejaVu Sans";
        yield return "Liberation Sans";
    }
}
