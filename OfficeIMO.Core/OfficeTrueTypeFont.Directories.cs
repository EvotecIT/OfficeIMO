using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    /// <summary>Standard host locations, including the fonts shipped with installed Office apps on macOS.</summary>
    private static IEnumerable<string> CandidateFontDirectories() {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrEmpty(windows)) yield return Path.Combine(windows, "Fonts");
        yield return "/System/Library/Fonts";
        yield return "/System/Library/Fonts/Supplemental";
        yield return "/Library/Fonts";
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            if (!string.IsNullOrEmpty(home)) yield return Path.Combine(home, "Library", "Fonts");
            foreach (string directory in OfficeInstalledApplicationFonts.GetDirectories()) yield return directory;
        }
        yield return "/usr/share/fonts/truetype/dejavu";
        yield return "/usr/share/fonts/truetype/liberation2";
        yield return "/usr/share/fonts/dejavu";
        yield return "/usr/share/fonts/liberation";
        yield return "/usr/share/fonts/truetype/crosextra";
        yield return "/usr/share/fonts/truetype/carlito";
        yield return "/usr/share/fonts/truetype/caladea";
        yield return "/usr/share/fonts/truetype/inter";
        if (!string.IsNullOrEmpty(home)) {
            yield return Path.Combine(home, ".fonts");
            yield return Path.Combine(home, ".local", "share", "fonts");
        }
    }
}
