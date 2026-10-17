using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeIMO.Drawing;

/// <summary>Shared host font locations for installed Office applications; no files are bundled or downloaded.</summary>
internal static class OfficeInstalledApplicationFonts {
    internal static IEnumerable<string> GetDirectories() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) yield break;
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        foreach (string application in new[] { "Microsoft Word.app", "Microsoft Excel.app", "Microsoft PowerPoint.app" }) {
            yield return Path.Combine("/Applications", application, "Contents", "Resources", "DFonts");
            if (!string.IsNullOrEmpty(home))
                yield return Path.Combine(home, "Applications", application, "Contents", "Resources", "DFonts");
        }
    }
}
