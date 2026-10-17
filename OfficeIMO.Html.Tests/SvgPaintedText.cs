using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

internal static partial class SvgPaintedText {
    internal static PdfEmbeddedFontFamily CreatePinnedFamily() {
        string folder = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts");
        return new PdfEmbeddedFontFamily("Pinned", File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-Regular.ttf")),
            File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-Bold.ttf")),
            File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-Italic.ttf")),
            File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-BoldItalic.ttf")));
    }
}
