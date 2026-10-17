using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

internal static class SvgPaintedText {
    // Kerning may split a word across text elements; outlined runs carry their
    // accessible label on the group which owns the painted glyph paths.
    internal static string Read(string svg) => string.Concat(XDocument.Parse(svg).Descendants()
        .Where(element => element.Name.LocalName == "text" || element.Name.LocalName == "g" &&
            (string?)element.Attribute("role") == "img" && element.Attribute("aria-label") != null &&
            element.Descendants().Any(child => child.Name.LocalName == "path"))
        .Select(element => element.Name.LocalName == "text" ? element.Value : (string?)element.Attribute("aria-label")));

    internal static PdfEmbeddedFontFamily CreatePinnedFamily() {
        string folder = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts");
        return new PdfEmbeddedFontFamily("Pinned", File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-Regular.ttf")),
            File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-Bold.ttf")),
            File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-Italic.ttf")),
            File.ReadAllBytes(Path.Combine(folder, "OfficeIMOBaselineSans-BoldItalic.ttf")));
    }
}
