using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Tests;

internal static partial class SvgPaintedText {
    // Kerning may split a word across text elements; outlined runs carry their
    // accessible label on the group which owns the painted glyph paths.
    internal static string Read(string svg) => string.Concat(XDocument.Parse(svg).Descendants()
        .Where(element => element.Name.LocalName == "text" || element.Name.LocalName == "g" &&
            (string?)element.Attribute("role") == "img" && element.Attribute("aria-label") != null &&
            element.Descendants().Any(child => child.Name.LocalName == "path"))
        .Select(element => element.Name.LocalName == "text" ? element.Value : (string?)element.Attribute("aria-label")));

}
