using System.Xml.Linq;

namespace OfficeIMO.Tests;

public partial class PowerPointImageExportTests {
    // Wrapping may split a word over text elements; verify rendered text rather than XML substrings.
    private static string ReadVisibleSvgText(string svg) => string.Concat(XDocument.Parse(svg)
        .Descendants().Where(element => element.Name.LocalName == "text").Select(element => element.Value));
}
