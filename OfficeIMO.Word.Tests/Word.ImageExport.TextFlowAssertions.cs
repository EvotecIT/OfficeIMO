using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    private static void AssertPaintedTokensAcrossPages(WordDocument document, string prefix, int count) {
        var pages = document.ExportImages(OfficeImageExportFormat.Svg);
        Assert.True(pages.Count > 1, "The authored flow must exercise more than one rendered page.");
        AssertPaintedTokens(pages.Select(page => Encoding.UTF8.GetString(page.Bytes)), prefix, count);
    }

    private static void AssertPaintedTokens(IEnumerable<string> pages, string prefix, int count) {
        XNamespace ns = "http://www.w3.org/2000/svg";
        var tokens = new List<string>();
        foreach (var page in pages) {
            var svg = XDocument.Parse(page);
            // Read painted text, rather than the source paragraph retained by a snapshot.
            // Text elements may split a word at a soft line break; its letters remain ordered.
            string painted = string.Concat(svg.Descendants(ns + "text").Select(item => item.Value));
            tokens.AddRange(Regex.Matches(painted, Regex.Escape(prefix) + @"\d{2}")
                .Cast<Match>().Select(match => match.Value));
        }
        Assert.Equal(Enumerable.Range(1, count).Select(index => prefix + index.ToString("00", CultureInfo.InvariantCulture)), tokens);
    }
}
