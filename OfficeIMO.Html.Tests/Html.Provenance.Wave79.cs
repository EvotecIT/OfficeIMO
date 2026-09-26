using AngleSharp.Html.Parser;
using System.Threading;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlProvenanceWave79Tests {
    [Theory]
    [InlineData("<table><template><select><option>A</option><tr><td>ignored</td></tr></select></template></table>")]
    [InlineData("<table><caption><select><option>A</option></caption><option>B</option></select></caption></table>")]
    [InlineData("<table><tbody><tr><td><select><option>A</option></tbody><option>B</option></select></td></tr></table>")]
    [InlineData("<table><tr><td><select><option>A</option></tbody><div>B</div></table>")]
    [InlineData("<table><td><select><option>A</option></tr><div>B</div></table>")]
    [InlineData("<table><td><select><option>A</option></tbody><div>B</div></table>")]
    public void SelectPreflightUsesTheParsedElementCountAcrossTableBoundaries(string html) {
        using var parsed = new HtmlParser().ParseDocument(html);
        var options = new OfficeProvenanceOptions { MaxContainerEntries = parsed.All.Length };

        HtmlProvenance.Inspect(html, options);
    }

    [Theory]
    [InlineData("<table><tr><td><select><option>A</option></tbody>")]
    [InlineData("<table><td><select><option>A</option></tbody>")]
    [InlineData("<table><td><select><option>A</option></tr>")]
    public void ImpliedTableSectionsCannotHideFollowingElementsFromPreflight(string prefix) {
        string html = prefix + string.Concat(Enumerable.Repeat("<div>x</div>", 200)) + "</table>";
        using var parsed = new HtmlParser().ParseDocument(html);
        Assert.NotNull(parsed.QuerySelector("tbody"));
        Assert.NotNull(parsed.QuerySelector("tr"));

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.ValidatePotentialElementCountCore(
            html, 50, CancellationToken.None));
    }
}
