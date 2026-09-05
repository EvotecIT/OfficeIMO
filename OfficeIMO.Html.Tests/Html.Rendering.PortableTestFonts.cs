using System;
using System.Linq;
using OfficeIMO.TestAssets;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private static string CreatePortableEmbeddedFontFaceCss(string familyName, params int[] additionalScalars) {
        int[] scalars = Enumerable.Range(32, 95)
            .Concat(additionalScalars ?? Array.Empty<int>())
            .Distinct()
            .ToArray();
        byte[] fontData = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(scalars);
        return "@font-face{font-family:'" + familyName + "';src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(fontData)
            + "') format('truetype')}";
    }
}
