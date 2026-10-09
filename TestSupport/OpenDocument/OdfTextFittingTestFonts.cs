using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;

namespace OfficeIMO.OpenDocument.Testing;

internal static class OdfTextFittingTestFonts {
    internal const string Family = "Scoped ODF Fitting";
    internal const string Body = "AAAAAAAAAA";

    // Share the repository's generated TrueType fixture; vary only advance widths to expose both report directions.
    internal static byte[] Create(bool wide) {
        byte[] bytes = ManagedTextShapingTestAssets.CreateFont('A');
        int U16(int offset) => bytes[offset] * 256 + bytes[offset + 1];
        int U32(int offset) => checked((int)((uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16 | (uint)bytes[offset + 2] << 8 | bytes[offset + 3]));
        int metrics = 0, count = 0;
        for (int i = 0; i < U16(4); i++) {
            int entry = 12 + i * 16;
            string name = Encoding.ASCII.GetString(bytes, entry, 4);
            if (name == "hmtx") metrics = U32(entry + 8);
            if (name == "hhea") count = U16(U32(entry + 8) + 34);
        }
        if (metrics == 0 || count == 0) throw new InvalidOperationException("Synthetic font metrics are missing.");
        int advance = wide ? 2000 : 200;
        for (int i = 0; i < count; i++) {
            bytes[metrics + i * 4] = (byte)(advance >> 8); bytes[metrics + i * 4 + 1] = (byte)advance;
        }
        return bytes;
    }

    internal static OfficeRenderingProfile Profile(bool wide, IOfficeTextShapingProvider? provider = null) {
        var fonts = new OfficeFontFaceCollection().Add(Family, Create(wide));
        return new OfficeRenderingProfile(wide ? "wide-test" : "narrow-test", fonts, provider, "en-US");
    }
}
