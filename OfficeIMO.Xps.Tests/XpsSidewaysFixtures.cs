using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Xps.Tests;

// Real glyph outlines with controlled metric tables exercise the native vertical/fallback contracts.
internal static class XpsSidewaysFixtures {
    internal static int U16(byte[] b, int p) => (b[p] << 8) | b[p + 1];
    internal static int I16(byte[] b, int p) => unchecked((short)U16(b, p));
    private static int U32(byte[] b, int p) => checked((int)(((uint)b[p] << 24) | ((uint)b[p + 1] << 16) | ((uint)b[p + 2] << 8) | b[p + 3]));
    private static void W16(byte[] b, int p, int n) { b[p] = (byte)(n >> 8); b[p + 1] = (byte)n; }
    private static void W32(byte[] b, int p, uint n) { for (int i = 3; i >= 0; i--) { b[p + i] = (byte)n; n >>= 8; } }
    internal static Dictionary<string, byte[]> Tables() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RobotoFlex.ttf"));
        var tables = new Dictionary<string, byte[]>();
        for (int i = 0; i < U16(bytes, 4); i++) {
            int p = 12 + i * 16; string tag = Encoding.ASCII.GetString(bytes, p, 4);
            tables[tag] = bytes.Skip(U32(bytes, p + 8)).Take(U32(bytes, p + 12)).ToArray();
        }
        return tables;
    }
    internal static byte[] Font(string metrics) {
        var tables = Tables();
        if (metrics == "hhea") tables.Remove("OS/2");
        if (metrics == "os2-distinct") { W16(tables["OS/2"], 68, 2100); W16(tables["OS/2"], 70, -450); }
        if (metrics.StartsWith("vertical", StringComparison.Ordinal)) {
            int glyphs = U16(tables["maxp"], 4);
            var header = new byte[36]; W32(header, 0, 0x00010000); W16(header, 34, 1);
            // A single long metric followed by distinct short bearings tests the compressed vmtx layout.
            var values = new byte[4 + (glyphs - 1) * 2];
            W16(values, 0, 2300);
            for (int g = 0; g < glyphs; g++) W16(values, 2 + g * 2, 100 + g);
            if (metrics == "vertical-truncated") Array.Resize(ref values, 4);
            if (metrics == "vertical-zero-count") W16(header, 34, 0);
            tables["vhea"] = header; tables["vmtx"] = values;
        }
        return Pack(tables);
    }
    private static uint Checksum(byte[] data) {
        uint sum = 0;
        for (int p = 0; p < data.Length; p += 4) {
            uint word = 0;
            for (int i = 0; i < 4; i++) word = (word << 8) | (p + i < data.Length ? data[p + i] : 0U);
            sum = unchecked(sum + word);
        }
        return sum;
    }
    private static byte[] Pack(Dictionary<string, byte[]> tables) {
        W32(tables["head"], 8, 0);
        int offset = 12 + tables.Count * 16;
        var result = new byte[offset + tables.Values.Sum(t => (t.Length + 3) & ~3)];
        W32(result, 0, 0x00010000); W16(result, 4, tables.Count);
        int power = 1, selector = 0; while (power * 2 <= tables.Count) { power *= 2; selector++; }
        W16(result, 6, power * 16); W16(result, 8, selector); W16(result, 10, tables.Count * 16 - power * 16);
        int record = 12, head = 0;
        foreach (var table in tables.OrderBy(t => t.Key, StringComparer.Ordinal)) {
            Encoding.ASCII.GetBytes(table.Key).CopyTo(result, record);
            W32(result, record + 4, Checksum(table.Value)); W32(result, record + 8, (uint)offset); W32(result, record + 12, (uint)table.Value.Length);
            table.Value.CopyTo(result, offset); if (table.Key == "head") head = offset;
            offset += (table.Value.Length + 3) & ~3; record += 16;
        }
        W32(result, head + 8, unchecked(0xB1B0AFBAU - Checksum(result)));
        return result;
    }
    internal static XpsDocument Create(string metrics, XpsFormat format = XpsFormat.Xps, bool positioned = false) {
        var doc = XpsDocument.Create(format);
        var page = doc.AddPage(260, 160).AddPath("M0,0H260V160H0Z", "#FFFFFFFF");
        page.AddText("AFQ", doc.AddFont(Font(metrics), false), 48, 30, 80);
        var xml = page.GetMarkup(); var glyphs = xml.Elements().Last();
        glyphs.SetAttributeValue("IsSideways", "true");
        if (positioned) {
            glyphs.SetAttributeValue("Indices", ";,100,30,10;");
            glyphs.SetAttributeValue("RenderTransform", "0,1,-1,0,160,0");
            glyphs.SetAttributeValue("Clip", "M0,0H190V130H0Z");
        }
        page.ReplaceMarkup(xml); return doc;
    }
}
