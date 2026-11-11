namespace OfficeIMO.Pdf {

    internal static partial class ResourceResolver {
        // A custom encoding stream can use an official-looking CMapName without the official semantics.
        // Only an actual predefined name and compatible descendant ROS select bundled mappings; an embedded
        // CMap contributes its own code-to-CID ranges, which the descendant's Adobe collection then maps to Unicode.
        private static Lazy<PdfPredefinedCMap>? ResolvePredefinedCMap(
            PdfDictionary font,
            Dictionary<int, PdfIndirectObject> objects) {
            if (font.Get<PdfName>("Subtype")?.Name != "Type0" ||
                !font.Items.TryGetValue("Encoding", out PdfObject? encodingObject) ||
                !font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) ||
                ResolveArray(descendantsObject, objects) is not PdfArray { Items.Count: 1 } descendants ||
                ResolveDict(descendants.Items[0], objects) is not PdfDictionary descendant ||
                descendant.Get<PdfName>("Subtype")?.Name is not ("CIDFontType0" or "CIDFontType2") ||
                !descendant.Items.TryGetValue("CIDSystemInfo", out PdfObject? infoObject) ||
                ResolveDict(infoObject, objects) is not PdfDictionary info) return null;
            string? ReadString(string key) => info.Items.TryGetValue(key, out PdfObject? value) &&
                ResolveObject(value, objects) is PdfStringObj text ? text.Value : null;
            string registry = ReadString("Registry") ?? string.Empty;
            string ordering = ReadString("Ordering") ?? string.Empty;
            PdfObject? encoding = ResolveObject(encodingObject, objects);
            if (encoding is PdfName name) return PdfPredefinedCMap.Find(name.Name, registry, ordering);
            if (encoding is not PdfStream stream || !PdfPredefinedCMap.HasUnicodeCollection(registry, ordering)) return null;
            return new Lazy<PdfPredefinedCMap>(() => {
                byte[]? data = Filters.StreamDecoder.Decode(stream, objects);
                return data == null ? PdfPredefinedCMap.Empty : PdfPredefinedCMap.FromEmbeddedCMap(data, ordering);
            });
        }

        private static double SumPredefinedCidWidths(byte[] bytes, CidWidthMap? widths, Lazy<PdfPredefinedCMap> predefined, PdfFontResource font) {
            if (bytes == null || bytes.Length < 2 || bytes.Length % 2 != 0) return 0D;
            double sum = 0D;
            PdfPredefinedCMap mapping = predefined.Value;
            // An embedded CMap this reader cannot use selects no CIDs. Measure as an unmapped composite font does,
            // not as CID 0 for every code, so text decoded through ToUnicode keeps its positions.
            if (ReferenceEquals(mapping, PdfPredefinedCMap.Empty)) return widths != null ? SumWidthsCid(bytes, widths) : bytes.Length / 2D * 1000D;
            for (int index = 0; index < bytes.Length; index += 2) {
                if (!mapping.TryGetPaintedCid(bytes, index, out ushort cid)) throw new PdfUnsupportedTextMappingException(font);
                sum += widths?.GetWidth(cid) ?? 1000D;
            }
            return sum;
        }
    }
}
