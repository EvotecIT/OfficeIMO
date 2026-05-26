namespace OfficeIMO.Pdf;

internal static class ResourceResolver {
    public static Dictionary<string, PdfFontResource> GetFontsForPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var fonts = new Dictionary<string, PdfFontResource>(System.StringComparer.Ordinal);
        var dict = GetInheritedDictionary(page, "Resources", objects);
        if (dict is null) return fonts;
        if (!dict.Items.TryGetValue("Font", out var fontDictObj)) return fonts;
        var fontDict = ResolveDict(fontDictObj, objects);
        if (fontDict is null) return fonts;
        foreach (var kv in fontDict.Items) {
            var fontVal = ResolveDict(kv.Value, objects);
            if (fontVal is null) continue;
            string baseFont = (fontVal.Get<PdfName>("BaseFont")?.Name) ?? "";
            string encoding = (fontVal.Get<PdfName>("Encoding")?.Name) ?? "WinAnsiEncoding"; // default heuristic
            bool hasToUnicode = fontVal.Items.ContainsKey("ToUnicode");
            ToUnicodeCMap? cmap = null;
            if (hasToUnicode) {
                if (fontVal.Items.TryGetValue("ToUnicode", out var tu) && tu is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                    var data = Filters.StreamDecoder.Decode(s.Dictionary, s.Data, objects);
                    if (!ToUnicodeCMap.TryParse(data, out cmap)) cmap = null;
                }
            }
            fonts[kv.Key] = new PdfFontResource(kv.Key, baseFont, encoding, hasToUnicode, cmap);
        }
        return fonts;
    }

    public static Dictionary<string, System.Func<byte[], string>> GetFontDecoders(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var map = new Dictionary<string, System.Func<byte[], string>>(System.StringComparer.Ordinal);
        var fonts = GetFontsForPage(page, objects);
        foreach (var kv in fonts) {
            var dec = BuildDecoderForFont(kv.Value);
            map[kv.Key] = dec;
        }
        return map;
    }

    public static Dictionary<string, System.Func<byte[], double>> GetFontWidthProviders(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var map = new Dictionary<string, System.Func<byte[], double>>(System.StringComparer.Ordinal);
        var dict = GetInheritedDictionary(page, "Resources", objects);
        if (dict is null) return map;
        if (!dict.Items.TryGetValue("Font", out var fontDictObj)) return map;
        var fontDict = ResolveDict(fontDictObj, objects);
        if (fontDict is null) return map;
        foreach (var kv in fontDict.Items) {
            var fontVal = ResolveDict(kv.Value, objects);
            if (fontVal is null) continue;
            var subtype = fontVal.Get<PdfName>("Subtype")?.Name ?? string.Empty;
            if (string.Equals(subtype, "Type0", System.StringComparison.Ordinal)) {
                if (TryBuildCidWidthMap(fontVal, objects, out var cidMap)) {
                    var localMap = cidMap!;
                    map[kv.Key] = bytes => SumWidthsCid(bytes, localMap);
                } else {
                    map[kv.Key] = bytes => bytes != null ? (bytes.Length / 2d) * 1000.0 : 0.0; // conservative default
                }
            } else {
                string baseFont = (fontVal.Get<PdfName>("BaseFont")?.Name) ?? "";
                int firstChar = (int)(fontVal.Get<PdfNumber>("FirstChar")?.Value ?? 0);
                var widths = ResolveArray(fontVal.Items.TryGetValue("Widths", out var wobj) ? wobj : null, objects);
                if (widths is null) {
                    if (PdfWriter.TryGetStandardFontByBaseFontName(baseFont, out var standardFont)) {
                        map[kv.Key] = bytes => PdfWriter.EstimateSimpleTextWidth1000(PdfWinAnsiEncoding.Decode(bytes), standardFont);
                    } else {
                        map[kv.Key] = bytes => (bytes?.Length ?? 0) * 500.0;
                    }

                    continue;
                } else {
                    var tbl = new double[widths.Items.Count];
                    for (int i = 0; i < widths.Items.Count; i++) tbl[i] = widths.Items[i] is PdfNumber num ? num.Value : 500.0;
                    map[kv.Key] = bytes => SumWidthsSimple(bytes, firstChar, tbl);
                }
            }
        }
        return map;
    }

    private static double SumWidthsSimple(byte[] bytes, int firstChar, double[] widths) {
        if (bytes == null) return 0.0;
        double sum = 0.0;
        for (int i = 0; i < bytes.Length; i++) {
            int code = bytes[i];
            int idx = code - firstChar;
            if (idx >= 0 && idx < widths.Length) sum += widths[idx]; else sum += 500.0;
        }
        return sum;
    }

    private sealed class CidWidthMap {
        public double DefaultWidth1000 { get; }
        public Dictionary<int, double> Widths { get; }
        public CidWidthMap(double dw, Dictionary<int, double> map) { DefaultWidth1000 = dw; Widths = map; }
    }

    private static bool TryBuildCidWidthMap(PdfDictionary type0Font, Dictionary<int, PdfIndirectObject> objects, out CidWidthMap? map) {
        map = null;
        if (!type0Font.Items.TryGetValue("DescendantFonts", out var dfObj)) return false;
        var dfArr = ResolveArray(dfObj, objects);
        if (dfArr is null || dfArr.Items.Count == 0) return false;
        var desc = ResolveDict(dfArr.Items[0], objects);
        if (desc is null) return false;
        double dw = (desc.Get<PdfNumber>("DW")?.Value) ?? 1000.0;
        var widthsObj = desc.Items.TryGetValue("W", out var w) ? w : null;
        var wArr = ResolveArray(widthsObj, objects);
        var dict = new Dictionary<int, double>();
        if (wArr is not null) {
            // Parse sequences: <startCid> [w1 w2 ...] | <startCid> <endCid> <w>
            for (int i = 0; i < wArr.Items.Count; i++) {
                var startObj = wArr.Items[i] as PdfNumber; if (startObj is null) break;
                int startCid = (int)startObj.Value; i++;
                if (i >= wArr.Items.Count) break;
                var next = wArr.Items[i];
                if (next is PdfArray list) {
                    for (int j = 0; j < list.Items.Count; j++) {
                        if (list.Items[j] is PdfNumber wn) dict[startCid + j] = wn.Value; else dict[startCid + j] = dw;
                    }
                } else if (next is PdfNumber endCidNum) {
                    int endCid = (int)endCidNum.Value; i++;
                    if (i >= wArr.Items.Count) break;
                    var wNum = wArr.Items[i] as PdfNumber; double wv = wNum?.Value ?? dw;
                    for (int cid = startCid; cid <= endCid; cid++) dict[cid] = wv;
                }
            }
        }
        map = new CidWidthMap(dw, dict); return true;
    }

    private static double SumWidthsCid(byte[] bytes, CidWidthMap map) {
        if (bytes == null || bytes.Length == 0) return 0.0;
        double sum = 0.0;
        // Assume Identity-H two-byte big-endian CIDs (common). If odd length, ignore trailing byte.
        for (int i = 0; i + 1 < bytes.Length; i += 2) {
            int cid = (bytes[i] << 8) | bytes[i + 1];
            if (!map.Widths.TryGetValue(cid, out var w)) w = map.DefaultWidth1000;
            sum += w;
        }
        return sum;
    }

    public static Dictionary<string, System.Func<byte[], string>> GetFontDecodersForForm(PdfDictionary formDict, Dictionary<int, PdfIndirectObject> objects) {
        var map = new Dictionary<string, System.Func<byte[], string>>(System.StringComparer.Ordinal);
        if (!formDict.Items.TryGetValue("Resources", out var resObj)) return map;
        var res = ResolveDict(resObj, objects);
        if (res is null) return map;
        if (!res.Items.TryGetValue("Font", out var fontObj)) return map;
        var fontDict = ResolveDict(fontObj, objects);
        if (fontDict is null) return map;
        foreach (var kv in fontDict.Items) {
            var fontVal = ResolveDict(kv.Value, objects);
            if (fontVal is null) continue;
            string baseFont = (fontVal.Get<PdfName>("BaseFont")?.Name) ?? "";
            string encoding = (fontVal.Get<PdfName>("Encoding")?.Name) ?? "WinAnsiEncoding";
            bool hasToUnicode = fontVal.Items.ContainsKey("ToUnicode");
            ToUnicodeCMap? cmap = null;
            if (hasToUnicode) {
                if (fontVal.Items.TryGetValue("ToUnicode", out var tu) && tu is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                    var data = Filters.StreamDecoder.Decode(s.Dictionary, s.Data, objects);
                    if (!ToUnicodeCMap.TryParse(data, out cmap)) cmap = null;
                }
            }
            var resName = kv.Key;
            var dec = BuildDecoderForFont(new PdfFontResource(resName, baseFont, encoding, hasToUnicode, cmap));
            map[resName] = dec;
        }
        return map;
    }

    public static Dictionary<string, byte[]> GetFormXObjectStreams(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var result = new Dictionary<string, byte[]>(System.StringComparer.Ordinal);
        var res = GetInheritedDictionary(page, "Resources", objects);
        if (res is null) return result;
        if (!res.Items.TryGetValue("XObject", out var xoObj)) return result;
        var xo = ResolveDict(xoObj, objects);
        if (xo is null) return result;
        foreach (var kv in xo.Items) {
            if (kv.Value is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                var subtype = s.Dictionary.Get<PdfName>("Subtype")?.Name;
                if (string.Equals(subtype, "Form", System.StringComparison.Ordinal)) {
                    var data = Filters.StreamDecoder.Decode(s.Dictionary, s.Data, objects);
                    result[kv.Key] = data;
                }
            }
        }
        return result;
    }

    public static IReadOnlyList<PdfExtractedImage> GetImageXObjectsForPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects, int pageNumber) {
        var result = new List<PdfExtractedImage>();
        var res = GetInheritedDictionary(page, "Resources", objects);
        if (res is null) return result;
        if (!res.Items.TryGetValue("XObject", out var xoObj)) return result;
        var xo = ResolveDict(xoObj, objects);
        if (xo is null) return result;

        foreach (var kv in xo.Items) {
            int objectNumber = 0;
            PdfStream? stream = null;
            if (kv.Value is PdfReference reference &&
                objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
                indirect.Value is PdfStream referencedStream) {
                objectNumber = reference.ObjectNumber;
                stream = referencedStream;
            } else if (kv.Value is PdfStream directStream) {
                stream = directStream;
            }

            if (stream is null) {
                continue;
            }

            var subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (!string.Equals(subtype, "Image", System.StringComparison.Ordinal)) {
                continue;
            }

            result.Add(BuildExtractedImage(pageNumber, kv.Key, objectNumber, stream, objects));
        }

        return result;
    }

    private static System.Func<byte[], string> BuildDecoderForFont(PdfFontResource font) {
        // Prefer font-specific ToUnicode map when present
        if (font.HasToUnicode && font.CMap is not null) return font.CMap.MapBytes;
        // Fall back to WinAnsi
        return PdfWinAnsiEncoding.Decode;
    }

    private static PdfDictionary? GetInheritedDictionary(PdfDictionary page, string key, Dictionary<int, PdfIndirectObject> objects) {
        PdfDictionary? current = page;
        while (current is not null) {
            if (current.Items.TryGetValue(key, out var v)) {
                var dict = ResolveDict(v, objects);
                if (dict is not null) return dict;
            }
            if (!current.Items.TryGetValue("Parent", out var p) || p is not PdfReference pr || !objects.TryGetValue(pr.ObjectNumber, out var indr) || indr.Value is not PdfDictionary parent) break;
            current = parent;
        }
        return null;
    }

    private static PdfDictionary? ResolveDict(PdfObject obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfDictionary d) return d;
        if (obj is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfDictionary dd) return dd;
        return null;
    }

    private static PdfArray? ResolveArray(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is null) return null;
        if (obj is PdfArray a) return a;
        if (obj is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfArray aa) return aa;
        return null;
    }

    private static PdfExtractedImage BuildExtractedImage(
        int pageNumber,
        string resourceName,
        int objectNumber,
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects) {
        int width = (int)(stream.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int height = (int)(stream.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int bitsPerComponent = (int)(stream.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        string colorSpace = GetNameOrEmpty(stream.Dictionary.Items.TryGetValue("ColorSpace", out var colorSpaceObj) ? colorSpaceObj : null, objects);
        string filter = GetFilterName(stream.Dictionary.Items.TryGetValue("Filter", out var filterObj) ? filterObj : null, objects);

        byte[] bytes = stream.Data;
        string? extension = null;
        string? mimeType = null;
        bool isImageFile = false;

        if (string.Equals(filter, "DCTDecode", System.StringComparison.Ordinal)) {
            extension = "jpg";
            mimeType = "image/jpeg";
            isImageFile = true;
        } else if (string.Equals(filter, "FlateDecode", System.StringComparison.Ordinal) &&
                   TryBuildPngFile(stream, width, height, bitsPerComponent, colorSpace, objects, out var pngBytes)) {
            bytes = pngBytes;
            extension = "png";
            mimeType = "image/png";
            isImageFile = true;
        }

        return new PdfExtractedImage(
            pageNumber,
            resourceName,
            objectNumber,
            width,
            height,
            bitsPerComponent,
            colorSpace,
            filter,
            bytes,
            extension,
            mimeType,
            isImageFile);
    }

    private static string GetFilterName(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        if (resolved is PdfName name) {
            return name.Name;
        }

        if (resolved is PdfArray array) {
            var names = new List<string>();
            foreach (var item in array.Items) {
                var itemResolved = ResolveObject(item, objects);
                if (itemResolved is PdfName itemName) {
                    names.Add(itemName.Name);
                }
            }

            return string.Join(",", names);
        }

        return string.Empty;
    }

    private static string GetNameOrEmpty(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        if (resolved is PdfName name) {
            return name.Name;
        }

        if (resolved is PdfArray array && array.Items.Count > 0) {
            var first = ResolveObject(array.Items[0], objects);
            if (first is PdfName firstName) {
                return firstName.Name;
            }
        }

        return string.Empty;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfReference reference && objects.TryGetValue(reference.ObjectNumber, out var indirect)) {
            return indirect.Value;
        }

        return obj;
    }

    private static bool TryBuildPngFile(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        string colorSpace,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (width <= 0 || height <= 0 || bitsPerComponent != 8) {
            return false;
        }

        int colorType;
        if (string.Equals(colorSpace, "DeviceGray", System.StringComparison.Ordinal)) {
            colorType = 0;
        } else if (string.Equals(colorSpace, "DeviceRGB", System.StringComparison.Ordinal)) {
            colorType = 2;
        } else {
            return false;
        }

        if (stream.Dictionary.Items.ContainsKey("SMask")) {
            return TryBuildPngFileWithSoftMask(stream, width, height, bitsPerComponent, colorType, objects, out pngBytes);
        }

        PdfDictionary? decodeParms = null;
        if (stream.Dictionary.Items.TryGetValue("DecodeParms", out var decodeParmsObj)) {
            decodeParms = ResolveDict(decodeParmsObj, objects);
        }

        int predictor = (int)(decodeParms?.Get<PdfNumber>("Predictor")?.Value ?? 1);
        if (predictor < 10 || predictor > 15) {
            return false;
        }

        using var ms = new MemoryStream();
        WritePngSignature(ms);
        WritePngChunk(ms, "IHDR", BuildIhdr(width, height, bitsPerComponent, colorType));
        WritePngChunk(ms, "IDAT", stream.Data);
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        pngBytes = ms.ToArray();
        return true;
    }

    private static bool TryBuildPngFileWithSoftMask(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        int colorType,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (!stream.Dictionary.Items.TryGetValue("SMask", out var softMaskObj)) {
            return false;
        }

        PdfStream? softMask = ResolveStream(softMaskObj, objects);
        if (softMask is null) {
            return false;
        }

        int softMaskWidth = (int)(softMask.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int softMaskHeight = (int)(softMask.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int softMaskBitsPerComponent = (int)(softMask.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        string softMaskColorSpace = GetNameOrEmpty(softMask.Dictionary.Items.TryGetValue("ColorSpace", out var softMaskColorSpaceObj) ? softMaskColorSpaceObj : null, objects);
        string softMaskFilter = GetFilterName(softMask.Dictionary.Items.TryGetValue("Filter", out var softMaskFilterObj) ? softMaskFilterObj : null, objects);
        if (softMaskWidth != width ||
            softMaskHeight != height ||
            softMaskBitsPerComponent != bitsPerComponent ||
            !string.Equals(softMaskColorSpace, "DeviceGray", System.StringComparison.Ordinal) ||
            !string.Equals(softMaskFilter, "FlateDecode", System.StringComparison.Ordinal)) {
            return false;
        }

        int baseColors;
        int alphaColorType;
        if (colorType == 0) {
            baseColors = 1;
            alphaColorType = 4;
        } else if (colorType == 2) {
            baseColors = 3;
            alphaColorType = 6;
        } else {
            return false;
        }

        byte[] basePixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        byte[] alphaPixels = Filters.StreamDecoder.Decode(softMask.Dictionary, softMask.Data, objects);
        int baseRowLength = width * baseColors;
        int alphaRowLength = width;
        int expectedBaseLength = baseRowLength * height;
        int expectedAlphaLength = alphaRowLength * height;
        if (basePixels.Length < expectedBaseLength || alphaPixels.Length < expectedAlphaLength) {
            return false;
        }

        int outputChannels = baseColors + 1;
        byte[] scanlines = new byte[(1 + width * outputChannels) * height];
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + width * outputChannels);
            int baseRow = row * baseRowLength;
            int alphaRow = row * alphaRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int outputPixel = outputRow + 1 + pixel * outputChannels;
                int basePixel = baseRow + pixel * baseColors;
                for (int channel = 0; channel < baseColors; channel++) {
                    scanlines[outputPixel + channel] = basePixels[basePixel + channel];
                }

                scanlines[outputPixel + baseColors] = alphaPixels[alphaRow + pixel];
            }
        }

        using var ms = new MemoryStream();
        WritePngSignature(ms);
        WritePngChunk(ms, "IHDR", BuildIhdr(width, height, bitsPerComponent, alphaColorType));
        WritePngChunk(ms, "IDAT", DeflateZlibStored(scanlines));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        pngBytes = ms.ToArray();
        return true;
    }

    private static PdfStream? ResolveStream(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        return resolved as PdfStream;
    }

    private static byte[] BuildIhdr(int width, int height, int bitDepth, int colorType) {
        var ihdr = new byte[13];
        WriteInt32BigEndian(ihdr, 0, width);
        WriteInt32BigEndian(ihdr, 4, height);
        ihdr[8] = (byte)bitDepth;
        ihdr[9] = (byte)colorType;
        ihdr[10] = 0;
        ihdr[11] = 0;
        ihdr[12] = 0;
        return ihdr;
    }

    private static void WritePngSignature(Stream stream) {
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        stream.Write(signature, 0, signature.Length);
    }

    private static void WritePngChunk(Stream stream, string type, byte[] data) {
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        var length = new byte[4];
        WriteInt32BigEndian(length, 0, data.Length);
        stream.Write(length, 0, length.Length);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        uint crc = ComputeCrc32(typeBytes, data);
        var crcBytes = new byte[4];
        WriteUInt32BigEndian(crcBytes, 0, crc);
        stream.Write(crcBytes, 0, crcBytes.Length);
    }

    private static void WriteInt32BigEndian(byte[] buffer, int offset, int value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static void WriteUInt32BigEndian(byte[] buffer, int offset, uint value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static byte[] DeflateZlibStored(byte[] data) {
        using var ms = new MemoryStream();
        ms.WriteByte(0x78);
        ms.WriteByte(0x01);

        int offset = 0;
        do {
            int blockLength = Math.Min(65535, data.Length - offset);
            bool final = offset + blockLength >= data.Length;
            ms.WriteByte(final ? (byte)1 : (byte)0);
            ms.WriteByte((byte)(blockLength & 0xFF));
            ms.WriteByte((byte)((blockLength >> 8) & 0xFF));
            ushort nlen = (ushort)~blockLength;
            ms.WriteByte((byte)(nlen & 0xFF));
            ms.WriteByte((byte)((nlen >> 8) & 0xFF));
            ms.Write(data, offset, blockLength);
            offset += blockLength;
        } while (offset < data.Length);

        uint adler = Adler32(data);
        ms.WriteByte((byte)((adler >> 24) & 0xFF));
        ms.WriteByte((byte)((adler >> 16) & 0xFF));
        ms.WriteByte((byte)((adler >> 8) & 0xFF));
        ms.WriteByte((byte)(adler & 0xFF));
        return ms.ToArray();
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private static uint ComputeCrc32(byte[] typeBytes, byte[] data) {
        uint crc = 0xFFFFFFFF;
        for (int i = 0; i < typeBytes.Length; i++) {
            crc = UpdateCrc32(crc, typeBytes[i]);
        }

        for (int i = 0; i < data.Length; i++) {
            crc = UpdateCrc32(crc, data[i]);
        }

        return crc ^ 0xFFFFFFFF;
    }

    private static uint UpdateCrc32(uint crc, byte value) {
        crc ^= value;
        for (int i = 0; i < 8; i++) {
            if ((crc & 1) != 0) {
                crc = (crc >> 1) ^ 0xEDB88320;
            } else {
                crc >>= 1;
            }
        }

        return crc;
    }
}
