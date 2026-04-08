using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Filters;

internal static class StreamDecoder {
    public static byte[] Decode(PdfDictionary dict, byte[] data, Dictionary<int, PdfIndirectObject>? objects = null) {
        if (data == null || data.Length == 0 || !dict.Items.TryGetValue("Filter", out var filterObj)) {
            return data ?? Array.Empty<byte>();
        }

        byte[] original = data;
        byte[] current = data;
        int filterIndex = 0;
        foreach (string filterName in EnumerateFilters(filterObj, objects)) {
            try {
                switch (filterName) {
                    case "FlateDecode":
                    case "Fl":
                        current = FlateDecoder.Decode(current);
                        current = ApplyDecodeParms(dict, filterIndex, current, objects);
                        break;
                    case "ASCIIHexDecode":
                    case "AHx":
                        current = AsciiHexDecoder.Decode(current);
                        break;
                    case "ASCII85Decode":
                    case "A85":
                        current = Ascii85Decoder.Decode(current);
                        break;
                    case "RunLengthDecode":
                    case "RL":
                        current = RunLengthDecoder.Decode(current);
                        break;
                    default:
                        return original;
                }
            } catch {
                return original;
            }

            filterIndex++;
        }

        return current;
    }

    private static byte[] ApplyDecodeParms(PdfDictionary dict, int filterIndex, byte[] data, Dictionary<int, PdfIndirectObject>? objects) {
        var decodeParms = GetDecodeParms(dict, filterIndex, objects);
        if (decodeParms is null) {
            return data;
        }

        int predictor = (int)(decodeParms.Get<PdfNumber>("Predictor")?.Value ?? 1);
        if (predictor <= 1) {
            return data;
        }

        int columns = (int)(decodeParms.Get<PdfNumber>("Columns")?.Value ?? 1);
        int colors = (int)(decodeParms.Get<PdfNumber>("Colors")?.Value ?? 1);
        int bitsPerComponent = (int)(decodeParms.Get<PdfNumber>("BitsPerComponent")?.Value ?? 8);
        if (predictor == 2) {
            return TiffPredictorDecoder.Decode(data, columns, colors, bitsPerComponent);
        }

        if (predictor < 10) {
            return data;
        }

        return PngPredictorDecoder.Decode(data, columns, colors, bitsPerComponent);
    }

    private static PdfDictionary? GetDecodeParms(PdfDictionary dict, int filterIndex, Dictionary<int, PdfIndirectObject>? objects) {
        if (!dict.Items.TryGetValue("DecodeParms", out var decodeParmsObj)) {
            return null;
        }

        PdfObject? resolvedDecodeParms = ResolveObject(decodeParmsObj, objects);

        if (resolvedDecodeParms is PdfDictionary directDict) {
            return filterIndex == 0 ? directDict : null;
        }

        if (resolvedDecodeParms is PdfArray decodeParmsArray &&
            filterIndex >= 0 &&
            filterIndex < decodeParmsArray.Items.Count &&
            ResolveDictionary(decodeParmsArray.Items[filterIndex], objects) is PdfDictionary indexedDict) {
            return indexedDict;
        }

        return null;
    }

    private static PdfDictionary? ResolveDictionary(PdfObject? obj, Dictionary<int, PdfIndirectObject>? objects) {
        if (ResolveObject(obj, objects) is PdfDictionary directDictionary) {
            return directDictionary;
        }

        return null;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject>? objects) {
        if (obj is PdfReference reference &&
            objects is not null &&
            objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
            indirect.Value is PdfObject resolvedObject) {
            return resolvedObject;
        }

        return obj;
    }

    private static IEnumerable<string> EnumerateFilters(PdfObject filterObj, Dictionary<int, PdfIndirectObject>? objects) {
        if (ResolveObject(filterObj, objects) is PdfName filterName) {
            yield return filterName.Name;
            yield break;
        }

        if (ResolveObject(filterObj, objects) is PdfArray filterArray) {
            foreach (var item in filterArray.Items) {
                if (ResolveObject(item, objects) is PdfName arrayFilterName) {
                    yield return arrayFilterName.Name;
                }
            }
        }
    }
}
