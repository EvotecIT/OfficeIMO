using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Filters;

internal static class StreamDecoder {
    private enum DecodeFilterKind {
        Unsupported,
        Flate,
        AsciiHex,
        Ascii85,
        RunLength,
        Lzw
    }

    public static byte[] Decode(
        PdfDictionary dict,
        byte[] data,
        Dictionary<int, PdfIndirectObject>? objects = null,
        int maxOutputBytes = PdfReadLimits.DefaultMaxDecodedStreamBytes) {
        if (maxOutputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxOutputBytes), maxOutputBytes, "Maximum decoded stream bytes must be positive.");
        }

        if (data == null || data.Length == 0 || !dict.Items.TryGetValue("Filter", out var filterObj)) {
            byte[] originalData = data ?? Array.Empty<byte>();
            ThrowIfDecodedLimitExceeded(originalData.LongLength, maxOutputBytes);
            return originalData;
        }

        byte[] original = data;
        byte[] current = data;
        int filterIndex = 0;
        foreach (string filterName in EnumerateFilters(filterObj, objects)) {
            try {
                switch (GetFilterKind(filterName)) {
                    case DecodeFilterKind.Flate:
                        if (!FlateDecoder.TryDecode(current, maxOutputBytes, out current, out bool flateLimitExceeded)) {
                            if (flateLimitExceeded) {
                                throw CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                            }

                            return ReturnWithinDecodedLimit(original, maxOutputBytes);
                        }

                        current = ApplyDecodeParms(dict, filterIndex, current, objects, maxOutputBytes);
                        break;
                    case DecodeFilterKind.AsciiHex:
                        if (!AsciiHexDecoder.TryDecode(current, maxOutputBytes, out current)) {
                            throw CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                        }

                        break;
                    case DecodeFilterKind.Ascii85:
                        if (!Ascii85Decoder.TryDecode(current, maxOutputBytes, out current)) {
                            throw CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                        }

                        break;
                    case DecodeFilterKind.RunLength:
                        if (!RunLengthDecoder.TryDecode(current, maxOutputBytes, out current)) {
                            throw CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                        }

                        break;
                    case DecodeFilterKind.Lzw:
                        if (!LzwDecoder.TryDecode(current, maxOutputBytes, out current, GetEarlyChange(dict, filterIndex, objects))) {
                            throw CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                        }

                        current = ApplyDecodeParms(dict, filterIndex, current, objects, maxOutputBytes);
                        break;
                    default:
                        return ReturnWithinDecodedLimit(original, maxOutputBytes);
                }
                ThrowIfDecodedLimitExceeded(current.LongLength, maxOutputBytes);
            } catch (PdfReadLimitException) {
                throw;
            } catch {
                return ReturnWithinDecodedLimit(original, maxOutputBytes);
            }

            filterIndex++;
        }

        return ReturnWithinDecodedLimit(current, maxOutputBytes);
    }

    public static bool TryDecode(PdfDictionary dict, byte[] data, int maxOutputBytes, out byte[] decoded, Dictionary<int, PdfIndirectObject>? objects = null) {
        decoded = Array.Empty<byte>();
        if (maxOutputBytes < 0) {
            return false;
        }

        if (data == null || data.Length == 0 || !dict.Items.TryGetValue("Filter", out var filterObj)) {
            return TryUseOriginal(data ?? Array.Empty<byte>(), maxOutputBytes, out decoded);
        }

        byte[] current = data;
        int filterIndex = 0;
        foreach (string filterName in EnumerateFilters(filterObj, objects)) {
            try {
                switch (GetFilterKind(filterName)) {
                    case DecodeFilterKind.Flate:
                        if (HasActiveDecodeParms(dict, filterIndex, objects)) {
                            return false;
                        }

                        if (!FlateDecoder.TryDecode(current, maxOutputBytes, out current)) {
                            return false;
                        }

                        break;
                    case DecodeFilterKind.AsciiHex:
                        if (HasActiveDecodeParms(dict, filterIndex, objects)) {
                            return false;
                        }

                        if (!AsciiHexDecoder.TryDecode(current, maxOutputBytes, out current)) {
                            return false;
                        }

                        break;
                    case DecodeFilterKind.Ascii85:
                        if (HasActiveDecodeParms(dict, filterIndex, objects)) {
                            return false;
                        }

                        if (!Ascii85Decoder.TryDecode(current, maxOutputBytes, out current)) {
                            return false;
                        }

                        break;
                    case DecodeFilterKind.RunLength:
                        if (!RunLengthDecoder.TryDecode(current, maxOutputBytes, out current)) {
                            return false;
                        }

                        break;
                    case DecodeFilterKind.Lzw:
                        if (!LzwDecoder.TryDecode(current, maxOutputBytes, out current, GetEarlyChange(dict, filterIndex, objects))) {
                            return false;
                        }

                        break;
                    default:
                        return false;
                }
            } catch {
                return false;
            }

            filterIndex++;
        }

        decoded = current;
        return true;
    }

    internal static List<string> GetUnsupportedFilters(PdfDictionary dict, Dictionary<int, PdfIndirectObject>? objects = null) {
        if (!dict.Items.TryGetValue("Filter", out var filterObj)) {
            return new List<string>(0);
        }

        var unsupported = new List<string>();
        foreach (string filterName in EnumerateFilters(filterObj, objects)) {
            if (!IsSupportedFilter(filterName) && !ContainsFilter(unsupported, filterName)) {
                unsupported.Add(filterName);
            }
        }

        return unsupported;
    }

    internal static bool IsSupportedFilter(string filterName) {
        return GetFilterKind(filterName) != DecodeFilterKind.Unsupported;
    }

    private static DecodeFilterKind GetFilterKind(string filterName) {
        switch (filterName) {
            case "FlateDecode":
            case "Fl":
                return DecodeFilterKind.Flate;
            case "ASCIIHexDecode":
            case "AHx":
                return DecodeFilterKind.AsciiHex;
            case "ASCII85Decode":
            case "A85":
                return DecodeFilterKind.Ascii85;
            case "RunLengthDecode":
            case "RL":
                return DecodeFilterKind.RunLength;
            case "LZWDecode":
            case "LZW":
                return DecodeFilterKind.Lzw;
            default:
                return DecodeFilterKind.Unsupported;
        }
    }

    private static bool ContainsFilter(List<string> filters, string filterName) {
        for (int i = 0; i < filters.Count; i++) {
            if (string.Equals(filters[i], filterName, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static bool TryUseOriginal(byte[] data, int maxOutputBytes, out byte[] decoded) {
        if (!IsWithinLimit(data, maxOutputBytes)) {
            decoded = Array.Empty<byte>();
            return false;
        }

        decoded = data;
        return true;
    }

    private static bool IsWithinLimit(byte[] data, int maxOutputBytes) {
        return data.LongLength <= maxOutputBytes;
    }

    private static void ThrowIfDecodedLimitExceeded(long actual, int maximum) {
        if (actual > maximum) {
            throw CreateDecodedLimitException(maximum, actual);
        }
    }

    private static PdfReadLimitException CreateDecodedLimitException(int maximum, long actual) =>
        PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximum, actual);

    private static byte[] ReturnWithinDecodedLimit(byte[] data, int maximum) {
        ThrowIfDecodedLimitExceeded(data.LongLength, maximum);
        return data;
    }

    private static bool HasActiveDecodeParms(PdfDictionary dict, int filterIndex, Dictionary<int, PdfIndirectObject>? objects) {
        var decodeParms = GetDecodeParms(dict, filterIndex, objects);
        if (decodeParms is null) {
            return false;
        }

        int predictor = (int)(decodeParms.Get<PdfNumber>("Predictor")?.Value ?? 1);
        return predictor > 1;
    }

    private static byte[] ApplyDecodeParms(
        PdfDictionary dict,
        int filterIndex,
        byte[] data,
        Dictionary<int, PdfIndirectObject>? objects,
        int maxOutputBytes) {
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
            return TiffPredictorDecoder.Decode(data, columns, colors, bitsPerComponent, maxOutputBytes);
        }

        if (predictor < 10) {
            return data;
        }

        return PngPredictorDecoder.Decode(data, columns, colors, bitsPerComponent, maxOutputBytes);
    }

    private static int GetEarlyChange(PdfDictionary dict, int filterIndex, Dictionary<int, PdfIndirectObject>? objects) {
        var decodeParms = GetDecodeParms(dict, filterIndex, objects);
        if (decodeParms is null) {
            return 1;
        }

        return (int)(decodeParms.Get<PdfNumber>("EarlyChange")?.Value ?? 1);
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
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
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
