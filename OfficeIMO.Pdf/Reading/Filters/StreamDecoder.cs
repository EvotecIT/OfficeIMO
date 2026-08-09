using OfficeIMO.Pdf;
using System.IO;

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

        if (data == null || !dict.Items.TryGetValue("Filter", out var filterObj)) {
            byte[] originalData = data ?? Array.Empty<byte>();
            ThrowIfDecodedLimitExceeded(originalData.LongLength, maxOutputBytes);
            return originalData;
        }

        if (!TryGetFilterNames(filterObj, objects, out List<string> filterNames)) {
            return ReturnWithinDecodedLimit(data, maxOutputBytes);
        }

        byte[] original = data;
        byte[] current = data;
        for (int filterIndex = 0; filterIndex < filterNames.Count; filterIndex++) {
            string filterName = filterNames[filterIndex];
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
        }

        return ReturnWithinDecodedLimit(current, maxOutputBytes);
    }

    public static bool TryDecode(PdfDictionary dict, byte[] data, int maxOutputBytes, out byte[] decoded, Dictionary<int, PdfIndirectObject>? objects = null) {
        return TryDecodeCore(dict, data, maxOutputBytes, out decoded, out _, objects);
    }

    internal static byte[] DecodeRequired(
        PdfDictionary dict,
        byte[] data,
        Dictionary<int, PdfIndirectObject>? objects = null,
        int maxOutputBytes = PdfReadLimits.DefaultMaxDecodedStreamBytes) {
        if (maxOutputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxOutputBytes), maxOutputBytes, "Maximum decoded stream bytes must be positive.");
        }

        if (TryDecodeCore(dict, data, maxOutputBytes, out byte[] decoded, out PdfReadLimitException? limitException, objects)) {
            return decoded;
        }

        if (limitException is not null) {
            throw limitException;
        }

        throw new InvalidDataException("PDF stream could not be decoded using its declared filters.");
    }

    private static bool TryDecodeCore(
        PdfDictionary dict,
        byte[] data,
        int maxOutputBytes,
        out byte[] decoded,
        out PdfReadLimitException? limitException,
        Dictionary<int, PdfIndirectObject>? objects) {
        decoded = Array.Empty<byte>();
        limitException = null;
        if (maxOutputBytes < 0) {
            return false;
        }

        if (data == null || !dict.Items.TryGetValue("Filter", out var filterObj)) {
            byte[] original = data ?? Array.Empty<byte>();
            if (!TryUseOriginal(original, maxOutputBytes, out decoded)) {
                limitException = CreateDecodedLimitException(maxOutputBytes, original.LongLength);
                return false;
            }

            return true;
        }

        if (!TryGetFilterNames(filterObj, objects, out List<string> filterNames) ||
            !HasValidDecodeParmsDeclaration(dict, filterNames.Count, objects)) {
            return false;
        }

        if (filterNames.Count == 0) {
            if (!TryUseOriginal(data, maxOutputBytes, out decoded)) {
                limitException = CreateDecodedLimitException(maxOutputBytes, data.LongLength);
                return false;
            }

            return true;
        }

        byte[] current = data;
        for (int filterIndex = 0; filterIndex < filterNames.Count; filterIndex++) {
            string filterName = filterNames[filterIndex];
            try {
                switch (GetFilterKind(filterName)) {
                    case DecodeFilterKind.Flate:
                        if (!FlateDecoder.TryDecode(current, maxOutputBytes, out current, out bool flateLimitExceeded)) {
                            if (flateLimitExceeded) {
                                limitException = CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                            }

                            return false;
                        }

                        current = ApplyDecodeParms(dict, filterIndex, current, objects, maxOutputBytes);
                        break;
                    case DecodeFilterKind.AsciiHex:
                        if (HasActiveDecodeParms(dict, filterIndex, objects)) {
                            return false;
                        }

                        if (!AsciiHexDecoder.TryDecode(current, maxOutputBytes, out current)) {
                            limitException = CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                            return false;
                        }

                        break;
                    case DecodeFilterKind.Ascii85:
                        if (HasActiveDecodeParms(dict, filterIndex, objects)) {
                            return false;
                        }

                        if (!Ascii85Decoder.TryDecode(current, maxOutputBytes, out current)) {
                            limitException = CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                            return false;
                        }

                        break;
                    case DecodeFilterKind.RunLength:
                        if (!RunLengthDecoder.TryDecode(current, maxOutputBytes, out current)) {
                            limitException = CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                            return false;
                        }

                        break;
                    case DecodeFilterKind.Lzw:
                        if (!LzwDecoder.TryDecode(current, maxOutputBytes, out current, GetEarlyChange(dict, filterIndex, objects))) {
                            limitException = CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
                            return false;
                        }

                        current = ApplyDecodeParms(dict, filterIndex, current, objects, maxOutputBytes);

                        break;
                    default:
                        return false;
                }
            } catch (PdfReadLimitException ex) {
                limitException = ex;
                return false;
            } catch {
                return false;
            }

            if (current.LongLength > maxOutputBytes) {
                limitException = CreateDecodedLimitException(maxOutputBytes, current.LongLength);
                return false;
            }
        }

        decoded = current;
        return true;
    }

    internal static List<string> GetUnsupportedFilters(PdfDictionary dict, Dictionary<int, PdfIndirectObject>? objects = null) {
        if (!dict.Items.TryGetValue("Filter", out var filterObj)) {
            return new List<string>(0);
        }

        if (!TryGetFilterNames(filterObj, objects, out List<string> filterNames)) {
            return new List<string> { "MalformedFilterDeclaration" };
        }

        var unsupported = new List<string>();
        foreach (string filterName in filterNames) {
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

        int predictor = ReadIntegerParameter(decodeParms, "Predictor", 1, objects);
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

        int predictor = ReadIntegerParameter(decodeParms, "Predictor", 1, objects);
        if (predictor == 1) {
            return data;
        }

        if (predictor != 2 && (predictor < 10 || predictor > 15)) {
            throw new FormatException($"Unsupported PDF predictor value '{predictor}'.");
        }

        int columns = ReadPositiveIntegerParameter(decodeParms, "Columns", 1, objects);
        int colors = ReadPositiveIntegerParameter(decodeParms, "Colors", 1, objects);
        int bitsPerComponent = ReadIntegerParameter(decodeParms, "BitsPerComponent", 8, objects);
        if (bitsPerComponent != 1 && bitsPerComponent != 2 && bitsPerComponent != 4 && bitsPerComponent != 8 && bitsPerComponent != 16) {
            throw new FormatException($"Unsupported PDF predictor bit depth '{bitsPerComponent}'.");
        }

        if (predictor == 2) {
            return TiffPredictorDecoder.Decode(data, columns, colors, bitsPerComponent, maxOutputBytes);
        }

        return PngPredictorDecoder.Decode(data, columns, colors, bitsPerComponent, maxOutputBytes);
    }

    private static int GetEarlyChange(PdfDictionary dict, int filterIndex, Dictionary<int, PdfIndirectObject>? objects) {
        var decodeParms = GetDecodeParms(dict, filterIndex, objects);
        if (decodeParms is null) {
            return 1;
        }

        int earlyChange = ReadIntegerParameter(decodeParms, "EarlyChange", 1, objects);
        if (earlyChange != 0 && earlyChange != 1) {
            throw new FormatException($"Unsupported LZW EarlyChange value '{earlyChange}'.");
        }

        return earlyChange;
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
        if (objects is null) {
            return obj;
        }

        HashSet<(int ObjectNumber, int Generation)>? visited = null;
        while (obj is PdfReference reference) {
            visited ??= new HashSet<(int ObjectNumber, int Generation)>();
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                return obj;
            }

            obj = indirect.Value;
        }

        return obj;
    }

    private static bool HasValidDecodeParmsDeclaration(
        PdfDictionary dict,
        int filterCount,
        Dictionary<int, PdfIndirectObject>? objects) {
        if (filterCount == 0) {
            return true;
        }

        if (!dict.Items.TryGetValue("DecodeParms", out PdfObject? decodeParmsObject)) {
            return true;
        }

        PdfObject? resolved = ResolveObject(decodeParmsObject, objects);
        if (resolved is PdfNull) {
            return true;
        }
        if (resolved is PdfDictionary) {
            return filterCount == 1;
        }

        if (resolved is not PdfArray decodeParmsArray || decodeParmsArray.Items.Count != filterCount) {
            return false;
        }

        foreach (PdfObject item in decodeParmsArray.Items) {
            PdfObject? entry = ResolveObject(item, objects);
            if (entry is not PdfNull && entry is not PdfDictionary) {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetFilterNames(
        PdfObject filterObject,
        Dictionary<int, PdfIndirectObject>? objects,
        out List<string> filterNames) {
        filterNames = new List<string>();
        PdfObject? resolved = ResolveObject(filterObject, objects);
        if (resolved is PdfNull) {
            return true;
        }

        if (resolved is PdfName filterName) {
            filterNames.Add(filterName.Name);
            return true;
        }

        if (resolved is not PdfArray filterArray || filterArray.Items.Count == 0) {
            return false;
        }

        foreach (PdfObject item in filterArray.Items) {
            if (ResolveObject(item, objects) is not PdfName arrayFilterName) {
                filterNames.Clear();
                return false;
            }

            filterNames.Add(arrayFilterName.Name);
        }

        return true;
    }

    private static int ReadPositiveIntegerParameter(
        PdfDictionary dictionary,
        string name,
        int defaultValue,
        Dictionary<int, PdfIndirectObject>? objects) {
        int value = ReadIntegerParameter(dictionary, name, defaultValue, objects);
        if (value <= 0) {
            throw new FormatException($"PDF decode parameter '{name}' must be positive.");
        }

        return value;
    }

    private static int ReadIntegerParameter(
        PdfDictionary dictionary,
        string name,
        int defaultValue,
        Dictionary<int, PdfIndirectObject>? objects) {
        if (!dictionary.Items.TryGetValue(name, out PdfObject? parameter)) {
            return defaultValue;
        }

        if (ResolveObject(parameter, objects) is not PdfNumber number ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value) ||
            number.Value != Math.Truncate(number.Value) ||
            number.Value < int.MinValue ||
            number.Value > int.MaxValue) {
            throw new FormatException($"PDF decode parameter '{name}' must be an integer.");
        }

        return (int)number.Value;
    }
}
