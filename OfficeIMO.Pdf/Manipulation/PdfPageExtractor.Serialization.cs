using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal static byte[] SerializePageDictionary(PdfDictionary dictionary, int sourceId, SerializationContext context) {
        var sb = new StringBuilder();
        sb.Append("<< ");
    
        bool hasType = false;
        context.PageOverrides.TryGetValue(sourceId, out var pageOverrides);
        foreach (var entry in dictionary.Items) {
            if (string.Equals(entry.Key, "Parent", StringComparison.Ordinal)) {
                continue;
            }
    
            if (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)) {
                continue;
            }
    
            if (string.Equals(entry.Key, "Type", StringComparison.Ordinal)) {
                hasType = true;
            }
    
            AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
        }
    
        if (!hasType) {
            sb.Append("/Type /Page ");
        }
    
        sb.Append("/Parent ")
            .Append(PdfSyntaxEscaper.IndirectReference(context.PagesObjectId))
            .Append(' ');
    
        if (context.MaterializedPageValues.TryGetValue(sourceId, out var inherited)) {
            foreach (var entry in inherited) {
                if (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)) {
                    continue;
                }
    
                if (!dictionary.Items.ContainsKey(entry.Key)) {
                    AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
                }
            }
        }
    
        if (pageOverrides is not null) {
            foreach (var entry in pageOverrides) {
                AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
            }
        }
    
        sb.Append(">>\n");
        return PdfEncoding.Latin1GetBytes(sb.ToString());
    }
    
    internal static byte[] SerializeObject(PdfObject value, SerializationContext context) {
        if (value is PdfStream stream) {
            return SerializeStream(stream, context);
        }
    
        var sb = new StringBuilder();
        AppendObject(sb, value, context);
        sb.Append('\n');
        return PdfEncoding.Latin1GetBytes(sb.ToString());
    }

    internal static void EnsureSerializedObjectWithinLimit(PdfObject value, SerializationContext context, long maximumBytes) {
        if (maximumBytes < 0 || EstimateSerializedObjectBytes(value, context, maximumBytes) > maximumBytes) {
            throw new InvalidDataException("The rewritten PDF exceeds the configured expanded container limit.");
        }
    }

    private static long EstimateSerializedObjectBytes(PdfObject value, SerializationContext context, long maximumBytes) {
        long total = EstimateValueBytes(value, context, maximumBytes);
        return AddEstimated(total, 1L, maximumBytes);
    }

    private static long EstimateValueBytes(PdfObject value, SerializationContext context, long maximumBytes) {
        switch (value) {
            case PdfStream stream:
                long streamBytes = EstimateDictionaryBytes(stream.Dictionary, context, maximumBytes, excludeLength: true);
                streamBytes = AddEstimated(streamBytes, 48L, maximumBytes);
                return AddEstimated(streamBytes, stream.Data.LongLength, maximumBytes);
            case PdfNumber:
                return 32L;
            case PdfBoolean boolean:
                return boolean.Value ? 4L : 5L;
            case PdfName name:
                return AddEstimated(1L, MultiplyEstimated(name.Name.Length, 5L, maximumBytes), maximumBytes);
            case PdfStringObj text:
                return context.PreserveRawStringBytes
                    ? AddEstimated(MultiplyEstimated(text.RawBytes.LongLength, 2L, maximumBytes), 2L, maximumBytes)
                    : AddEstimated(MultiplyEstimated(text.Value.Length, 4L, maximumBytes), 6L, maximumBytes);
            case PdfNull:
                return 4L;
            case PdfReference reference:
                ValidateReferenceGeneration(reference, context);
                if (!context.NumberMap.ContainsKey(reference.ObjectNumber)) {
                    throw new InvalidOperationException("PDF object " + reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced but not copied.");
                }
                return 32L;
            case PdfArray array:
                long arrayBytes = 3L;
                foreach (PdfObject item in array.Items) {
                    arrayBytes = AddEstimated(arrayBytes, EstimateValueBytes(item, context, maximumBytes), maximumBytes);
                    arrayBytes = AddEstimated(arrayBytes, 1L, maximumBytes);
                }
                return AddEstimated(arrayBytes, 1L, maximumBytes);
            case PdfDictionary dictionary:
                return EstimateDictionaryBytes(dictionary, context, maximumBytes, excludeLength: false);
            default:
                throw new NotSupportedException("Unsupported PDF object type: " + value.GetType().Name);
        }
    }

    private static long EstimateDictionaryBytes(PdfDictionary dictionary, SerializationContext context, long maximumBytes, bool excludeLength) {
        long total = 3L;
        foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
            if (excludeLength && string.Equals(entry.Key, "Length", StringComparison.Ordinal)) continue;
            total = AddEstimated(total, AddEstimated(2L, MultiplyEstimated(entry.Key.Length, 5L, maximumBytes), maximumBytes), maximumBytes);
            total = AddEstimated(total, EstimateValueBytes(entry.Value, context, maximumBytes), maximumBytes);
            total = AddEstimated(total, 1L, maximumBytes);
        }
        return AddEstimated(total, 2L, maximumBytes);
    }

    private static long MultiplyEstimated(long value, long multiplier, long maximumBytes) =>
        value > maximumBytes / multiplier ? ExceededEstimate(maximumBytes) : value * multiplier;

    private static long AddEstimated(long current, long added, long maximumBytes) =>
        current > maximumBytes || added > maximumBytes - current ? ExceededEstimate(maximumBytes) : current + added;

    private static long ExceededEstimate(long maximumBytes) =>
        maximumBytes == long.MaxValue ? long.MaxValue : maximumBytes + 1L;
    
    private static byte[] SerializeStream(PdfStream stream, SerializationContext context) {
        string dictionary = BuildStreamDictionary(stream, context);
        return SerializeStreamBody(dictionary, stream.Data);
    }
    
    private static string BuildStreamDictionary(PdfStream stream, SerializationContext context) {
        var sb = new StringBuilder();
        sb.Append("<< ");
        foreach (var entry in stream.Dictionary.Items) {
            if (!string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
            }
        }
    
        sb.Append("/Length ")
            .Append(stream.Data.Length.ToString(CultureInfo.InvariantCulture))
            .Append(" >>");
    
        return sb.ToString();
    }
    
    private static byte[] SerializeStreamBody(string dictionary, byte[] data) {
        return PdfObjectBytes.WrapStreamBody(dictionary, data);
    }
    
    private static void AppendDictionaryEntry(StringBuilder sb, string key, PdfObject value, SerializationContext context) {
        sb.Append('/').Append(PdfSyntaxEscaper.Name(key)).Append(' ');
        AppendObject(sb, value, context);
        sb.Append(' ');
    }
    
    private static void AppendObject(StringBuilder sb, PdfObject value, SerializationContext context) {
        switch (value) {
            case PdfNumber number:
                sb.Append(FormatNumber(number.Value));
                break;
            case PdfBoolean boolean:
                sb.Append(boolean.Value ? "true" : "false");
                break;
            case PdfName name:
                sb.Append('/').Append(PdfSyntaxEscaper.Name(name.Name));
                break;
            case PdfStringObj text:
                sb.Append(context.PreserveRawStringBytes
                    ? PdfSyntaxEscaper.HexString(text.RawBytes)
                    : text.UseTextStringEncoding
                        ? PdfSyntaxEscaper.TextString(text.Value)
                        : PdfSyntaxEscaper.LiteralString(text.Value));
                break;
            case PdfNull:
                sb.Append("null");
                break;
            case PdfReference reference:
                ValidateReferenceGeneration(reference, context);
                if (!context.NumberMap.TryGetValue(reference.ObjectNumber, out int newObjectNumber)) {
                    throw new InvalidOperationException("PDF object " + reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced but not copied.");
                }

                int generation = context.PreserveReferenceGenerations && newObjectNumber == reference.ObjectNumber
                    ? reference.Generation
                    : 0;
                sb.Append(PdfSyntaxEscaper.IndirectReference(newObjectNumber, generation));
                break;
            case PdfArray array:
                sb.Append("[ ");
                foreach (var item in array.Items) {
                    AppendObject(sb, item, context);
                    sb.Append(' ');
                }
                sb.Append(']');
                break;
            case PdfDictionary dictionary:
                sb.Append("<< ");
                foreach (var entry in dictionary.Items) {
                    AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
                }
                sb.Append(">>");
                break;
            case PdfStream:
                throw new NotSupportedException("Direct PDF streams inside arrays or dictionaries are not supported by page extraction yet.");
            default:
                throw new NotSupportedException("Unsupported PDF object type: " + value.GetType().Name);
        }
    }
    
    private static void ValidateReferenceGeneration(PdfReference reference, SerializationContext context) {
        if (context.SourceObjectGenerations.TryGetValue(reference.ObjectNumber, out int activeGeneration)) {
            if (reference.Generation != activeGeneration) {
                throw BuildGenerationMismatchException(reference, activeGeneration);
            }
    
            return;
        }
    
        if (reference.ObjectNumber < 0 && reference.Generation != 0) {
            throw new InvalidOperationException("Additional PDF object " + reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced with generation " + reference.Generation.ToString(CultureInfo.InvariantCulture) + "; additional rewrite objects must use generation 0.");
        }
    }
    
    private static InvalidOperationException BuildGenerationMismatchException(PdfReference reference, int activeGeneration) {
        return new InvalidOperationException(
            "PDF object " +
            reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) +
            " " +
            reference.Generation.ToString(CultureInfo.InvariantCulture) +
            " R was referenced, but the active object generation is " +
            activeGeneration.ToString(CultureInfo.InvariantCulture) +
            ".");
    }
    
    internal static string BuildInfoDictionary(PdfMetadata metadata) {
        return PdfInfoDictionaryBuilder.Build(metadata);
    }
    
    internal static byte[] WrapObject(int objectNumber, byte[] body) {
        return PdfObjectBytes.WrapIndirectObject(objectNumber, body);
    }
    
    internal static byte[] Assemble(List<byte[]> objects, int catalogId, int infoId, PdfFileVersion fileVersion = PdfFileVersion.Pdf14) {
        return PdfFileAssembler.Assemble(objects, catalogId, infoId, fileVersion);
    }

    internal static PdfFileVersion GetSourceFileVersion(byte[] pdf) {
        return PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(pdf));
    }
    
    private static string FormatNumber(double value) {
        if (Math.Abs(value % 1) < 0.0000001) {
            return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
        }
    
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }
    
}
