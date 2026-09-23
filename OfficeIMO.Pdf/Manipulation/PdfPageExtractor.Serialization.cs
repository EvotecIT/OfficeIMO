using System.Globalization;

using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal static PdfDictionary BuildPageDictionaryForSizeCheck(PdfDictionary dictionary, int sourceId, SerializationContext context) {
        context.CancellationToken.ThrowIfCancellationRequested();
        var result = new PdfDictionary();
        context.PageOverrides.TryGetValue(sourceId, out var pageOverrides);
        foreach (var entry in dictionary.Items) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (string.Equals(entry.Key, "Parent", StringComparison.Ordinal) ||
                pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)) {
                continue;
            }
            result.Items[entry.Key] = entry.Value;
        }

        if (!result.Items.ContainsKey("Type")) result.Items["Type"] = new PdfName("Page");
        if (context.MaterializedPageValues.TryGetValue(sourceId, out var inherited)) {
            foreach (var entry in inherited) {
                context.CancellationToken.ThrowIfCancellationRequested();
                if (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)) continue;
                if (!result.Items.ContainsKey(entry.Key)) result.Items[entry.Key] = entry.Value;
            }
        }
        if (pageOverrides is not null) {
            foreach (var entry in pageOverrides) {
                context.CancellationToken.ThrowIfCancellationRequested();
                result.Items[entry.Key] = entry.Value;
            }
        }
        return result;
    }

    internal static byte[] SerializePageDictionary(PdfDictionary dictionary, int sourceId, SerializationContext context) {
        context.CancellationToken.ThrowIfCancellationRequested();
        var sb = new StringBuilder();
        sb.Append("<< ");
    
        bool hasType = false;
        context.PageOverrides.TryGetValue(sourceId, out var pageOverrides);
        foreach (var entry in dictionary.Items) {
            context.CancellationToken.ThrowIfCancellationRequested();
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
    
        sb.Append("/Parent ");
        PdfSyntaxEscaper.AppendIndirectReference(sb, context.PagesObjectId);
        sb.Append(' ');
    
        if (context.MaterializedPageValues.TryGetValue(sourceId, out var inherited)) {
            foreach (var entry in inherited) {
                context.CancellationToken.ThrowIfCancellationRequested();
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
                context.CancellationToken.ThrowIfCancellationRequested();
                AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
            }
        }
    
        sb.Append(">>\n");
        return PdfEncoding.Latin1GetBytes(sb);
    }
    
    internal static byte[] SerializeObject(PdfObject value, SerializationContext context) {
        context.CancellationToken.ThrowIfCancellationRequested();
        if (value is PdfStream stream) {
            return SerializeStream(stream, context);
        }
    
        var sb = new StringBuilder();
        AppendObject(sb, value, context);
        sb.Append('\n');
        return PdfEncoding.Latin1GetBytes(sb);
    }

    /// <summary>Serializes an indirect object without buffering a stream body a second time.</summary>
    internal static byte[] SerializeIndirectObject(int objectNumber, PdfObject value, SerializationContext context) =>
        value is PdfStream stream
            ? PdfObjectBytes.WrapStreamObject(objectNumber, BuildStreamDictionary(stream, context), stream)
            : WrapObject(objectNumber, SerializeObject(value, context));

    /// <summary>Serializes an indirect object for final assembly without copying a retained stream payload.</summary>
    internal static PdfSerializedObject SerializeIndirectObjectForAssembly(
        int objectNumber,
        PdfObject value,
        SerializationContext context) =>
        value is PdfStream stream
            ? PdfObjectBytes.SegmentStreamObject(objectNumber, BuildStreamDictionary(stream, context), stream)
            : PdfSerializedObject.FromBytes(WrapObject(objectNumber, SerializeObject(value, context)));

    internal static void EnsureSerializedObjectWithinLimit(PdfObject value, SerializationContext context, long maximumBytes) {
        if (maximumBytes < 0 || CountSerializedObjectBytes(value, context, maximumBytes) > maximumBytes) {
            throw PdfOutputLimitErrors.Create("The rewritten PDF exceeds the configured output limit.");
        }
    }

    internal static void EnsureSerializedIndirectObjectWithinLimit(
        PdfObject value,
        SerializationContext context,
        int objectNumber,
        long maximumBytes) {
        if (maximumBytes < 0) {
            throw PdfOutputLimitErrors.Create("The rewritten PDF exceeds the configured output limit.");
        }
        long total = CountSerializedObjectBytes(value, context, maximumBytes);
        total = AddCounted(total, objectNumber.ToString(CultureInfo.InvariantCulture).Length + 14L, maximumBytes);
        if (maximumBytes < 0 || total > maximumBytes) {
            throw PdfOutputLimitErrors.Create("The rewritten PDF exceeds the configured output limit.");
        }
    }

    private static long CountSerializedObjectBytes(PdfObject value, SerializationContext context, long maximumBytes) {
        if (value is PdfStream stream) return CountStreamBytes(stream, context, maximumBytes);
        return AddCounted(CountValueBytes(value, context, maximumBytes), 1L, maximumBytes);
    }

    private static long CountValueBytes(PdfObject value, SerializationContext context, long maximumBytes) {
        context.CancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfStream:
                throw new NotSupportedException("Direct PDF streams inside arrays or dictionaries are not supported by page extraction yet.");
            case PdfNumber number:
                return FormatNumber(number.Value).Length;
            case PdfBoolean boolean:
                return boolean.Value ? 4L : 5L;
            case PdfName name:
                return AddCounted(1L, CountNameBytes(name.Name, maximumBytes, context.CancellationToken), maximumBytes);
            case PdfStringObj text:
                return context.PreserveRawStringBytes
                    ? CountHexStringBytes(text.RawBytes.LongLength, maximumBytes)
                    : text.UseTextStringEncoding
                        ? CountTextStringBytes(text.Value, maximumBytes, context.CancellationToken)
                        : CountLiteralStringBytes(text.Value, maximumBytes, context.CancellationToken);
            case PdfNull:
                return 4L;
            case PdfReference reference:
                ValidateReferenceGeneration(reference, context);
                if (!context.NumberMap.TryGetValue(reference.ObjectNumber, out int newObjectNumber)) {
                    throw new InvalidOperationException("PDF object " + reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced but not copied.");
                }
                int generation = context.PreserveReferenceGenerations && newObjectNumber == reference.ObjectNumber
                    ? reference.Generation
                    : 0;
                return newObjectNumber.ToString(CultureInfo.InvariantCulture).Length +
                    generation.ToString(CultureInfo.InvariantCulture).Length + 3L;
            case PdfArray array:
                long arrayBytes = 2L;
                foreach (PdfObject item in array.Items) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    arrayBytes = AddCounted(arrayBytes, CountValueBytes(item, context, maximumBytes), maximumBytes);
                    arrayBytes = AddCounted(arrayBytes, 1L, maximumBytes);
                }
                return AddCounted(arrayBytes, 1L, maximumBytes);
            case PdfDictionary dictionary:
                return CountDictionaryBytes(dictionary, context, maximumBytes, excludeLength: false);
            default:
                throw new NotSupportedException("Unsupported PDF object type: " + value.GetType().Name);
        }
    }

    private static long CountStreamBytes(PdfStream stream, SerializationContext context, long maximumBytes) {
        long dictionaryBytes = 3L;
        foreach (KeyValuePair<string, PdfObject> entry in stream.Dictionary.Items) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (string.Equals(entry.Key, "Length", StringComparison.Ordinal)) continue;
            dictionaryBytes = AddCounted(dictionaryBytes, 2L, maximumBytes);
            dictionaryBytes = AddCounted(dictionaryBytes, CountNameBytes(entry.Key, maximumBytes, context.CancellationToken), maximumBytes);
            dictionaryBytes = AddCounted(dictionaryBytes, CountValueBytes(entry.Value, context, maximumBytes), maximumBytes);
            dictionaryBytes = AddCounted(dictionaryBytes, 1L, maximumBytes);
        }
        dictionaryBytes = AddCounted(dictionaryBytes, 8L, maximumBytes); // /Length plus trailing separator.
        dictionaryBytes = AddCounted(dictionaryBytes, stream.DataLength.ToString(CultureInfo.InvariantCulture).Length, maximumBytes);
        dictionaryBytes = AddCounted(dictionaryBytes, 3L, maximumBytes); //  >>
        long total = AddCounted(dictionaryBytes, 8L, maximumBytes); // \nstream\n
        total = AddCounted(total, stream.DataLongLength, maximumBytes);
        return AddCounted(total, 11L, maximumBytes); // \nendstream\n
    }

    private static long CountDictionaryBytes(PdfDictionary dictionary, SerializationContext context, long maximumBytes, bool excludeLength) {
        long total = 3L;
        foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (excludeLength && string.Equals(entry.Key, "Length", StringComparison.Ordinal)) continue;
            total = AddCounted(total, 2L, maximumBytes);
            total = AddCounted(total, CountNameBytes(entry.Key, maximumBytes, context.CancellationToken), maximumBytes);
            total = AddCounted(total, CountValueBytes(entry.Value, context, maximumBytes), maximumBytes);
            total = AddCounted(total, 1L, maximumBytes);
        }
        return AddCounted(total, 2L, maximumBytes);
    }

    private static long CountNameBytes(string value, long maximumBytes, CancellationToken cancellationToken) {
        long asciiBytes = 0L;
        long escapedAsciiBytes = 0L;
        foreach (char character in value) {
            cancellationToken.ThrowIfCancellationRequested();
            if (character >= 0x80) continue;
            asciiBytes++;
            if (character <= 0x20 || character == 0x7F || IsNameDelimiter(character)) escapedAsciiBytes++;
        }

        long utf8Bytes = 0L;
        for (int index = 0; index < value.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = Math.Min(1024, value.Length - index);
            if (index + length < value.Length && char.IsHighSurrogate(value[index + length - 1]) && char.IsLowSurrogate(value[index + length])) length++;
#if NET6_0_OR_GREATER
            utf8Bytes += Encoding.UTF8.GetByteCount(value.AsSpan(index, length));
#else
            utf8Bytes += Encoding.UTF8.GetByteCount(value.Substring(index, length));
#endif
            index += length;
        }
        long encodedNonAsciiBytes = utf8Bytes - asciiBytes;
        long total = AddCounted(asciiBytes, MultiplyCounted(escapedAsciiBytes, 2L, maximumBytes), maximumBytes);
        return AddCounted(total, MultiplyCounted(encodedNonAsciiBytes, 3L, maximumBytes), maximumBytes);
    }

    private static long CountLiteralStringBytes(string value, long maximumBytes, CancellationToken cancellationToken) {
        foreach (char character in value) {
            cancellationToken.ThrowIfCancellationRequested();
            if (character > byte.MaxValue) return CountTextStringBytes(value, maximumBytes, cancellationToken);
        }
        long total = 2L;
        foreach (char character in value) {
            cancellationToken.ThrowIfCancellationRequested();
            long count;
            if (character is '\\' or '(' or ')' or '\r' or '\n' or '\t' or '\b' or '\f') count = 2L;
            else if (character < 32 || character == 127) count = 4L;
            else count = 1L;
            total = AddCounted(total, count, maximumBytes);
        }
        return total;
    }

    private static long CountTextStringBytes(string value, long maximumBytes, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        long encodedBytes = PdfWinAnsiEncoding.CanEncode(value, out _, cancellationToken)
            ? value.Length
            : AddCounted(2L, MultiplyCounted(value.Length, 2L, maximumBytes), maximumBytes);
        cancellationToken.ThrowIfCancellationRequested();
        return CountHexStringBytes(encodedBytes, maximumBytes);
    }

    private static long CountHexStringBytes(long byteCount, long maximumBytes) =>
        AddCounted(MultiplyCounted(byteCount, 2L, maximumBytes), 2L, maximumBytes);

    private static bool IsNameDelimiter(char character) =>
        character is '(' or ')' or '<' or '>' or '[' or ']' or '{' or '}' or '/' or '%' or '#';

    private static long MultiplyCounted(long value, long multiplier, long maximumBytes) =>
        value > maximumBytes / multiplier ? ExceededCount(maximumBytes) : value * multiplier;

    private static long AddCounted(long current, long added, long maximumBytes) =>
        current > maximumBytes || added > maximumBytes - current ? ExceededCount(maximumBytes) : current + added;

    private static long ExceededCount(long maximumBytes) =>
        maximumBytes == long.MaxValue ? long.MaxValue : maximumBytes + 1L;
    
    private static byte[] SerializeStream(PdfStream stream, SerializationContext context) {
        string dictionary = BuildStreamDictionary(stream, context);
        return PdfObjectBytes.WrapStreamBody(dictionary, stream);
    }
    
    private static string BuildStreamDictionary(PdfStream stream, SerializationContext context) {
        context.CancellationToken.ThrowIfCancellationRequested();
        var sb = new StringBuilder();
        sb.Append("<< ");
        foreach (var entry in stream.Dictionary.Items) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
            }
        }
    
        sb.Append("/Length ")
            .Append(stream.DataLength.ToString(CultureInfo.InvariantCulture))
            .Append(" >>");
    
        return sb.ToString();
    }
    
    private static byte[] SerializeStreamBody(string dictionary, byte[] data) {
        return PdfObjectBytes.WrapStreamBody(dictionary, data);
    }
    
    private static void AppendDictionaryEntry(StringBuilder sb, string key, PdfObject value, SerializationContext context) {
        context.CancellationToken.ThrowIfCancellationRequested();
        sb.Append('/');
        PdfSyntaxEscaper.AppendName(sb, key, context.CancellationToken);
        sb.Append(' ');
        AppendObject(sb, value, context);
        sb.Append(' ');
    }
    
    private static void AppendObject(StringBuilder sb, PdfObject value, SerializationContext context) {
        context.CancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfNumber number:
                AppendNumber(sb, number.Value);
                break;
            case PdfBoolean boolean:
                sb.Append(boolean.Value ? "true" : "false");
                break;
            case PdfName name:
                sb.Append('/');
                PdfSyntaxEscaper.AppendName(sb, name.Name, context.CancellationToken);
                break;
            case PdfStringObj text:
                sb.Append(context.PreserveRawStringBytes
                    ? PdfSyntaxEscaper.HexString(text.RawBytes, context.CancellationToken)
                    : text.UseTextStringEncoding
                        ? PdfSyntaxEscaper.TextString(text.Value, context.CancellationToken)
                        : PdfSyntaxEscaper.LiteralString(text.Value, context.CancellationToken));
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
                PdfSyntaxEscaper.AppendIndirectReference(sb, newObjectNumber, generation);
                break;
            case PdfArray array:
                sb.Append("[ ");
                foreach (var item in array.Items) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    AppendObject(sb, item, context);
                    sb.Append(' ');
                }
                sb.Append(']');
                break;
            case PdfDictionary dictionary:
                sb.Append("<< ");
                foreach (var entry in dictionary.Items) {
                    context.CancellationToken.ThrowIfCancellationRequested();
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
        if (context.SourceObjects is not null &&
            context.SourceObjects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? activeObject)) {
            if (reference.Generation != activeObject.Generation) {
                throw BuildGenerationMismatchException(reference, activeObject.Generation);
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
    
    internal static byte[] Assemble(
        List<byte[]> objects,
        int catalogId,
        int infoId,
        PdfFileVersion fileVersion = PdfFileVersion.Pdf14,
        CancellationToken cancellationToken = default) {
        return PdfFileAssembler.Assemble(objects, catalogId, infoId, fileVersion, cancellationToken: cancellationToken);
    }

    internal static byte[] Assemble(
        List<PdfSerializedObject> objects,
        int catalogId,
        int infoId,
        PdfFileVersion fileVersion = PdfFileVersion.Pdf14,
        CancellationToken cancellationToken = default) =>
        PdfFileAssembler.Assemble(objects, catalogId, infoId, fileVersion, cancellationToken: cancellationToken);

    internal static PdfFileVersion GetSourceFileVersion(byte[] pdf) {
        return PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(pdf));
    }
    
    private static string FormatNumber(double value) {
        if (Math.Abs(value % 1) < 0.0000001) {
            return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
        }
    
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static void AppendNumber(StringBuilder destination, double value) {
#if NET6_0_OR_GREATER
        Span<char> buffer = stackalloc char[64];
        if (Math.Abs(value % 1) < 0.0000001) {
            long rounded = (long)Math.Round(value);
            if (rounded.TryFormat(buffer, out int integerWritten, default, CultureInfo.InvariantCulture)) {
                destination.Append(buffer.Slice(0, integerWritten));
                return;
            }
        } else if (value.TryFormat(buffer, out int realWritten, "0.###", CultureInfo.InvariantCulture)) {
            destination.Append(buffer.Slice(0, realWritten));
            return;
        }
#endif
        destination.Append(FormatNumber(value));
    }
    
}
