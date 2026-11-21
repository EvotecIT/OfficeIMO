using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfObjectBytes {
    internal static byte[] WrapIndirectObject(int objectNumber, string body) {
        Guard.NotNull(body, nameof(body));
        if (objectNumber < 1) {
            return WrapIndirectObject(objectNumber, PdfEncoding.Latin1GetBytes(body));
        }

        int headerLength = GetObjectHeaderLength(objectNumber);
        // The suffix follows encoded bytes: high code units may also have a CR/LF low byte.
        int lastByte = body.Length == 0 ? -1 : body[body.Length - 1] & 0xFF;
        string suffix = lastByte == '\n' || lastByte == '\r' ? "endobj\n" : "\nendobj\n";
        var result = new byte[headerLength + body.Length + suffix.Length];
        WriteObjectHeader(objectNumber, result, headerLength);
        PdfEncoding.Latin1GetBytes(body, result, headerLength);
        PdfEncoding.Latin1GetBytes(suffix, result, headerLength + body.Length);
        return result;
    }

    /// <summary>Wraps a dictionary still held in a builder, encoding it straight into the object bytes.</summary>
    internal static byte[] WrapIndirectObject(int objectNumber, System.Text.StringBuilder body) {
        Guard.NotNull(body, nameof(body));
#if NET6_0_OR_GREATER
        if (objectNumber >= 1) {
            int headerLength = GetObjectHeaderLength(objectNumber);
            int lastByte = body.Length == 0 ? -1 : body[body.Length - 1] & 0xFF;
            string suffix = lastByte == '\n' || lastByte == '\r' ? "endobj\n" : "\nendobj\n";
            var result = new byte[headerLength + body.Length + suffix.Length];
            WriteObjectHeader(objectNumber, result, headerLength);
            int position = headerLength;
            foreach (ReadOnlyMemory<char> chunk in body.GetChunks()) {
                ReadOnlySpan<char> span = chunk.Span;
                for (int index = 0; index < span.Length; index++) result[position++] = (byte)(span[index] & 0xFF);
            }
            PdfEncoding.Latin1GetBytes(suffix, result, position);
            return result;
        }
#endif
        return WrapIndirectObject(objectNumber, body.ToString());
    }

    // "N 0 obj\n" for a positive object number, written as bytes rather than built as a string per object.
    private static int GetObjectHeaderLength(int objectNumber) {
        int digits = 1;
        for (int value = objectNumber; value >= 10; value /= 10) digits++;
        return digits + " 0 obj\n".Length;
    }

    private static void WriteObjectHeader(int objectNumber, byte[] result, int headerLength) {
        int position = headerLength - " 0 obj\n".Length;
        for (int value = objectNumber; position > 0; value /= 10) {
            result[--position] = (byte)('0' + value % 10);
        }
        PdfEncoding.Latin1GetBytes(" 0 obj\n", result, headerLength - " 0 obj\n".Length);
    }

    internal static byte[] WrapIndirectObject(int objectNumber, byte[] body) {
        return WrapIndirectObject(objectNumber, 0, body);
    }

    internal static byte[] WrapIndirectObjectCancellable(int objectNumber, byte[] body, CancellationToken cancellationToken) {
        Guard.NotNull(body, nameof(body));
        if (objectNumber < 1) throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        cancellationToken.ThrowIfCancellationRequested();
        byte[] prefix = PdfEncoding.Latin1GetBytes(objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
        byte[] suffix = CreateIndirectObjectSuffix(body);
        var result = new byte[checked(prefix.Length + body.Length + suffix.Length)];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        for (int offset = 0; offset < body.Length; offset += 65536) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(body, offset, result, prefix.Length + offset, Math.Min(65536, body.Length - offset));
        }
        cancellationToken.ThrowIfCancellationRequested();
        Buffer.BlockCopy(suffix, 0, result, prefix.Length + body.Length, suffix.Length);
        return result;
    }

    internal static byte[] WrapIndirectObject(int objectNumber, int generation, byte[] body) {
        Guard.NotNull(body, nameof(body));
        if (objectNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        }

        if (generation < 0) {
            throw new ArgumentOutOfRangeException(nameof(generation), "PDF object generation cannot be negative.");
        }

        return Concat(
            PdfEncoding.Latin1GetBytes(objectNumber.ToString(CultureInfo.InvariantCulture) + " " + generation.ToString(CultureInfo.InvariantCulture) + " obj\n"),
            body,
            CreateIndirectObjectSuffix(body));
    }

    // PDF/A requires an EOL before endobj. Preserve the body, including any existing EOL.
    internal static byte[] CreateIndirectObjectSuffix(byte[] body) {
        Guard.NotNull(body, nameof(body));
        bool endsWithEol = body.Length > 0 && (body[body.Length - 1] == '\n' || body[body.Length - 1] == '\r');
        return PdfEncoding.Latin1GetBytes(endsWithEol ? "endobj\n" : "\nendobj\n");
    }

    internal static byte[] WrapStreamObject(int objectNumber, string dictionary, byte[] content) {
        return Concat(CreateStreamObjectSegments(objectNumber, dictionary, content));
    }

    internal static byte[] WrapStreamObject(int objectNumber, string dictionary, PdfStream content, CancellationToken cancellationToken = default) {
        Guard.NotNull(content, nameof(content));
        CreateStreamObjectEnvelope(objectNumber, dictionary, out byte[] prefix, out byte[] suffix, cancellationToken);
        var result = new byte[checked(prefix.Length + content.DataLength + suffix.Length)];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        content.CopyDataTo(result, prefix.Length, cancellationToken);
        Buffer.BlockCopy(suffix, 0, result, prefix.Length + content.DataLength, suffix.Length);
        return result;
    }

    internal static PdfSerializedObject SegmentStreamObject(int objectNumber, string dictionary, PdfStream content, CancellationToken cancellationToken = default) {
        Guard.NotNull(content, nameof(content));
        CreateStreamObjectEnvelope(objectNumber, dictionary, out byte[] prefix, out byte[] suffix, cancellationToken);
        return PdfSerializedObject.FromStream(prefix, content, suffix);
    }

    internal static byte[][] CreateStreamObjectSegments(int objectNumber, string dictionary, byte[] content) {
        Guard.NotNull(content, nameof(content));
        CreateStreamObjectEnvelope(objectNumber, dictionary, out byte[] prefix, out byte[] suffix);
        return new[] {
            prefix,
            content,
            suffix
        };
    }

    private static void CreateStreamObjectEnvelope(
        int objectNumber,
        string dictionary,
        out byte[] prefix,
        out byte[] suffix,
        CancellationToken cancellationToken = default) {
        ValidateStreamDictionary(dictionary, cancellationToken);
        if (objectNumber < 1) throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        if (ContainsStreamMarker(dictionary, cancellationToken)) throw new ArgumentException("Stream dictionaries must not include stream markers.", nameof(dictionary));

        string header = objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n";
        prefix = CreateStreamPrefix(header, dictionary, cancellationToken);
        suffix = PdfEncoding.Latin1GetBytes("\nendstream\nendobj\n");
    }

    internal static byte[] WrapStreamBody(string dictionary, byte[] content) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNullOrWhiteSpace(dictionary, nameof(dictionary));
        if (ContainsStreamMarker(dictionary)) {
            throw new ArgumentException("Stream dictionaries must not include stream markers.", nameof(dictionary));
        }

        return Concat(
            PdfEncoding.Latin1GetBytes(dictionary.TrimEnd() + "\nstream\n"),
            content,
            PdfEncoding.Latin1GetBytes("\nendstream\n"));
    }

    internal static byte[] WrapStreamBody(string dictionary, PdfStream content, CancellationToken cancellationToken = default) {
        Guard.NotNull(content, nameof(content));
        ValidateStreamDictionary(dictionary, cancellationToken);
        if (ContainsStreamMarker(dictionary, cancellationToken)) {
            throw new ArgumentException("Stream dictionaries must not include stream markers.", nameof(dictionary));
        }

        byte[] prefix = CreateStreamPrefix(string.Empty, dictionary, cancellationToken);
        byte[] suffix = PdfEncoding.Latin1GetBytes("\nendstream\n");
        var result = new byte[checked(prefix.Length + content.DataLength + suffix.Length)];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        content.CopyDataTo(result, prefix.Length, cancellationToken);
        Buffer.BlockCopy(suffix, 0, result, prefix.Length + content.DataLength, suffix.Length);
        return result;
    }

    private static byte[] CreateStreamPrefix(string header, string dictionary, CancellationToken cancellationToken) {
        int dictionaryLength = dictionary.Length;
        while (dictionaryLength > 0 && char.IsWhiteSpace(dictionary[dictionaryLength - 1])) {
            cancellationToken.ThrowIfCancellationRequested();
            dictionaryLength--;
        }
        const string streamStart = "\nstream\n";
        byte[] prefix = new byte[checked(header.Length + dictionaryLength + streamStart.Length)];
        int offset = 0;
        for (int i = 0; i < header.Length; i++) prefix[offset++] = (byte)header[i];
        for (int i = 0; i < dictionaryLength; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            prefix[offset++] = (byte)(dictionary[i] & 0xFF);
        }
        cancellationToken.ThrowIfCancellationRequested();
        for (int i = 0; i < streamStart.Length; i++) prefix[offset++] = (byte)streamStart[i];
        return prefix;
    }

    private static void ValidateStreamDictionary(string dictionary, CancellationToken cancellationToken) {
        Guard.NotNull(dictionary, nameof(dictionary));
        bool hasContent = false;
        for (int index = 0; index < dictionary.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (!char.IsWhiteSpace(dictionary[index])) hasContent = true;
        }
        if (!hasContent) throw new ArgumentException("Value cannot be null or whitespace.", nameof(dictionary));
    }

    private static bool ContainsStreamMarker(string dictionary, CancellationToken cancellationToken = default) {
        for (int index = 0; index <= dictionary.Length - 6; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (dictionary[index] != 's' || dictionary[index + 1] != 't' ||
                dictionary[index + 2] != 'r' || dictionary[index + 3] != 'e' ||
                dictionary[index + 4] != 'a' || dictionary[index + 5] != 'm') continue;
            bool before = index == 0 || char.IsWhiteSpace(dictionary[index - 1]);
            int afterIndex = index + 6;
            bool after = afterIndex == dictionary.Length || char.IsWhiteSpace(dictionary[afterIndex]);
            if (before && after) return true;
        }
        return false;
    }

    internal static byte[] Concat(params byte[][] parts) {
        int length = 0;
        foreach (byte[] part in parts) {
            length += part.Length;
        }

        var result = new byte[length];
        int offset = 0;
        foreach (byte[] part in parts) {
            Buffer.BlockCopy(part, 0, result, offset, part.Length);
            offset += part.Length;
        }

        return result;
    }
}
