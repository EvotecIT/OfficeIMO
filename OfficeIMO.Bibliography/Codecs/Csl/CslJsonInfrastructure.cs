using System.Buffers;
using System.Text.Json;

namespace OfficeIMO.Bibliography;

internal static class CslJsonInfrastructure {
    internal static JsonDocument ParseCancellable(string value, JsonDocumentOptions options, CancellationToken cancellationToken) =>
        ParseCancellable(EncodeUtf8Cancellable(value, cancellationToken), options, cancellationToken);

    internal static JsonDocument ParseCancellable(byte[] utf8, JsonDocumentOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using var sequence = new CancellationAwareJsonSequence(utf8, 0, utf8.Length, cancellationToken);
        JsonDocument document = JsonDocument.Parse(sequence.Sequence, options);
        if (cancellationToken.IsCancellationRequested) {
            document.Dispose();
            cancellationToken.ThrowIfCancellationRequested();
        }
        return document;
    }

    internal static bool IsStrict(byte[] utf8, CancellationToken cancellationToken) {
        try {
            using JsonDocument parsed = ParseCancellable(utf8, new JsonDocumentOptions { MaxDepth = CslJsonCodec.NativeJsonMaximumDepth }, cancellationToken);
            return true;
        } catch (JsonException) {
            return false;
        }
    }

    internal static byte[] EncodeUtf8Cancellable(string value, CancellationToken cancellationToken) {
        const int ChunkCharacters = 4096;
        using var stream = new MemoryStream(Math.Min(value.Length, 1024 * 1024));
        var encoder = new UTF8Encoding(false, true).GetEncoder();
        var characters = new char[Math.Min(ChunkCharacters, Math.Max(1, value.Length))];
        var bytes = new byte[Encoding.UTF8.GetMaxByteCount(characters.Length)];
        int position = 0;
        while (position < value.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int characterCount = Math.Min(characters.Length, value.Length - position);
            value.CopyTo(position, characters, 0, characterCount);
            bool flush = position + characterCount == value.Length;
            try {
                encoder.Convert(characters, 0, characterCount, bytes, 0, bytes.Length, flush, out int charactersUsed, out int bytesUsed, out _);
                stream.Write(bytes, 0, bytesUsed);
                position += charactersUsed;
            } catch (EncoderFallbackException exception) {
                throw new JsonException("CSL JSON native value contains invalid UTF-16.", exception);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return stream.ToArray();
    }

    internal static void WriteElementCancellable(Utf8JsonWriter writer, JsonElement value, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value.ValueKind) {
            case JsonValueKind.Object:
                writer.WriteStartObject();
                foreach (JsonProperty property in value.EnumerateObject()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    writer.WritePropertyName(property.Name);
                    WriteElementCancellable(writer, property.Value, cancellationToken);
                }
                writer.WriteEndObject();
                break;
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (JsonElement element in value.EnumerateArray()) WriteElementCancellable(writer, element, cancellationToken);
                writer.WriteEndArray();
                break;
            case JsonValueKind.String:
                cancellationToken.ThrowIfCancellationRequested();
                writer.WriteStringValue(value.GetString());
                break;
            case JsonValueKind.Number:
                writer.WriteRawValue(EncodeUtf8Cancellable(value.GetRawText(), cancellationToken), skipInputValidation: true);
                break;
            case JsonValueKind.True:
                writer.WriteBooleanValue(true);
                break;
            case JsonValueKind.False:
                writer.WriteBooleanValue(false);
                break;
            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                writer.WriteNullValue();
                break;
            default:
                throw new JsonException("Unsupported CSL JSON native value kind.");
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    internal sealed class CancellationAwareJsonSequence : IDisposable {
        private const int SegmentSize = 4096;
        private readonly List<CancellationAwareMemoryManager> _managers = new List<CancellationAwareMemoryManager>();

        internal CancellationAwareJsonSequence(byte[] source, int offset, int length, CancellationToken cancellationToken) {
            if (length == 0) { Sequence = ReadOnlySequence<byte>.Empty; return; }
            JsonSequenceSegment? first = null;
            JsonSequenceSegment? last = null;
            int end = checked(offset + length);
            for (int position = offset; position < end;) {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(SegmentSize, end - position);
                var manager = new CancellationAwareMemoryManager(source, position, count, cancellationToken);
                _managers.Add(manager);
                if (first == null) first = last = new JsonSequenceSegment(manager.Memory);
                else last = last!.Append(manager.Memory);
                position += count;
            }
            Sequence = new ReadOnlySequence<byte>(first!, 0, last!, last!.Memory.Length);
        }

        internal ReadOnlySequence<byte> Sequence { get; }

        public void Dispose() {
            foreach (CancellationAwareMemoryManager manager in _managers) ((IDisposable)manager).Dispose();
        }
    }

    private sealed class JsonSequenceSegment : ReadOnlySequenceSegment<byte> {
        internal JsonSequenceSegment(ReadOnlyMemory<byte> memory) => Memory = memory;
        internal JsonSequenceSegment Append(ReadOnlyMemory<byte> memory) {
            var segment = new JsonSequenceSegment(memory) { RunningIndex = RunningIndex + Memory.Length };
            Next = segment;
            return segment;
        }
    }

    private sealed class CancellationAwareMemoryManager : MemoryManager<byte> {
        private readonly byte[] _source;
        private readonly int _offset;
        private readonly int _length;
        private readonly CancellationToken _cancellationToken;

        internal CancellationAwareMemoryManager(byte[] source, int offset, int length, CancellationToken cancellationToken) {
            _source = source;
            _offset = offset;
            _length = length;
            _cancellationToken = cancellationToken;
        }

        public override Span<byte> GetSpan() {
            _cancellationToken.ThrowIfCancellationRequested();
            return new Span<byte>(_source, _offset, _length);
        }

        public override MemoryHandle Pin(int elementIndex = 0) => throw new NotSupportedException();
        public override void Unpin() { }
        protected override void Dispose(bool disposing) { }
    }
}
