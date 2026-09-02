using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Incrementally fingerprints a bounded PDF object graph without materializing encoded payload text.</summary>
internal sealed class PdfObjectGraphFingerprint : IDisposable {
    private const int PayloadChunkSize = 64 * 1024;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
    private readonly Dictionary<(int ObjectNumber, int Generation), int> _references = new();
    private readonly int _maximumDepth;
    private readonly int _maximumNodes;
    private readonly CancellationToken _cancellationToken;
    private readonly byte[] _int32Bytes = new byte[sizeof(int)];
    private readonly byte[] _singleByte = new byte[1];
    private int _nodes;

    internal PdfObjectGraphFingerprint(
        Dictionary<int, PdfIndirectObject> objects,
        int maximumDepth,
        int maximumNodes,
        CancellationToken cancellationToken = default) {
        _objects = objects;
        _maximumDepth = maximumDepth;
        _maximumNodes = maximumNodes;
        _cancellationToken = cancellationToken;
    }

    internal void AppendRoot(PdfObject value) {
        AppendByte(1);
        AppendObject(value, 0);
    }

    internal byte[] Complete() => _hash.GetHashAndReset();

    public void Dispose() => _hash.Dispose();

    private void AppendObject(PdfObject value, int depth) {
        _cancellationToken.ThrowIfCancellationRequested();
        _nodes++;
        if (depth > _maximumDepth || _nodes > _maximumNodes) {
            throw new InvalidDataException("PDF rendering identity exceeded its bounded object graph.");
        }

        switch (value) {
            case PdfNumber number:
                AppendByte(2);
                AppendString(number.Value.ToString("R", CultureInfo.InvariantCulture));
                break;
            case PdfBoolean boolean:
                AppendByte(boolean.Value ? (byte)3 : (byte)4);
                break;
            case PdfName name:
                AppendByte(5);
                AppendString(name.Name);
                break;
            case PdfStringObj text:
                AppendByte(6);
                AppendBytes(text.RawBytes);
                break;
            case PdfArray array:
                AppendByte(7);
                AppendInt32(array.Items.Count);
                for (int i = 0; i < array.Items.Count; i++) AppendObject(array.Items[i], depth + 1);
                break;
            case PdfDictionary dictionary:
                AppendDictionary(8, dictionary, depth);
                break;
            case PdfReference reference:
                AppendReference(reference, depth);
                break;
            case PdfStream stream:
                AppendDictionary(10, stream.Dictionary, depth);
                AppendBytes(stream.Data);
                AppendByte(stream.DecodingFailed ? (byte)16 : (byte)17);
                break;
            case PdfNull:
                AppendByte(11);
                break;
            default:
                throw new InvalidDataException("PDF rendering identity contains an unsupported object type.");
        }
    }

    private void AppendReference(PdfReference reference, int depth) {
        var key = (reference.ObjectNumber, reference.Generation);
        if (_references.TryGetValue(key, out int existingId)) {
            AppendByte(12);
            AppendInt32(existingId);
            return;
        }

        int id = _references.Count;
        _references[key] = id;
        AppendByte(13);
        AppendInt32(id);
        if (PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject? indirect)) {
            AppendByte(14);
            AppendObject(indirect.Value, depth + 1);
        } else {
            AppendByte(15);
        }
    }

    private void AppendDictionary(byte marker, PdfDictionary dictionary, int depth) {
        AppendByte(marker);
        AppendInt32(dictionary.Items.Count);
        foreach (KeyValuePair<string, PdfObject> item in dictionary.Items.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
            AppendString(item.Key);
            AppendObject(item.Value, depth + 1);
        }
    }

    private void AppendString(string value) => AppendBytes(Encoding.UTF8.GetBytes(value));

    private void AppendBytes(byte[] value) {
        AppendInt32(value.Length);
        for (int offset = 0; offset < value.Length; offset += PayloadChunkSize) {
            _cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(PayloadChunkSize, value.Length - offset);
            _hash.AppendData(value, offset, count);
        }
    }

    private void AppendInt32(int value) {
        _int32Bytes[0] = (byte)value;
        _int32Bytes[1] = (byte)(value >> 8);
        _int32Bytes[2] = (byte)(value >> 16);
        _int32Bytes[3] = (byte)(value >> 24);
        _hash.AppendData(_int32Bytes);
    }

    private void AppendByte(byte value) {
        _singleByte[0] = value;
        _hash.AppendData(_singleByte);
    }
}
