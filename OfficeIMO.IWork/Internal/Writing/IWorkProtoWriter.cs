using System.Text;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Encodes the small protobuf subset used by the owned native writer and tracks object references for IWA headers.</summary>
internal sealed class IWorkProtoWriter {
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);
    private readonly MemoryStream _output = new();
    private readonly int _maximumBytes;
    private readonly CancellationToken _cancellationToken;
    private readonly SortedSet<ulong> _references = new();

    internal IWorkProtoWriter(int maximumBytes, CancellationToken cancellationToken) {
        _maximumBytes = maximumBytes;
        _cancellationToken = cancellationToken;
    }

    internal IReadOnlyCollection<ulong> References => _references;
    internal int Length => checked((int)_output.Length);
    internal byte[] ToArray() => _output.ToArray();

    internal IWorkProtoWriter UInt(int field, ulong value) {
        Key(field, 0);
        Ensure(VarintLength(value));
        WriteVarint(_output, value);
        return this;
    }

    internal IWorkProtoWriter Bool(int field, bool value) => UInt(field, value ? 1u : 0u);

    internal IWorkProtoWriter Float(int field, float value) {
        Key(field, 5);
        byte[] bytes = BitConverter.GetBytes(value);
        if (!BitConverter.IsLittleEndian) Array.Reverse(bytes);
        Append(bytes);
        return this;
    }

    internal IWorkProtoWriter Double(int field, double value) {
        Key(field, 1);
        byte[] bytes = BitConverter.GetBytes(value);
        if (!BitConverter.IsLittleEndian) Array.Reverse(bytes);
        Append(bytes);
        return this;
    }

    internal IWorkProtoWriter String(int field, string value) {
        _cancellationToken.ThrowIfCancellationRequested();
        int length = StrictUtf8.GetByteCount(value);
        Ensure(length);
        return Bytes(field, StrictUtf8.GetBytes(value));
    }

    internal IWorkProtoWriter Bytes(int field, byte[] bytes) {
        Key(field, 2);
        Ensure(VarintLength((ulong)bytes.Length));
        WriteVarint(_output, (ulong)bytes.Length);
        Append(bytes);
        return this;
    }

    internal IWorkProtoWriter Message(int field, IWorkProtoWriter value) {
        Key(field, 2);
        Ensure(VarintLength((ulong)value.Length));
        WriteVarint(_output, (ulong)value.Length);
        Ensure(value.Length);
        value._output.WriteTo(_output);
        _references.UnionWith(value._references);
        return this;
    }

    internal IWorkProtoWriter Reference(int field, ulong identifier) {
        Message(field, new IWorkProtoWriter(_maximumBytes, _cancellationToken).UInt(1, identifier));
        _references.Add(identifier);
        return this;
    }

    internal void Append(byte[] bytes) {
        Ensure(bytes.Length);
        _output.Write(bytes, 0, bytes.Length);
    }

    internal void AppendLength(int length) {
        Ensure(VarintLength((ulong)length));
        WriteVarint(_output, (ulong)length);
    }

    private void Key(int field, int wire) {
        if (field <= 0 || field > 536_870_911) throw new ArgumentOutOfRangeException(nameof(field));
        uint key = (uint)field << 3 | (uint)wire;
        Ensure(VarintLength(key));
        WriteVarint(_output, key);
    }

    private void Ensure(int count) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_output.Length > (long)_maximumBytes - count) {
            throw new InvalidDataException($"Native Keynote encoding exceeds the configured {_maximumBytes}-byte limit.");
        }
    }

    internal static void WriteVarint(Stream destination, ulong value) {
        while (value >= 128) {
            destination.WriteByte((byte)(value | 128));
            value >>= 7;
        }
        destination.WriteByte((byte)value);
    }

    private static int VarintLength(ulong value) {
        int count = 1;
        while (value >= 128) { count++; value >>= 7; }
        return count;
    }
}
