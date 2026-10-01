#nullable enable
#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

// Replays bounded detection lookahead without reopening or seeking the source.
internal sealed class CsvPrefixTextReader(string prefix, TextReader source) : TextReader
{
    private int _position;

    public override int Read(char[] buffer, int index, int count)
    {
        if (_position == prefix.Length) return source.Read(buffer, index, count);
        int copied = Math.Min(count, prefix.Length - _position);
        prefix.AsSpan(_position, copied).CopyTo(buffer.AsSpan(index, count));
        _position += copied;
        return copied;
    }

    public override ValueTask<int> ReadAsync(Memory<char> buffer, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (_position == prefix.Length) return source.ReadAsync(buffer, cancellationToken);
        int copied = Math.Min(buffer.Length, prefix.Length - _position);
        prefix.AsMemory(_position, copied).CopyTo(buffer);
        _position += copied;
        return new ValueTask<int>(copied);
    }

    public override Task<int> ReadAsync(char[] buffer, int index, int count) =>
        ReadAsync(buffer.AsMemory(index, count)).AsTask();

    protected override void Dispose(bool disposing)
    {
        if (disposing) source.Dispose();
        base.Dispose(disposing);
    }
}
#endif
