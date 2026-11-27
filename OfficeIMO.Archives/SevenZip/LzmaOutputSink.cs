namespace OfficeIMO.Archives.SevenZip;

/// <summary>Receives completed output from a streaming LZMA decoder and enforces its output limit before any
/// history is grown for it.</summary>
internal abstract class LzmaOutputSink {
    internal long Written { get; private set; }

    /// <summary>Checks that <paramref name="total"/> decoded bytes stay within the limit.</summary>
    internal abstract void Admit(long total);

    internal void Write(byte[] buffer, int offset, int count) {
        Admit(Written + count);
        Consume(buffer, offset, count);
        Written += count;
    }

    protected abstract void Consume(byte[] buffer, int offset, int count);
}
