using System.IO;

namespace OfficeIMO.Adf;

/// <summary>Stops a JSON writer before its UTF-8 output grows beyond the operation budget.</summary>
internal sealed class AdfJsonOutputStream : MemoryStream {
    private readonly AdfProcessingOptions _options;
    internal AdfJsonOutputStream(AdfProcessingOptions options) => _options = options;
    private void CheckWrite(int count) {
        _options.CancellationToken.ThrowIfCancellationRequested();
        if (Length + count > _options.MaxOutputBytes) throw AdfGraphGuard.Limit("MaxOutputBytes");
    }
    public override void Write(byte[] buffer, int offset, int count) {
        CheckWrite(count);
        base.Write(buffer, offset, count);
    }
#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer) {
        CheckWrite(buffer.Length);
        base.Write(buffer);
    }
#endif
}
