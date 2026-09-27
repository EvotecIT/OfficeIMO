#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvFile
{
    // Inspect only the final encoded character, retaining the same file handle for
    // writing so a competing writer cannot change the boundary after inspection.
    internal static bool NeedsAppendRecordSeparator(Stream stream, Encoding encoding)
    {
        long length = stream.Length;
        byte[] preamble = encoding.GetPreamble();
        if (length == 0) return false;
        byte[] carriageReturn = encoding.GetBytes("\r");
        byte[] lineFeed = encoding.GetBytes("\n");
        int tailLength = (int)Math.Min(length, Math.Max(preamble.Length, Math.Max(carriageReturn.Length, lineFeed.Length)));
        byte[] tail = new byte[tailLength];
        stream.Position = length - tailLength;
        int read = 0;
        while (read < tail.Length)
        {
            int count = stream.Read(tail, read, tail.Length - read);
            if (count == 0) throw new EndOfStreamException("CSV append boundary could not be read.");
            read += count;
        }
        stream.Position = length;
        if (length == preamble.Length && EndsWith(tail, preamble)) return false;
        return !EndsWith(tail, carriageReturn) && !EndsWith(tail, lineFeed);
    }

    private static bool EndsWith(byte[] bytes, byte[] suffix)
    {
        if (suffix.Length == 0 || bytes.Length < suffix.Length) return false;
        int offset = bytes.Length - suffix.Length;
        for (int index = 0; index < suffix.Length; index++)
        {
            if (bytes[offset + index] != suffix[index]) return false;
        }
        return true;
    }

    private static TextWriter CreateAppendTextWriter(string path, CsvSaveOptions options, int bufferSize)
    {
        if (ResolveCompression(options.CompressionType, path) != CsvCompressionType.None)
            throw new NotSupportedException("Appending to compressed CSV files is not supported.");
        Encoding encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var stream = new FileStream(path, options.NoClobber ? FileMode.CreateNew : FileMode.OpenOrCreate,
            FileAccess.ReadWrite, FileShare.Read, bufferSize, FileOptions.SequentialScan);
        try
        {
            bool needsSeparator = NeedsAppendRecordSeparator(stream, encoding);
            var writer = new StreamWriter(stream, encoding, bufferSize);
            return needsSeparator ? new AppendBoundaryWriter(writer, options.NewLine) : writer;
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    private sealed class AppendBoundaryWriter : TextWriter
    {
        private readonly TextWriter _inner;
        private string? _separator;

        internal AppendBoundaryWriter(TextWriter inner, string separator)
        {
            _inner = inner;
            _separator = separator;
        }

        public override Encoding Encoding => _inner.Encoding;

        private void BeginWrite()
        {
            if (_separator is null) return;
            _inner.Write(_separator);
            _separator = null;
        }

        public override void Write(char value) { BeginWrite(); _inner.Write(value); }
        public override void Write(string? value)
        {
            if (string.IsNullOrEmpty(value)) return;
            BeginWrite();
            _inner.Write(value);
        }
        public override void Write(char[] buffer, int index, int count)
        {
            if (count == 0) return;
            BeginWrite();
            _inner.Write(buffer, index, count);
        }
#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<char> buffer)
        {
            if (buffer.IsEmpty) return;
            BeginWrite();
            _inner.Write(buffer);
        }
#endif
        public override void Flush() => _inner.Flush();
        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
