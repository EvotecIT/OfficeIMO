using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

/// <summary>Immutable native source fragments; hexadecimal and bin payload bytes are decoded only as a caller reads.</summary>
internal sealed class RtfPayloadSource {
    private readonly Fragment[] _fragments;
    private RtfPayloadSource(Fragment[] fragments) {
        _fragments = fragments;
        ActualLength = fragments.Any(fragment => fragment.Length < 0) ? -1 : fragments.Sum(fragment => (long)fragment.Length);
        DeclaredLength = fragments.Any(fragment => fragment.DeclaredLength < 0) ? -1 : fragments.Sum(fragment => (long)fragment.DeclaredLength);
    }
    internal long ActualLength { get; }
    internal long DeclaredLength { get; }
    internal static RtfPayloadSource FromGroup(RtfGroup group, bool recursive, Action checkCancellation, bool countHex = true) {
        var fragments = new List<Fragment>();
        void Visit(RtfGroup current) {
            foreach (RtfNode node in current.Children) {
                checkCancellation();
                if (node is RtfBinary binary) fragments.Add(new Fragment(binary.RawText, true, binary.Length, binary.DeclaredLength, binary.Position));
                else if (node is RtfControlWord control && control.Name == "bin" && !IsValidBinaryDeclaration(control))
                    fragments.Add(new Fragment(string.Empty, false, 0, -1, control.Position, invalidBinaryDeclaration: true));
                else if (node is RtfText text) {
                    int bytes = -1;
                    if (countHex) {
                        int digits = 0;
                        foreach (char character in text.Text) { checkCancellation(); if (HexValue(character) >= 0) digits++; }
                        bytes = digits / 2;
                    }
                    fragments.Add(new Fragment(text.Text, false, bytes, bytes, text.Position));
                } else if (recursive && node is RtfGroup child) Visit(child);
            }
        }
        Visit(group);
        return new RtfPayloadSource(fragments.ToArray());
    }
    internal void VisitLengths(Action<int, int> visit) {
        foreach (Fragment fragment in _fragments) visit(fragment.Length, fragment.Position);
    }
    internal Stream OpenRead(CancellationToken token) => new PayloadStream(_fragments, token, strict: true);
    internal byte[] ReadBytes() {
        // Semantic getters retain the native binder's compatibility treatment of malformed hex.
        // Explicit complete original-source reads reject missing nibbles or non-whitespace garbage.
        using Stream source = new PayloadStream(_fragments, CancellationToken.None, strict: false);
        return OfficeIMO.Core.Internal.OfficeArchiveSafety.ReadEntryBytes(source, ActualLength, ActualLength, CancellationToken.None);
    }
    private static int HexValue(char value) => value >= '0' && value <= '9' ? value - '0'
        : value >= 'a' && value <= 'f' ? value - 'a' + 10 : value >= 'A' && value <= 'F' ? value - 'A' + 10 : -1;
    private static bool IsValidBinaryDeclaration(RtfControlWord control) {
        // The tokenizer keeps its compatibility parameter value (zero after overflow).
        // Strict source reads must also validate the original declaration rather than losing that fact.
        if (!control.HasParameter || control.RawText.Length < 5) return false;
        string parameter = control.RawText.Substring(4).TrimEnd(' ');
        return int.TryParse(parameter, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int length) && length >= 0;
    }
    private readonly struct Fragment {
        internal Fragment(string text, bool binary, int length, int declaredLength, int position, bool invalidBinaryDeclaration = false) {
            Text = text; Binary = binary; Length = length; DeclaredLength = declaredLength; Position = position;
            InvalidBinaryDeclaration = invalidBinaryDeclaration;
        }
        internal string Text { get; }
        internal bool Binary { get; }
        internal int Length { get; }
        internal int DeclaredLength { get; }
        internal int Position { get; }
        internal bool InvalidBinaryDeclaration { get; }
    }
    private sealed class PayloadStream : Stream {
        private readonly Fragment[] _fragments;
        private readonly CancellationToken _token;
        private readonly bool _strict;
        private int _fragment, _character, _highNibble = -1;
        private bool _disposed;
        internal PayloadStream(Fragment[] fragments, CancellationToken token, bool strict) { _fragments = fragments; _token = token; _strict = strict; }
        public override bool CanRead => !_disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            if (_disposed) throw new ObjectDisposedException(nameof(PayloadStream));
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
            int written = 0;
            while (written < count && _fragment < _fragments.Length) {
                _token.ThrowIfCancellationRequested();
                Fragment fragment = _fragments[_fragment];
                if (_strict && fragment.InvalidBinaryDeclaration) {
                    if (written > 0) return written;
                    throw new InvalidDataException("The original RTF payload contains an invalid bin declaration.");
                }
                if (_character >= fragment.Text.Length) {
                    if (_strict && fragment.Binary && fragment.Length != fragment.DeclaredLength) {
                        if (written > 0) return written;
                        throw new InvalidDataException("The original RTF binary payload is shorter than its bin declaration.");
                    }
                    if (_strict && _highNibble >= 0) {
                        if (written > 0) return written;
                        throw new InvalidDataException("The original RTF hexadecimal payload ends with an incomplete byte.");
                    }
                    _fragment++; _character = 0; _highNibble = -1; continue;
                }
                char character = fragment.Text[_character++];
                if (fragment.Binary) { buffer[offset + written++] = (byte)(character & 255); continue; }
                int nibble = HexValue(character);
                if (nibble < 0) {
                    if (_strict && !char.IsWhiteSpace(character)) {
                        _character--;
                        if (written > 0) return written;
                        throw new InvalidDataException("The original RTF hexadecimal payload contains non-hexadecimal data.");
                    }
                    continue;
                }
                if (_highNibble < 0) _highNibble = nibble;
                else { buffer[offset + written++] = (byte)((_highNibble << 4) | nibble); _highNibble = -1; }
            }
            return written;
        }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { _disposed = true; base.Dispose(disposing); }
    }
}
