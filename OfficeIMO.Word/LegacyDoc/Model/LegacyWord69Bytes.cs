namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Views a bounded portion of the source array without copying its bytes.</summary>
    internal readonly struct LegacyWord69Bytes {
        private readonly byte[] _bytes;
        private readonly int _offset;

        internal LegacyWord69Bytes(byte[] bytes) : this(bytes, 0, bytes.Length) { }

        private LegacyWord69Bytes(byte[] bytes, int offset, int length) {
            _bytes = bytes;
            _offset = offset;
            Length = length;
        }

        internal int Length { get; }
        internal bool IsEmpty => Length == 0;

        internal byte this[int index] {
            get {
                if ((uint)index >= (uint)Length) throw new ArgumentOutOfRangeException(nameof(index));
                return _bytes[_offset + index];
            }
        }

        internal LegacyWord69Bytes Slice(int start) => Slice(start, Length - start);

        internal LegacyWord69Bytes Slice(int start, int length) {
            if ((uint)start > (uint)Length) throw new ArgumentOutOfRangeException(nameof(start));
            if ((uint)length > (uint)(Length - start)) throw new ArgumentOutOfRangeException(nameof(length));
            return new LegacyWord69Bytes(_bytes, _offset + start, length);
        }

        internal int IndexOf(byte value) {
            if (IsEmpty) return -1;
            int index = Array.IndexOf(_bytes, value, _offset, Length);
            return index < 0 ? -1 : index - _offset;
        }
    }
}
