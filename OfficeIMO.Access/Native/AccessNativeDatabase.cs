using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    /// <summary>Immutable bounded source snapshot and shared Jet3/Jet4/ACE page substrate. No providers or external references are opened.</summary>
    internal sealed partial class AccessNativeDatabase : IDisposable {
        private byte[] _bytes;
        private readonly AccessDocument _document;
        private readonly Dictionary<int, AccessNativeTable> _definitions = new Dictionary<int, AccessNativeTable>();
        private long _metadataBytes;
        internal void AccountMetadata(int bytes) {
            if (bytes < 0 || _metadataBytes > MaxMetadataBytes - (long)bytes) throw new InvalidDataException("Native Access decoded metadata exceeds MaxMetadataBytes.");
            _metadataBytes += bytes;
        }
        internal readonly int MaxCatalogObjects, MaxMetadataBytes, MaxValueBytes, MaxChainLength;
        internal readonly long MaxRows;
        internal readonly HashSet<string>? SelectedTables;
        internal readonly bool DecodeApplicationObjects;
        internal readonly AccessNativeLayout Layout;
        internal AccessNativeDatabase(AccessDocument document, byte[] bytes, AccessLoadOptions options) {
            Layout = new AccessNativeLayout(document.Profile == AccessFormatProfile.Jet3);
            _document = document; _bytes = bytes; MaxCatalogObjects = options.MaxCatalogObjects; MaxMetadataBytes = options.MaxMetadataBytes;
            MaxValueBytes = options.MaxValueBytes; MaxChainLength = options.MaxChainLength; MaxRows = options.MaxRows;
            SelectedTables = options.TableNames == null ? null : new HashSet<string>(options.TableNames, StringComparer.OrdinalIgnoreCase);
            DecodeApplicationObjects = options.DecodeApplicationObjects;
        }
        internal int PageCount => _bytes.Length / Layout.PageSize;
        internal AccessDocument Document => _document;
        internal byte[] Snapshot() { _document.EnsureNotDisposed(); return _bytes; }
        internal OfficeByteView Page(int page, byte? expectedType = null) {
            _document.EnsureNotDisposed();
            if (page <= 0 || page >= PageCount) throw new InvalidDataException("Native Access page reference is outside the source snapshot.");
            OfficeByteView data = new OfficeByteView(_bytes).Slice(checked(page * Layout.PageSize), Layout.PageSize);
            if (expectedType.HasValue && data[0] != expectedType.Value) throw new InvalidDataException("Native Access page type does not match its reference.");
            return data;
        }
        internal bool CanDecode(out string reason) {
            byte[] header = new OfficeByteView(_bytes).Slice(0, Layout.PageSize).ToArray();
            OfficeRc4Transform mask = new OfficeRc4Transform(new byte[] { 0xc7, 0xda, 0x39, 0x6b });
            for (int i = 24; i < (Layout.IsJet3 ? 150 : 152); i++) header[i] ^= mask.NextByte();
            _document.CodePage = U16(header, 60); _document.SortOrder = U16(header, Layout.IsJet3 ? 58 : 110);
            if (U32(header, 62) != 0) { reason = "Native Access page encryption is not qualified; no encrypted pages are decoded."; return false; }
            if (Layout.IsJet3) {
                if (header.Skip(66).Take(20).Any(value => value != 0)) { reason = "Native Access password protection is not qualified; only inert header evidence is available."; return false; }
                if (!QualifiedJet3CodePage(_document.CodePage!.Value)) { reason = "Jet3 catalog decoding is not qualified for this code page; the snapshot remains preserve-only."; return false; }
                reason = string.Empty; return true;
            }
            double date = F64(header, 114);
            if (double.IsNaN(date) || double.IsInfinity(date) || date < int.MinValue || date > int.MaxValue) throw new InvalidDataException("Native Access header creation date is invalid.");
            int passwordMask = (int)date;
            for (int i = 0; i < 40; i++) {
                if ((byte)(header[66 + i] ^ (byte)(passwordMask >> ((i % 4) * 8))) != 0) {
                    reason = "Native Access password protection is not qualified; only inert header evidence is available."; return false;
                }
            }
            reason = string.Empty; return true;
        }
        internal OfficeByteView Row(int pageNumber, int rowNumber, bool followOverflow, CancellationToken cancellation) {
            HashSet<long> visited = new HashSet<long>();
            for (int depth = 0; ; depth++) {
                cancellation.ThrowIfCancellationRequested();
                if (depth >= MaxChainLength || !visited.Add(((long)pageNumber << 8) | (uint)rowNumber)) throw new InvalidDataException("Native Access overflow chain exceeds its limit or contains a cycle.");
                OfficeByteView page = Page(pageNumber, 1); int directory = Layout.DataRowDirectory, count = U16(page, Layout.DataRowCount);
                if (count > 255 || rowNumber < 0 || rowNumber >= count) throw new InvalidDataException("Native Access row reference is invalid.");
                int slot = U16(page, directory + rowNumber * 2); int start = slot & 0x1fff;
                int end = rowNumber == 0 ? Layout.PageSize : U16(page, directory + (rowNumber - 1) * 2) & 0x1fff;
                if (start < directory + count * 2 || end < start || end > Layout.PageSize) throw new InvalidDataException("Native Access row directory is malformed.");
                OfficeByteView row = Slice(page, start, end - start);
                if (!followOverflow || (slot & 0x4000) == 0) return row;
                uint pointer = U32(row, 0); rowNumber = (int)(pointer & 255); pageNumber = checked((int)(pointer >> 8));
            }
        }
        internal IEnumerable<int> OwnedPages(uint pointer, CancellationToken cancellation) {
            OfficeByteView map = Row(checked((int)(pointer >> 8)), (int)(pointer & 255), false, cancellation);
            if (map.Length < 5) throw new InvalidDataException("Native Access usage map is truncated.");
            if (map[0] == 0) {
                uint start = U32(map, 1);
                foreach (int page in MapBits(Slice(map, 5, map.Length - 5), start, cancellation)) yield return page;
            } else if (map[0] == 1) {
                if ((map.Length - 1) % 4 != 0) throw new InvalidDataException("Native Access reference usage map is malformed.");
                HashSet<int> visited = new HashSet<int>();
                for (int position = 1; position < map.Length; position += 4) {
                    cancellation.ThrowIfCancellationRequested(); int reference = I32(map, position); if (reference == 0) continue;
                    if (!visited.Add(reference)) throw new InvalidDataException("Native Access reference usage map repeats a page.");
                    OfficeByteView page = Page(reference, 5); int mapBytes = Layout.PageSize - 4;
                    uint start = checked((uint)((position - 1) / 4) * (uint)(mapBytes * 8));
                    foreach (int owned in MapBits(Slice(page, 4, mapBytes), start, cancellation)) yield return owned;
                }
            } else { throw new InvalidDataException("Native Access usage map type is unsupported."); }
        }
        private IEnumerable<int> MapBits(OfficeByteView bits, uint start, CancellationToken cancellation) {
            for (int offset = 0; offset < bits.Length; offset++) {
                cancellation.ThrowIfCancellationRequested(); byte mask = bits[offset];
                for (int bit = 0; bit < 8; bit++) if ((mask & (1 << bit)) != 0) {
                    long page = (long)start + (long)offset * 8 + bit;
                    if (page <= 0 || page >= PageCount) throw new InvalidDataException("Native Access usage map refers outside the source snapshot.");
                    yield return (int)page;
                }
            }
        }
        public void Dispose() { _bytes = Array.Empty<byte>(); _definitions.Clear(); }
    }
}
