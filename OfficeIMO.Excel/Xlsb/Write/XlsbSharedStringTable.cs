using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Collects plain text cells and encodes the workbook's BIFF12 shared-string part.</summary>
    internal sealed class XlsbSharedStringTable {
        private readonly Dictionary<string, int> _indexes = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<string> _values = new List<string>();
        private int _referenceCount;

        internal int GetOrAdd(string value) {
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            if (!_indexes.TryGetValue(value, out int index)) {
                index = _values.Count;
                _indexes.Add(value, index);
                _values.Add(value);
            }
            // BrtBeginSst counts are restricted to signed 31-bit values by MS-XLSB.
            _referenceCount = checked(_referenceCount + 1);
            return index;
        }

        internal ArraySegment<byte> CreatePart(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            using var output = new MemoryStream();
            using var writer = new XlsbDirectRecordWriter(output);
            writer.WriteHeader(159, payloadLength: 8); // BrtBeginSst
            writer.WriteUInt32(checked((uint)_referenceCount));
            writer.WriteUInt32(checked((uint)_values.Count));
            for (int index = 0; index < _values.Count; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                string value = _values[index];
                writer.WriteHeader(19, checked(5 + value.Length * 2)); // BrtSSTItem
                writer.WriteByte(0); // RichStr: no formatting runs or phonetic extensions.
                writer.WriteWideString(value);
            }
            writer.WriteRecord(160); // BrtEndSst
            writer.Flush();
            return new ArraySegment<byte>(output.GetBuffer(), 0, checked((int)output.Length));
        }
    }
}
