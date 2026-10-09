namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocParagraphFormattingReader {
        private const ushort SprmPHugePapx = 0x6646;
        private const ushort SprmPTableProps = 0x646B;
        private const int MaximumPrcDataLength = 0x3FA2;

        /// <summary>
        /// Resolves the ordered paragraph properties stored in the Data stream.
        /// A followed pointer replaces the remainder of its containing property array.
        /// </summary>
        internal static byte[] ResolveDataProperties(byte[] properties, int offset, int count,
            byte[] dataStream, ushort styleIndex = 0) {
            if (offset < 0 || count < 0 || offset > properties.Length - count)
                throw new InvalidDataException("The paragraph property array is truncated.");
            var result = new List<byte>();
            var visited = new HashSet<int>();
            long maximumExpandedBytes = (long)count + dataStream.Length;
            int start = offset;
            int end = offset + count;
            bool inPapx = true;
            while (offset < end) {
                // Older OfficeIMO page PAPX records included one alignment zero in
                // the even-length property count. Keep reading those stored files;
                // referenced PrcData arrays still require complete property records.
                if (inPapx && end - offset == 1 && properties[offset] == 0) break;
                if (end - offset < 2)
                    throw new InvalidDataException("The paragraph property code is truncated.");
                ushort sprm = LegacyDocFib.ReadUInt16(properties, offset);
                if (sprm == SprmPHugePapx && offset != start) {
                    // MS-DOC requires a non-leading PHugePapx to be ignored.
                    if (end - offset < 6)
                        throw new InvalidDataException("The ignored paragraph Data pointer is truncated.");
                    offset += 6;
                    continue;
                }
                if (sprm == SprmPHugePapx || sprm == SprmPTableProps) {
                    if (end - offset < 6)
                        throw new InvalidDataException("The paragraph Data pointer is truncated.");
                    if (inPapx && sprm == SprmPHugePapx && (count != 6 || styleIndex != 0))
                        throw new InvalidDataException("A PAPX PHugePapx must be its only property and use style zero.");
                    uint unsignedOffset = unchecked((uint)LegacyDocFib.ReadInt32(properties, offset + 2));
                    if (unsignedOffset > int.MaxValue || unsignedOffset > (uint)Math.Max(0, dataStream.Length - 2)
                        || dataStream.Length < 2)
                        throw new InvalidDataException("A paragraph Data pointer is outside the Data stream.");
                    int dataOffset = (int)unsignedOffset;
                    if (!visited.Add(dataOffset))
                        throw new InvalidDataException("The paragraph Data property chain contains a cycle.");
                    int length = ReadInt16(dataStream, dataOffset);
                    if (length < 10 || length > MaximumPrcDataLength || length > dataStream.Length - dataOffset - 2)
                        throw new InvalidDataException("The referenced paragraph Data record has an invalid length.");
                    properties = dataStream;
                    start = offset = dataOffset + 2;
                    end = offset + length;
                    inPapx = false;
                    continue;
                }
                int operandLength;
                if (sprm == SprmTDefTable) {
                    if (end - offset < 4)
                        throw new InvalidDataException("The table definition length is truncated.");
                    operandLength = LegacyDocFib.ReadUInt16(properties, offset + 2) + 1;
                    if (operandLength < 5 || operandLength > end - offset - 2)
                        throw new InvalidDataException("The table definition operand is truncated.");
                } else if (!TryGetSprmOperandLength(properties, offset, end, out operandLength)) {
                    throw new InvalidDataException("A paragraph property operand is truncated.");
                }
                // Overlapping adversarial records must not amplify their property bytes
                // beyond the physical paragraph and Data stream input budget.
                if (result.Count + 2L + operandLength > maximumExpandedBytes)
                    throw new InvalidDataException("The paragraph Data chain exceeds its stored property budget.");
                for (int index = offset; index < offset + 2 + operandLength; index++) result.Add(properties[index]);
                offset += 2 + operandLength;
            }
            return result.ToArray();
        }
    }
}
