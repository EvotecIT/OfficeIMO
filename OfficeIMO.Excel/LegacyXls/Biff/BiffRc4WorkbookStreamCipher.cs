namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffRc4WorkbookStreamCipher {
        private const int EncryptionBlockSize = 1024;

        internal static byte[] Transform(byte[] workbookStream, Func<uint, byte[]> deriveBlockKey) {
            byte[] transformed = new byte[workbookStream.Length];
            Buffer.BlockCopy(workbookStream, 0, transformed, 0, workbookStream.Length);

            BiffRc4Transform? transform = null;
            int currentBlock = -1;
            int offset = 0;
            while (offset + 4 <= transformed.Length) {
                ushort recordType = BiffRecordReader.ReadUInt16(transformed, offset);
                ushort length = BiffRecordReader.ReadUInt16(transformed, offset + 2);
                int recordEnd = offset + 4 + length;
                if (recordEnd > transformed.Length) {
                    break;
                }

                for (int position = offset; position < recordEnd; position++) {
                    int block = position / EncryptionBlockSize;
                    if (block != currentBlock) {
                        currentBlock = block;
                        transform = new BiffRc4Transform(deriveBlockKey(checked((uint)block)));
                    }

                    byte keyByte = transform!.NextByte();
                    if (IsEncryptedByte(recordType, offset, position)) {
                        transformed[position] = (byte)(transformed[position] ^ keyByte);
                    }
                }

                offset = recordEnd;
            }

            return transformed;
        }

        private static bool IsEncryptedByte(ushort recordType, int recordOffset, int position) {
            if (position < recordOffset + 4) {
                return false;
            }

            if (recordType == (ushort)BiffRecordType.BoundSheet8) {
                return position >= recordOffset + 8;
            }

            return !IsUnencryptedRecord(recordType);
        }

        private static bool IsUnencryptedRecord(ushort recordType) {
            return recordType == (ushort)BiffRecordType.Bof
                || recordType == (ushort)BiffRecordType.FilePass
                || recordType == (ushort)BiffRecordType.UsrExcl
                || recordType == (ushort)BiffRecordType.FileLock
                || recordType == (ushort)BiffRecordType.InterfaceHdr
                || recordType == (ushort)BiffRecordType.RrdInfo
                || recordType == (ushort)BiffRecordType.RrdHead;
        }
    }
}
