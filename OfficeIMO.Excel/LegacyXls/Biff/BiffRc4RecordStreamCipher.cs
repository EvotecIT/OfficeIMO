using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffRc4RecordStreamCipher {
        private const int EncryptionBlockSize = 1024;

        internal static byte[] Transform(byte[] workbookStream, Func<uint, byte[]> deriveBlockKey, bool requireCompleteRecords = false) {
            return TransformBlocks(workbookStream, block => {
                var pad = new byte[EncryptionBlockSize];
                BiffRc4Transform.Xor(deriveBlockKey(block), pad, 0, pad.Length);
                return pad;
            }, requireCompleteRecords);
        }

        internal static byte[] Transform(byte[] workbookStream, OfficeBinaryRc4CryptoApiSession session) =>
            TransformBlocks(workbookStream, block => {
                var pad = new byte[EncryptionBlockSize];
                session.TransformInPlace(pad, 0, pad.Length, block);
                return pad;
            }, false);

        private static byte[] TransformBlocks(byte[] workbookStream, Func<uint, byte[]> createPad, bool requireCompleteRecords) {
            if (requireCompleteRecords) ValidateRecordStream(workbookStream);
            byte[] transformed = new byte[workbookStream.Length];
            Buffer.BlockCopy(workbookStream, 0, transformed, 0, workbookStream.Length);

            byte[]? pad = null;
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
                        pad = createPad(checked((uint)block));
                    }

                    byte keyByte = pad![position % EncryptionBlockSize];
                    if (IsEncryptedByte(recordType, offset, position)) {
                        transformed[position] = (byte)(transformed[position] ^ keyByte);
                    }
                }

                offset = recordEnd;
            }

            return transformed;
        }

        internal static void ValidateRecordStream(byte[] bytes) {
            int offset = 0;
            while (offset <= bytes.Length - 4) {
                int length = BiffRecordReader.ReadUInt16(bytes, offset + 2);
                if (length > bytes.Length - offset - 4)
                    throw new InvalidDataException("The BIFF record stream ends inside an encrypted record payload.");
                offset += length + 4;
            }
            for (; offset < bytes.Length; offset++)
                if (bytes[offset] != 0) throw new InvalidDataException("The BIFF record stream has a truncated record header.");
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
