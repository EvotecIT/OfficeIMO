using OfficeIMO.Excel.LegacyXls.Diagnostics;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffXorObfuscation {
        private const ushort XorEncryptionType = 0x0000;
        private const int MaxPasswordLength = 15;

        private static readonly byte[] PadArray = {
            0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80,
            0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00
        };

        private static readonly ushort[] InitialCode = {
            0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C,
            0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139,
            0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3
        };

        private static readonly ushort[] XorMatrix = {
            0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09,
            0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF,
            0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0,
            0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40,
            0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5,
            0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A,
            0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9,
            0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0,
            0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC,
            0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10,
            0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
            0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
            0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
            0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC,
            0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4
        };

        internal static bool IsXorFilePass(BiffRecord record) {
            return record.Type == (ushort)BiffRecordType.FilePass
                && record.Payload.Length >= 2
                && BiffRecordReader.ReadUInt16(record.Payload, 0) == XorEncryptionType;
        }

        internal static bool TryDecrypt(
            byte[] workbookStream,
            BiffRecord filePassRecord,
            string password,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out byte[] decryptedWorkbookStream) {
            decryptedWorkbookStream = workbookStream;
            if (!TryReadDescriptor(filePassRecord.Payload, out XorDescriptor descriptor, out string? error)) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-FILEPASS-INVALID",
                    "The workbook contains an XOR FilePass record, but its obfuscation descriptor could not be parsed. " + error,
                    recordOffset: filePassRecord.Offset,
                    recordType: filePassRecord.Type,
                    detailCode: "Encryption:FilePass:XorObfuscation"));
                return false;
            }

            byte[] passwordBytes = GetSingleBytePassword(password);
            if (passwordBytes.Length == 0 || passwordBytes.Length > MaxPasswordLength) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-FILEPASS-PASSWORD-INVALID",
                    "The supplied legacy XLS password cannot be used with XOR obfuscation because Excel XOR passwords must contain 1 through 15 single-byte characters.",
                    recordOffset: filePassRecord.Offset,
                    recordType: filePassRecord.Type,
                    detailCode: "Encryption:FilePass:XorObfuscation"));
                return false;
            }

            if (CreateXorKey(passwordBytes) != descriptor.Key
                || CreatePasswordVerifier(passwordBytes) != descriptor.VerificationBytes) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-FILEPASS-PASSWORD-INVALID",
                    "The supplied legacy XLS password did not match the workbook XOR FilePass verifier.",
                    recordOffset: filePassRecord.Offset,
                    recordType: filePassRecord.Type,
                    detailCode: "Encryption:FilePass:XorObfuscation"));
                return false;
            }

            decryptedWorkbookStream = TransformWorkbookStream(workbookStream, passwordBytes, decrypt: true);
            return true;
        }

        internal static ushort CreatePasswordVerifier(string password) {
            return CreatePasswordVerifier(GetSingleBytePassword(password));
        }

        internal static ushort CreateXorKey(string password) {
            return CreateXorKey(GetSingleBytePassword(password));
        }

        internal static byte[] ObfuscateWorkbookStream(byte[] workbookStream, string password) {
            return TransformWorkbookStream(workbookStream, GetSingleBytePassword(password), decrypt: false);
        }

        private static bool TryReadDescriptor(byte[] payload, out XorDescriptor descriptor, out string? error) {
            descriptor = default;
            error = null;
            if (payload.Length < 6) {
                error = "The FilePass payload is shorter than the XOR obfuscation structure.";
                return false;
            }

            descriptor = new XorDescriptor(
                BiffRecordReader.ReadUInt16(payload, 2),
                BiffRecordReader.ReadUInt16(payload, 4));
            return true;
        }

        private static byte[] TransformWorkbookStream(byte[] workbookStream, byte[] passwordBytes, bool decrypt) {
            byte[] transformed = new byte[workbookStream.Length];
            Buffer.BlockCopy(workbookStream, 0, transformed, 0, workbookStream.Length);
            byte[] xorArray = CreateXorArray(passwordBytes);

            int offset = 0;
            while (offset + 4 <= transformed.Length) {
                ushort recordType = BiffRecordReader.ReadUInt16(transformed, offset);
                ushort length = BiffRecordReader.ReadUInt16(transformed, offset + 2);
                int payloadOffset = offset + 4;
                int recordEnd = payloadOffset + length;
                if (recordEnd > transformed.Length) {
                    break;
                }

                TransformRecordPayload(transformed, xorArray, recordType, payloadOffset, length, decrypt);
                offset = recordEnd;
            }

            return transformed;
        }

        private static void TransformRecordPayload(byte[] bytes, byte[] xorArray, ushort recordType, int payloadOffset, int payloadLength, bool decrypt) {
            if (payloadLength == 0 || IsUnobfuscatedRecord(recordType)) {
                return;
            }

            int transformOffset = payloadOffset;
            int transformLength = payloadLength;
            int indexAdjustment = 0;
            if (recordType == (ushort)BiffRecordType.BoundSheet8) {
                if (payloadLength <= 4) {
                    return;
                }

                transformOffset += 4;
                transformLength -= 4;
                indexAdjustment = 4;
            }

            int xorArrayIndex = (transformOffset + transformLength + indexAdjustment) % xorArray.Length;
            for (int i = 0; i < transformLength; i++) {
                int position = transformOffset + i;
                bytes[position] = decrypt
                    ? RotateRight((byte)(bytes[position] ^ xorArray[xorArrayIndex]), 5)
                    : (byte)(RotateLeft(bytes[position], 5) ^ xorArray[xorArrayIndex]);
                xorArrayIndex = (xorArrayIndex + 1) % xorArray.Length;
            }
        }

        private static bool IsUnobfuscatedRecord(ushort recordType) {
            return recordType == (ushort)BiffRecordType.Bof
                || recordType == (ushort)BiffRecordType.FilePass
                || recordType == (ushort)BiffRecordType.UsrExcl
                || recordType == (ushort)BiffRecordType.FileLock
                || recordType == (ushort)BiffRecordType.InterfaceHdr
                || recordType == (ushort)BiffRecordType.RrdInfo
                || recordType == (ushort)BiffRecordType.RrdHead;
        }

        private static byte[] GetSingleBytePassword(string password) {
            int length = Math.Min(password.Length, MaxPasswordLength);
            byte[] bytes = new byte[length];
            for (int i = 0; i < length; i++) {
                char ch = password[i];
                bytes[i] = ch <= 0xFF ? (byte)ch : (byte)'?';
            }

            return bytes;
        }

        private static ushort CreatePasswordVerifier(byte[] passwordBytes) {
            ushort verifier = 0;
            for (int i = passwordBytes.Length; i >= 0; i--) {
                byte passwordByte = i == 0 ? checked((byte)passwordBytes.Length) : passwordBytes[i - 1];
                ushort intermediate1 = (verifier & 0x4000) == 0 ? (ushort)0 : (ushort)1;
                ushort intermediate2 = (ushort)((verifier << 1) & 0x7FFF);
                verifier = (ushort)((intermediate1 | intermediate2) ^ passwordByte);
            }

            return (ushort)(verifier ^ 0xCE4B);
        }

        private static ushort CreateXorKey(byte[] passwordBytes) {
            ushort xorKey = InitialCode[passwordBytes.Length - 1];
            int currentElement = 0x68;
            for (int i = passwordBytes.Length - 1; i >= 0; i--) {
                int ch = passwordBytes[i];
                for (int bit = 0; bit < 7; bit++) {
                    if ((ch & 0x40) != 0) {
                        xorKey = (ushort)(xorKey ^ XorMatrix[currentElement]);
                    }

                    ch = (ch << 1) & 0xFF;
                    currentElement--;
                }
            }

            return xorKey;
        }

        private static byte[] CreateXorArray(byte[] passwordBytes) {
            ushort xorKey = CreateXorKey(passwordBytes);
            var obfuscationArray = new byte[16];
            int index = passwordBytes.Length;

            if ((index % 2) == 1) {
                byte temp = (byte)((xorKey & 0xFF00) >> 8);
                obfuscationArray[index] = XorRotateRight(PadArray[0], temp);
                index--;

                temp = (byte)(xorKey & 0x00FF);
                obfuscationArray[index] = XorRotateRight(passwordBytes[passwordBytes.Length - 1], temp);
            }

            while (index > 0) {
                index--;
                byte temp = (byte)((xorKey & 0xFF00) >> 8);
                obfuscationArray[index] = XorRotateRight(passwordBytes[index], temp);

                index--;
                temp = (byte)(xorKey & 0x00FF);
                obfuscationArray[index] = XorRotateRight(passwordBytes[index], temp);
            }

            index = 15;
            int padIndex = 15 - passwordBytes.Length;
            while (padIndex > 0) {
                byte temp = (byte)((xorKey & 0xFF00) >> 8);
                obfuscationArray[index] = XorRotateRight(PadArray[padIndex], temp);
                index--;
                padIndex--;

                temp = (byte)(xorKey & 0x00FF);
                obfuscationArray[index] = XorRotateRight(PadArray[padIndex], temp);
                index--;
                padIndex--;
            }

            return obfuscationArray;
        }

        private static byte XorRotateRight(byte left, byte right) {
            return RotateRight((byte)(left ^ right), 1);
        }

        private static byte RotateRight(byte value, int count) {
            return (byte)(((value >> count) | (value << (8 - count))) & 0xFF);
        }

        private static byte RotateLeft(byte value, int count) {
            return (byte)(((value << count) | (value >> (8 - count))) & 0xFF);
        }

        private readonly struct XorDescriptor {
            internal XorDescriptor(ushort key, ushort verificationBytes) {
                Key = key;
                VerificationBytes = verificationBytes;
            }

            internal ushort Key { get; }

            internal ushort VerificationBytes { get; }
        }
    }
}
