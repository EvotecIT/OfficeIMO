using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffRc4Encryption {
        private const ushort FilePassRc4EncryptionType = 0x0001;
        private const ushort Rc4MajorVersion = 0x0001;
        private const ushort Rc4MinorVersion = 0x0001;
        private const int SaltSize = 16;
        private const int VerifierSize = 16;
        private const int Md5HashSize = 16;
        private const int TruncatedHashSize = 5;

        internal static bool IsRc4FilePass(BiffRecord record) {
            if (record.Type != (ushort)BiffRecordType.FilePass || record.Payload.Length < 6) {
                return false;
            }

            ushort encryptionType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            ushort majorVersion = BiffRecordReader.ReadUInt16(record.Payload, 2);
            ushort minorVersion = BiffRecordReader.ReadUInt16(record.Payload, 4);
            return encryptionType == FilePassRc4EncryptionType
                && majorVersion == Rc4MajorVersion
                && minorVersion == Rc4MinorVersion;
        }

        internal static bool TryDecrypt(
            byte[] workbookStream,
            BiffRecord filePassRecord,
            string password,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out byte[] decryptedWorkbookStream) {
            decryptedWorkbookStream = workbookStream;
            if (!TryReadDescriptor(filePassRecord.Payload, out Rc4Descriptor? descriptor, out string? error)) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-FILEPASS-INVALID",
                    "The workbook contains an RC4 FilePass record, but its encryption descriptor could not be parsed. " + error,
                    recordOffset: filePassRecord.Offset,
                    recordType: filePassRecord.Type,
                    detailCode: "Encryption:FilePass:Rc4"));
                return false;
            }

            if (!VerifyPassword(descriptor!, password)) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-FILEPASS-PASSWORD-INVALID",
                    "The supplied legacy XLS password did not match the workbook RC4 FilePass verifier.",
                    recordOffset: filePassRecord.Offset,
                    recordType: filePassRecord.Type,
                    detailCode: "Encryption:FilePass:Rc4"));
                return false;
            }

            decryptedWorkbookStream = TransformWorkbookStream(workbookStream, descriptor!, password);
            return true;
        }

        internal static byte[] CreateFilePassPayload(string password, byte[] salt, byte[] verifier) {
            if (salt.Length != SaltSize) {
                throw new ArgumentException("Classic RC4 FilePass salt must be 16 bytes.", nameof(salt));
            }

            if (verifier.Length != VerifierSize) {
                throw new ArgumentException("Classic RC4 FilePass verifier must be 16 bytes.", nameof(verifier));
            }

            byte[] verifierAndHash = new byte[VerifierSize + Md5HashSize];
            Buffer.BlockCopy(verifier, 0, verifierAndHash, 0, VerifierSize);
            Buffer.BlockCopy(ComputeMd5(verifier), 0, verifierAndHash, VerifierSize, Md5HashSize);
            byte[] key = DeriveKey(salt, password, blockNumber: 0);
            BiffRc4Transform.Xor(key, verifierAndHash, 0, verifierAndHash.Length);

            byte[] payload = new byte[2 + 2 + 2 + SaltSize + VerifierSize + Md5HashSize];
            WriteUInt16(payload, 0, FilePassRc4EncryptionType);
            WriteUInt16(payload, 2, Rc4MajorVersion);
            WriteUInt16(payload, 4, Rc4MinorVersion);
            Buffer.BlockCopy(salt, 0, payload, 6, SaltSize);
            Buffer.BlockCopy(verifierAndHash, 0, payload, 22, VerifierSize);
            Buffer.BlockCopy(verifierAndHash, VerifierSize, payload, 38, Md5HashSize);
            return payload;
        }

        internal static byte[] EncryptWorkbookStream(byte[] workbookStream, string password, byte[] salt) {
            return BiffRc4WorkbookStreamCipher.Transform(
                workbookStream,
                block => DeriveKey(salt, password, block));
        }

        private static bool TryReadDescriptor(byte[] payload, out Rc4Descriptor? descriptor, out string? error) {
            descriptor = null;
            error = null;
            if (payload.Length < 2 + 2 + 2 + SaltSize + VerifierSize + Md5HashSize) {
                error = "The FilePass payload is shorter than the RC4 encryption header.";
                return false;
            }

            descriptor = new Rc4Descriptor(
                CopyBytes(payload, 6, SaltSize),
                CopyBytes(payload, 22, VerifierSize),
                CopyBytes(payload, 38, Md5HashSize));
            return true;
        }

        private static bool VerifyPassword(Rc4Descriptor descriptor, string password) {
            byte[] key = DeriveKey(descriptor.Salt, password, blockNumber: 0);
            byte[] verifierAndHash = new byte[VerifierSize + Md5HashSize];
            Buffer.BlockCopy(descriptor.EncryptedVerifier, 0, verifierAndHash, 0, VerifierSize);
            Buffer.BlockCopy(descriptor.EncryptedVerifierHash, 0, verifierAndHash, VerifierSize, Md5HashSize);
            BiffRc4Transform.Xor(key, verifierAndHash, 0, verifierAndHash.Length);
            byte[] verifier = CopyBytes(verifierAndHash, 0, VerifierSize);
            byte[] verifierHash = CopyBytes(verifierAndHash, VerifierSize, Md5HashSize);
            byte[] expectedHash = ComputeMd5(verifier);
            return verifierHash.AsSpan(0, expectedHash.Length).SequenceEqual(expectedHash);
        }

        private static byte[] TransformWorkbookStream(byte[] workbookStream, Rc4Descriptor descriptor, string password) {
            return BiffRc4WorkbookStreamCipher.Transform(
                workbookStream,
                block => DeriveKey(descriptor.Salt, password, block));
        }

        private static byte[] DeriveKey(byte[] salt, string password, uint blockNumber) {
            byte[] passwordHash = ComputeMd5(Encoding.Unicode.GetBytes(password));
            byte[] truncatedHash = CopyBytes(passwordHash, 0, TruncatedHashSize);
            byte[] intermediateBuffer = new byte[(TruncatedHashSize + SaltSize) * 16];
            int offset = 0;
            for (int i = 0; i < 16; i++) {
                Buffer.BlockCopy(truncatedHash, 0, intermediateBuffer, offset, TruncatedHashSize);
                offset += TruncatedHashSize;
                Buffer.BlockCopy(salt, 0, intermediateBuffer, offset, SaltSize);
                offset += SaltSize;
            }

            byte[] intermediateHash = ComputeMd5(intermediateBuffer);
            byte[] keySource = new byte[TruncatedHashSize + 4];
            Buffer.BlockCopy(intermediateHash, 0, keySource, 0, TruncatedHashSize);
            Buffer.BlockCopy(BitConverter.GetBytes(blockNumber), 0, keySource, TruncatedHashSize, 4);
            return ComputeMd5(keySource);
        }

        private static byte[] CopyBytes(byte[] source, int offset, int length) {
            byte[] result = new byte[length];
            Buffer.BlockCopy(source, offset, result, 0, length);
            return result;
        }

        private static byte[] ComputeMd5(byte[] bytes) {
            using MD5 md5 = MD5.Create();
            return md5.ComputeHash(bytes);
        }

        private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
            bytes[offset] = (byte)(value & 0xff);
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private sealed class Rc4Descriptor {
            internal Rc4Descriptor(byte[] salt, byte[] encryptedVerifier, byte[] encryptedVerifierHash) {
                Salt = salt;
                EncryptedVerifier = encryptedVerifier;
                EncryptedVerifierHash = encryptedVerifierHash;
            }

            internal byte[] Salt { get; }

            internal byte[] EncryptedVerifier { get; }

            internal byte[] EncryptedVerifierHash { get; }
        }
    }
}
