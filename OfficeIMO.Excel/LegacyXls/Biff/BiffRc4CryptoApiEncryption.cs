using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffRc4CryptoApiEncryption {
        private const ushort FilePassRc4EncryptionType = 0x0001;
        private const ushort CryptoApiMajorVersionMin = 0x0002;
        private const ushort CryptoApiMajorVersionMax = 0x0004;
        private const ushort CryptoApiMinorVersion = 0x0002;
        private const uint Rc4AlgorithmId = 0x00006801;
        private const uint Sha1AlgorithmId = 0x00008004;
        private const int SaltSize = 16;
        private const int VerifierSize = 16;
        private const int Sha1HashSize = 20;

        internal static bool IsRc4CryptoApiFilePass(BiffRecord record) {
            if (record.Type != (ushort)BiffRecordType.FilePass || record.Payload.Length < 6) {
                return false;
            }

            ushort encryptionType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            ushort majorVersion = BiffRecordReader.ReadUInt16(record.Payload, 2);
            ushort minorVersion = BiffRecordReader.ReadUInt16(record.Payload, 4);
            return encryptionType == FilePassRc4EncryptionType
                && majorVersion >= CryptoApiMajorVersionMin
                && majorVersion <= CryptoApiMajorVersionMax
                && minorVersion == CryptoApiMinorVersion;
        }

        internal static bool TryDecrypt(
            byte[] workbookStream,
            BiffRecord filePassRecord,
            string password,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out byte[] decryptedWorkbookStream) {
            decryptedWorkbookStream = workbookStream;
            if (!TryReadDescriptor(filePassRecord.Payload, out Rc4CryptoApiDescriptor? descriptor, out string? error)) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-FILEPASS-INVALID",
                    "The workbook contains an RC4 CryptoAPI FilePass record, but its encryption descriptor could not be parsed. " + error,
                    recordOffset: filePassRecord.Offset,
                    recordType: filePassRecord.Type,
                    detailCode: "Encryption:FilePass:Rc4CryptoApi"));
                return false;
            }

            if (!VerifyPassword(descriptor!, password)) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-FILEPASS-PASSWORD-INVALID",
                    "The supplied legacy XLS password did not match the workbook FilePass verifier.",
                    recordOffset: filePassRecord.Offset,
                    recordType: filePassRecord.Type,
                    detailCode: "Encryption:FilePass:Rc4CryptoApi"));
                return false;
            }

            decryptedWorkbookStream = DecryptWorkbookStream(workbookStream, descriptor!, password);
            return true;
        }

        private static bool TryReadDescriptor(byte[] payload, out Rc4CryptoApiDescriptor? descriptor, out string? error) {
            descriptor = null;
            error = null;
            if (payload.Length < 14) {
                error = "The FilePass payload is shorter than the RC4 CryptoAPI header.";
                return false;
            }

            uint headerSize = BiffRecordReader.ReadUInt32(payload, 10);
            if (headerSize < 32 || headerSize > int.MaxValue) {
                error = "The FilePass encryption header size is invalid.";
                return false;
            }

            int headerOffset = 14;
            int verifierOffset = checked(headerOffset + (int)headerSize);
            if (verifierOffset + 40 > payload.Length) {
                error = "The FilePass payload ends before the encryption verifier.";
                return false;
            }

            uint algorithmId = BiffRecordReader.ReadUInt32(payload, headerOffset + 8);
            uint hashAlgorithmId = BiffRecordReader.ReadUInt32(payload, headerOffset + 12);
            uint keySizeBits = BiffRecordReader.ReadUInt32(payload, headerOffset + 16);
            if (algorithmId != Rc4AlgorithmId || hashAlgorithmId != Sha1AlgorithmId) {
                error = "Only RC4 with SHA-1 CryptoAPI FilePass encryption is supported.";
                return false;
            }

            if (keySizeBits == 0) {
                keySizeBits = 40;
            }

            if (keySizeBits < 40 || keySizeBits > 128 || (keySizeBits % 8) != 0) {
                error = "The RC4 CryptoAPI key size is outside the supported range.";
                return false;
            }

            uint saltSize = BiffRecordReader.ReadUInt32(payload, verifierOffset);
            if (saltSize != SaltSize || verifierOffset + 40 + Sha1HashSize > payload.Length) {
                error = "The encryption verifier salt or hash size is invalid.";
                return false;
            }

            byte[] salt = CopyBytes(payload, verifierOffset + 4, SaltSize);
            byte[] encryptedVerifier = CopyBytes(payload, verifierOffset + 20, VerifierSize);
            uint verifierHashSize = BiffRecordReader.ReadUInt32(payload, verifierOffset + 36);
            if (verifierHashSize != Sha1HashSize || verifierOffset + 40 + verifierHashSize > payload.Length) {
                error = "The encryption verifier hash size is invalid.";
                return false;
            }

            byte[] encryptedVerifierHash = CopyBytes(payload, verifierOffset + 40, checked((int)verifierHashSize));
            descriptor = new Rc4CryptoApiDescriptor(salt, encryptedVerifier, encryptedVerifierHash, checked((int)keySizeBits));
            return true;
        }

        private static bool VerifyPassword(Rc4CryptoApiDescriptor descriptor, string password) {
            byte[] key = DeriveKey(descriptor.Salt, password, blockNumber: 0, descriptor.KeySizeBits);
            byte[] verifierAndHash = new byte[VerifierSize + descriptor.EncryptedVerifierHash.Length];
            Buffer.BlockCopy(descriptor.EncryptedVerifier, 0, verifierAndHash, 0, VerifierSize);
            Buffer.BlockCopy(descriptor.EncryptedVerifierHash, 0, verifierAndHash, VerifierSize, descriptor.EncryptedVerifierHash.Length);
            BiffRc4Transform.Xor(key, verifierAndHash, 0, verifierAndHash.Length);

            byte[] verifier = CopyBytes(verifierAndHash, 0, VerifierSize);
            byte[] verifierHash = CopyBytes(verifierAndHash, VerifierSize, descriptor.EncryptedVerifierHash.Length);
            byte[] expectedHash = ComputeSha1(verifier);
            return verifierHash.AsSpan(0, expectedHash.Length).SequenceEqual(expectedHash);
        }

        private static byte[] DecryptWorkbookStream(byte[] workbookStream, Rc4CryptoApiDescriptor descriptor, string password) {
            return BiffRc4WorkbookStreamCipher.Transform(
                workbookStream,
                block => DeriveKey(descriptor.Salt, password, block, descriptor.KeySizeBits));
        }

        private static byte[] DeriveKey(byte[] salt, string password, uint blockNumber, int keySizeBits) {
            byte[] passwordBytes = Encoding.Unicode.GetBytes(password);
            byte[] saltedPassword = new byte[salt.Length + passwordBytes.Length];
            Buffer.BlockCopy(salt, 0, saltedPassword, 0, salt.Length);
            Buffer.BlockCopy(passwordBytes, 0, saltedPassword, salt.Length, passwordBytes.Length);
            byte[] passwordHash = ComputeSha1(saltedPassword);

            byte[] blockBytes = BitConverter.GetBytes(blockNumber);
            byte[] hashInput = new byte[passwordHash.Length + blockBytes.Length];
            Buffer.BlockCopy(passwordHash, 0, hashInput, 0, passwordHash.Length);
            Buffer.BlockCopy(blockBytes, 0, hashInput, passwordHash.Length, blockBytes.Length);
            byte[] finalHash = ComputeSha1(hashInput);

            byte[] key = new byte[16];
            if (keySizeBits == 40) {
                Buffer.BlockCopy(finalHash, 0, key, 0, 5);
            } else {
                Buffer.BlockCopy(finalHash, 0, key, 0, checked(keySizeBits / 8));
            }

            return key;
        }

        private static byte[] CopyBytes(byte[] source, int offset, int length) {
            byte[] result = new byte[length];
            Buffer.BlockCopy(source, offset, result, 0, length);
            return result;
        }

        private static byte[] ComputeSha1(byte[] bytes) {
            using SHA1 sha1 = SHA1.Create();
            return sha1.ComputeHash(bytes);
        }

        private sealed class Rc4CryptoApiDescriptor {
            internal Rc4CryptoApiDescriptor(byte[] salt, byte[] encryptedVerifier, byte[] encryptedVerifierHash, int keySizeBits) {
                Salt = salt;
                EncryptedVerifier = encryptedVerifier;
                EncryptedVerifierHash = encryptedVerifierHash;
                KeySizeBits = keySizeBits;
            }

            internal byte[] Salt { get; }

            internal byte[] EncryptedVerifier { get; }

            internal byte[] EncryptedVerifierHash { get; }

            internal int KeySizeBits { get; }
        }

    }
}
