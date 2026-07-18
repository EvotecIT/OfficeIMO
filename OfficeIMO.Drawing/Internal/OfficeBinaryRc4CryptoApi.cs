using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Implements the shared password verifier and record transform used by Office binary
    /// document RC4 CryptoAPI encryption.
    /// </summary>
    internal sealed class OfficeBinaryRc4CryptoApiSession {
        private const uint CryptoApiFlag = 0x00000004;
        private const uint DocumentPropertiesFlag = 0x00000008;
        private const uint Rc4AlgorithmId = 0x00006801;
        private const uint Sha1AlgorithmId = 0x00008004;
        private const uint Rc4ProviderType = 0x00000001;
        private const int SaltSize = 16;
        private const int VerifierSize = 16;
        private const int Sha1HashSize = 20;
        private const string EnhancedProviderName = "Microsoft Enhanced Cryptographic Provider v1.0";

        private readonly byte[] _passwordHash;

        private OfficeBinaryRc4CryptoApiSession(uint flags, int keySizeBits,
            byte[] passwordHash) {
            Flags = flags;
            KeySizeBits = keySizeBits;
            _passwordHash = passwordHash;
        }

        internal uint Flags { get; }

        internal int KeySizeBits { get; }

        internal bool EncryptsDocumentProperties =>
            (Flags & DocumentPropertiesFlag) == 0;

        internal static OfficeBinaryRc4CryptoApiSession Create(string password,
            int keySizeBits, bool encryptDocumentProperties,
            out byte[] encryptionHeader) {
            ValidatePassword(password);
            ValidateKeySize(keySizeBits);

            var salt = new byte[SaltSize];
            var verifier = new byte[VerifierSize];
            using (RandomNumberGenerator random = RandomNumberGenerator.Create()) {
                random.GetBytes(salt);
                random.GetBytes(verifier);
            }

            uint flags = CryptoApiFlag
                | (encryptDocumentProperties ? 0U : DocumentPropertiesFlag);
            byte[] passwordHash = ComputePasswordHash(salt, password);
            var session = new OfficeBinaryRc4CryptoApiSession(flags,
                keySizeBits, passwordHash);
            byte[] verifierAndHash = new byte[VerifierSize + Sha1HashSize];
            Buffer.BlockCopy(verifier, 0, verifierAndHash, 0, VerifierSize);
            Buffer.BlockCopy(ComputeSha1(verifier), 0, verifierAndHash,
                VerifierSize, Sha1HashSize);
            session.TransformInPlace(verifierAndHash, 0,
                verifierAndHash.Length, blockNumber: 0);
            encryptionHeader = BuildEncryptionHeader(flags, keySizeBits,
                salt, verifierAndHash);
            return session;
        }

        internal static OfficeBinaryRc4CryptoApiSession Open(byte[] encryptionHeader,
            string password) {
            if (encryptionHeader == null) throw new ArgumentNullException(nameof(encryptionHeader));
            ValidatePassword(password);
            ParsedHeader parsed = ParseEncryptionHeader(encryptionHeader);
            byte[] passwordHash = ComputePasswordHash(parsed.Salt, password);
            var session = new OfficeBinaryRc4CryptoApiSession(parsed.Flags,
                parsed.KeySizeBits, passwordHash);

            byte[] verifierAndHash = new byte[VerifierSize + Sha1HashSize];
            Buffer.BlockCopy(parsed.EncryptedVerifier, 0, verifierAndHash,
                0, VerifierSize);
            Buffer.BlockCopy(parsed.EncryptedVerifierHash, 0,
                verifierAndHash, VerifierSize, Sha1HashSize);
            session.TransformInPlace(verifierAndHash, 0,
                verifierAndHash.Length, blockNumber: 0);
            byte[] expectedHash = ComputeSha1(CopyBytes(verifierAndHash,
                0, VerifierSize));
            if (!FixedTimeEquals(expectedHash, verifierAndHash,
                    VerifierSize, Sha1HashSize)) {
                throw new CryptographicException("The password is incorrect.");
            }
            return session;
        }

        internal void TransformInPlace(byte[] bytes, int offset, int length,
            uint blockNumber,
            CancellationToken cancellationToken = default) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (offset < 0 || length < 0 || offset > bytes.Length - length) {
                throw new ArgumentOutOfRangeException(nameof(offset));
            }
            byte[] key = DeriveKey(blockNumber);
            var transform = new OfficeRc4Transform(key);
            for (int index = 0; index < length; index++) {
                if ((index & 0xFFFF) == 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                }
                int position = offset + index;
                bytes[position] = unchecked((byte)(bytes[position]
                    ^ transform.NextByte()));
            }
            cancellationToken.ThrowIfCancellationRequested();
        }

        private byte[] DeriveKey(uint blockNumber) {
            byte[] blockBytes = new byte[4];
            WriteUInt32(blockBytes, 0, blockNumber);
            byte[] hashInput = new byte[_passwordHash.Length + blockBytes.Length];
            Buffer.BlockCopy(_passwordHash, 0, hashInput, 0,
                _passwordHash.Length);
            Buffer.BlockCopy(blockBytes, 0, hashInput, _passwordHash.Length,
                blockBytes.Length);
            byte[] finalHash = ComputeSha1(hashInput);
            int keyByteCount = KeySizeBits / 8;
            int actualKeyByteCount = KeySizeBits == 40 ? 16 : keyByteCount;
            var key = new byte[actualKeyByteCount];
            Buffer.BlockCopy(finalHash, 0, key, 0, keyByteCount);
            return key;
        }

        private static ParsedHeader ParseEncryptionHeader(byte[] bytes) {
            if (bytes.Length < 12 + 32 + 4 + SaltSize + VerifierSize
                    + 4 + Sha1HashSize) {
                throw new InvalidDataException(
                    "The RC4 CryptoAPI encryption header is truncated.");
            }
            ushort majorVersion = ReadUInt16(bytes, 0);
            ushort minorVersion = ReadUInt16(bytes, 2);
            if (majorVersion < 2 || majorVersion > 4 || minorVersion != 2) {
                throw new NotSupportedException(
                    $"RC4 CryptoAPI version {majorVersion}.{minorVersion} is not supported.");
            }
            uint flags = ReadUInt32(bytes, 4);
            uint headerSize = ReadUInt32(bytes, 8);
            if (headerSize < 32 || headerSize > int.MaxValue
                || 12L + headerSize > bytes.Length) {
                throw new InvalidDataException(
                    "The RC4 CryptoAPI EncryptionHeader size is invalid.");
            }
            int headerOffset = 12;
            uint headerFlags = ReadUInt32(bytes, headerOffset);
            if (flags != headerFlags || (flags & CryptoApiFlag) == 0
                || (flags & 0x00000030) != 0) {
                throw new NotSupportedException(
                    "Only Office binary RC4 CryptoAPI encryption is supported.");
            }
            if (ReadUInt32(bytes, headerOffset + 4) != 0
                || ReadUInt32(bytes, headerOffset + 8) != Rc4AlgorithmId
                || ReadUInt32(bytes, headerOffset + 12) != Sha1AlgorithmId
                || ReadUInt32(bytes, headerOffset + 20) != Rc4ProviderType
                || ReadUInt32(bytes, headerOffset + 28) != 0) {
                throw new NotSupportedException(
                    "The encryption header does not describe RC4 with SHA-1 CryptoAPI encryption.");
            }
            uint keySize = ReadUInt32(bytes, headerOffset + 16);
            int keySizeBits = keySize == 0 ? 40 : checked((int)keySize);
            ValidateKeySize(keySizeBits);

            int verifierOffset = checked(12 + (int)headerSize);
            if (verifierOffset > bytes.Length - (4 + SaltSize + VerifierSize
                    + 4 + Sha1HashSize)) {
                throw new InvalidDataException(
                    "The RC4 CryptoAPI encryption verifier is truncated.");
            }
            if (ReadUInt32(bytes, verifierOffset) != SaltSize
                || ReadUInt32(bytes, verifierOffset + 36) != Sha1HashSize) {
                throw new InvalidDataException(
                    "The RC4 CryptoAPI encryption verifier has an invalid size.");
            }
            return new ParsedHeader(flags, keySizeBits,
                CopyBytes(bytes, verifierOffset + 4, SaltSize),
                CopyBytes(bytes, verifierOffset + 20, VerifierSize),
                CopyBytes(bytes, verifierOffset + 40, Sha1HashSize));
        }

        private static byte[] BuildEncryptionHeader(uint flags, int keySizeBits,
            byte[] salt, byte[] verifierAndHash) {
            byte[] providerName = Encoding.Unicode.GetBytes(
                EnhancedProviderName + "\0");
            int headerSize = checked(32 + providerName.Length);
            var bytes = new byte[checked(12 + headerSize + 4 + SaltSize
                + VerifierSize + 4 + Sha1HashSize)];
            WriteUInt16(bytes, 0, 4);
            WriteUInt16(bytes, 2, 2);
            WriteUInt32(bytes, 4, flags);
            WriteUInt32(bytes, 8, unchecked((uint)headerSize));
            int headerOffset = 12;
            WriteUInt32(bytes, headerOffset, flags);
            WriteUInt32(bytes, headerOffset + 8, Rc4AlgorithmId);
            WriteUInt32(bytes, headerOffset + 12, Sha1AlgorithmId);
            WriteUInt32(bytes, headerOffset + 16,
                unchecked((uint)keySizeBits));
            WriteUInt32(bytes, headerOffset + 20, Rc4ProviderType);
            Buffer.BlockCopy(providerName, 0, bytes, headerOffset + 32,
                providerName.Length);
            int verifierOffset = checked(12 + headerSize);
            WriteUInt32(bytes, verifierOffset, SaltSize);
            Buffer.BlockCopy(salt, 0, bytes, verifierOffset + 4, SaltSize);
            Buffer.BlockCopy(verifierAndHash, 0, bytes,
                verifierOffset + 20, VerifierSize);
            WriteUInt32(bytes, verifierOffset + 36, Sha1HashSize);
            Buffer.BlockCopy(verifierAndHash, VerifierSize, bytes,
                verifierOffset + 40, Sha1HashSize);
            return bytes;
        }

        private static byte[] ComputePasswordHash(byte[] salt, string password) {
            byte[] passwordBytes = Encoding.Unicode.GetBytes(
                password.Length > 255 ? password.Substring(0, 255) : password);
            byte[] saltedPassword = new byte[salt.Length + passwordBytes.Length];
            Buffer.BlockCopy(salt, 0, saltedPassword, 0, salt.Length);
            Buffer.BlockCopy(passwordBytes, 0, saltedPassword, salt.Length,
                passwordBytes.Length);
            return ComputeSha1(saltedPassword);
        }

        private static void ValidatePassword(string password) {
            if (password == null) throw new ArgumentNullException(nameof(password));
        }

        private static void ValidateKeySize(int keySizeBits) {
            if (keySizeBits < 40 || keySizeBits > 128
                || keySizeBits % 8 != 0) {
                throw new ArgumentOutOfRangeException(nameof(keySizeBits),
                    "RC4 CryptoAPI keys must contain 40 through 128 bits in 8-bit increments.");
            }
        }

        private static bool FixedTimeEquals(byte[] expected, byte[] actual,
            int actualOffset, int length) {
            if (expected.Length != length || actualOffset < 0
                || actualOffset > actual.Length - length) return false;
            int difference = 0;
            for (int index = 0; index < length; index++) {
                difference |= expected[index] ^ actual[actualOffset + index];
            }
            return difference == 0;
        }

        private static byte[] ComputeSha1(byte[] bytes) {
            using SHA1 sha1 = SHA1.Create();
            return sha1.ComputeHash(bytes);
        }

        private static byte[] CopyBytes(byte[] source, int offset, int length) {
            var result = new byte[length];
            Buffer.BlockCopy(source, offset, result, 0, length);
            return result;
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) =>
            unchecked((ushort)(bytes[offset] | bytes[offset + 1] << 8));

        private static uint ReadUInt32(byte[] bytes, int offset) =>
            unchecked((uint)(bytes[offset] | bytes[offset + 1] << 8
                | bytes[offset + 2] << 16 | bytes[offset + 3] << 24));

        private static void WriteUInt16(byte[] bytes, int offset,
            ushort value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
        }

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
            bytes[offset + 2] = unchecked((byte)(value >> 16));
            bytes[offset + 3] = unchecked((byte)(value >> 24));
        }

        private sealed class ParsedHeader {
            internal ParsedHeader(uint flags, int keySizeBits, byte[] salt,
                byte[] encryptedVerifier, byte[] encryptedVerifierHash) {
                Flags = flags;
                KeySizeBits = keySizeBits;
                Salt = salt;
                EncryptedVerifier = encryptedVerifier;
                EncryptedVerifierHash = encryptedVerifierHash;
            }

            internal uint Flags { get; }
            internal int KeySizeBits { get; }
            internal byte[] Salt { get; }
            internal byte[] EncryptedVerifier { get; }
            internal byte[] EncryptedVerifierHash { get; }
        }

        private sealed class OfficeRc4Transform {
            private readonly byte[] _state = new byte[256];
            private int _i;
            private int _j;

            internal OfficeRc4Transform(byte[] key) {
                for (int index = 0; index < _state.Length; index++) {
                    _state[index] = unchecked((byte)index);
                }
                int j = 0;
                for (int index = 0; index < _state.Length; index++) {
                    j = (j + _state[index] + key[index % key.Length]) & 0xFF;
                    Swap(index, j);
                }
            }

            internal byte NextByte() {
                _i = (_i + 1) & 0xFF;
                _j = (_j + _state[_i]) & 0xFF;
                Swap(_i, _j);
                return _state[(_state[_i] + _state[_j]) & 0xFF];
            }

            private void Swap(int left, int right) {
                byte value = _state[left];
                _state[left] = _state[right];
                _state[right] = value;
            }
        }
    }
}
