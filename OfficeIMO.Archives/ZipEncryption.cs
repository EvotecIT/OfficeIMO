using System.IO.Compression;
using System.Text;
using System.Security.Cryptography;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

internal enum ZipCipher {
    None,
    ZipCrypto,
    WinZipAes
}

/// <summary>How one ZIP member is protected and compressed, resolved from its central-directory record.</summary>
internal readonly struct ZipMemberScheme {
    internal ZipMemberScheme(ZipCipher cipher, ushort method, int keyLength, bool checksCrc) {
        Cipher = cipher;
        Method = method;
        KeyLength = keyLength;
        ChecksCrc = checksCrc;
    }

    internal ZipCipher Cipher { get; }
    /// <summary>The compression method of the plaintext: 0 (stored) or 8 (Deflate).</summary>
    internal ushort Method { get; }
    /// <summary>WinZip AES key length in bytes (16, 24 or 32); zero otherwise.</summary>
    internal int KeyLength { get; }
    /// <summary>False for WinZip AES AE-2, which stores no CRC and relies on the authentication code.</summary>
    internal bool ChecksCrc { get; }
    internal bool Encrypted => Cipher != ZipCipher.None;
}

/// <summary>Traditional PKWARE encryption (APPNOTE 6.3.10, section 6.1) and WinZip AES (AE-1 and AE-2) for ZIP members.</summary>
internal static class ZipEncryption {
    private const ushort Stored = 0;
    private const ushort Deflate = 8;
    private const ushort AesMethod = 99;
    private const ushort AesExtraId = 0x9901;
    private const int ZipCryptoHeaderLength = 12;
    private const int AesVerifierLength = 2;
    private const int AesMacLength = 10;
    private const int AesIterations = 1000;
    // Encryption flag, data descriptor and UTF-8 names; Deflate also allows its two level-hint bits. Strong encryption
    // (bit 6) and central-directory encryption (bit 13) are outside these two schemes.
    private const ushort EncryptedFlags = 0x0809;

    /// <summary>Resolves the member's scheme, or returns why it cannot be read.</summary>
    internal static ArchiveInspectionDiagnostic? Inspect(OfficeArchiveSafety.ZipCentralDirectoryRecord record, out ZipMemberScheme scheme) {
        scheme = default;
        if ((record.Flags & 1) == 0) {
            if (record.CompressionMethod != Stored && record.CompressionMethod != Deflate)
                return new ArchiveInspectionDiagnostic("ZIP_COMPRESSION", "The ZIP member compression method is unsupported.");
            if (!OfficeArchiveSafety.HasSupportedZipFlags(record.Flags, record.CompressionMethod))
                return new ArchiveInspectionDiagnostic("ZIP_FLAGS", "The ZIP member uses unsupported flags.");
            scheme = new ZipMemberScheme(ZipCipher.None, record.CompressionMethod, 0, checksCrc: true);
            return null;
        }
        if ((record.Flags & 0x0040) != 0 || (record.Flags & 0x2000) != 0)
            return new ArchiveInspectionDiagnostic("ZIP_ENCRYPTION", "The ZIP member uses PKWARE strong encryption, which is unsupported.");
        ushort method = record.CompressionMethod;
        ZipCipher cipher = ZipCipher.ZipCrypto;
        int keyLength = 0;
        bool checksCrc = true;
        long overhead = ZipCryptoHeaderLength;
        if (method == AesMethod) {
            if (!TryReadAesExtra(record.CentralExtra, out int version, out int strength, out method))
                return new ArchiveInspectionDiagnostic("ZIP_AES_EXTRA", "The WinZip AES extra field is missing or malformed.");
            if (version != 1 && version != 2 || strength < 1 || strength > 3)
                return new ArchiveInspectionDiagnostic("ZIP_ENCRYPTION", "The ZIP member uses an unknown WinZip AES version or key strength.");
            cipher = ZipCipher.WinZipAes;
            keyLength = 8 * (strength + 1);
            checksCrc = version == 1;
            overhead = keyLength / 2 + AesVerifierLength + AesMacLength;
        }
        if (method != Stored && method != Deflate)
            return new ArchiveInspectionDiagnostic("ZIP_COMPRESSION", "The encrypted ZIP member compression method is unsupported.");
        if ((record.Flags & ~(EncryptedFlags | (method == Deflate ? 0x0006 : 0))) != 0)
            return new ArchiveInspectionDiagnostic("ZIP_FLAGS", "The ZIP member uses unsupported flags.");
        if (record.CompressedSize < overhead || method == Stored && record.CompressedSize - overhead != record.UncompressedSize)
            return new ArchiveInspectionDiagnostic("ZIP_ENCRYPTED_SIZE", "The encrypted ZIP member is shorter than its encryption header or its sizes disagree.");
        scheme = new ZipMemberScheme(cipher, method, keyLength, checksCrc);
        return null;
    }

    /// <summary>Resolves the member's scheme, throwing <see cref="NotSupportedException"/> or <see cref="InvalidDataException"/>.</summary>
    internal static ZipMemberScheme Describe(OfficeArchiveSafety.ZipCentralDirectoryRecord record) {
        ArchiveInspectionDiagnostic? problem = Inspect(record, out ZipMemberScheme scheme);
        if (problem == null) return scheme;
        if (problem.Code == "ZIP_AES_EXTRA" || problem.Code == "ZIP_ENCRYPTED_SIZE") throw new InvalidDataException(problem.Message);
        throw new NotSupportedException(problem.Message);
    }

    /// <summary>Checks the local header and payload range of a member that may be skipped.</summary>
    internal static void Validate(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, ZipMemberScheme scheme) =>
        OfficeArchiveSafety.ValidateEntryPayload(source, record, scheme.Encrypted);

    /// <summary>Opens the member's decoded plaintext. Encrypted members are checked before any plaintext is returned:
    /// WinZip AES members are authenticated and Deflate framing is validated over the decrypted bytes.</summary>
    internal static Stream Open(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, ZipMemberScheme scheme,
        string? password, string name, CancellationToken token) {
        if (!scheme.Encrypted) return OfficeArchiveSafety.OpenEntryPayload(source, record, token);
        if (password == null)
            throw new ArchivePasswordException(name, false, $"ZIP member '{name}' is encrypted; supply ZipPassword to read it.");
        long payload = OfficeArchiveSafety.ValidateEntryPayload(source, record, encrypted: true);
        byte[] secret = Encoding.UTF8.GetBytes(password);
        return scheme.Cipher == ZipCipher.ZipCrypto
            ? OpenZipCrypto(source, record, scheme, secret, payload, name, token)
            : OpenWinZipAes(source, record, scheme, secret, payload, name, token);
    }

    private static Stream OpenZipCrypto(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, ZipMemberScheme scheme,
        byte[] password, long payload, string name, CancellationToken token) {
        var keys = ZipCryptoKeys.Create(password);
        var header = new byte[ZipCryptoHeaderLength];
        using (var window = new SeekableReadWindowStream(source, payload, ZipCryptoHeaderLength, cancellationToken: token))
            ReadExactly(window, header);
        keys.Decrypt(header, 0, header.Length);
        // The last header byte repeats the CRC's high byte; with a data descriptor, writers use the high byte of the
        // modification time instead, so either is accepted there. One byte is all the format offers.
        byte check = header[ZipCryptoHeaderLength - 1];
        bool matches = check == (byte)(record.Crc32 >> 24) || (record.Flags & 8) != 0 && check == (byte)(record.DosTime >> 8);
        if (!matches) throw new ArchivePasswordException(name, true, $"The password does not open ZIP member '{name}'.");
        long start = payload + ZipCryptoHeaderLength, length = record.CompressedSize - ZipCryptoHeaderLength;
        Stream Plain() => new ZipCryptoStream(new SeekableReadWindowStream(source, start, length, cancellationToken: token), keys);
        if (scheme.Method == Stored) return Plain();
        using (Stream validation = Plain()) {
            if (!OfficeDeflateStreamValidator.TryValidateExact(validation, record.UncompressedSize, token))
                throw new ArchivePasswordException(name, true, $"The password does not open ZIP member '{name}', or the member is damaged.");
        }
        return new DeflateStream(Plain(), CompressionMode.Decompress);
    }

    private static Stream OpenWinZipAes(Stream source, OfficeArchiveSafety.ZipCentralDirectoryRecord record, ZipMemberScheme scheme,
        byte[] password, long payload, string name, CancellationToken token) {
        int saltLength = scheme.KeyLength / 2;
        var prefix = new byte[saltLength + AesVerifierLength];
        var mac = new byte[AesMacLength];
        long macOffset = payload + record.CompressedSize - AesMacLength;
        using (var window = new SeekableReadWindowStream(source, payload, prefix.Length, cancellationToken: token)) ReadExactly(window, prefix);
        using (var window = new SeekableReadWindowStream(source, macOffset, AesMacLength, cancellationToken: token)) ReadExactly(window, mac);
        var salt = new byte[saltLength];
        Buffer.BlockCopy(prefix, 0, salt, 0, saltLength);
        byte[] derived = DeriveKeys(password, salt, 2 * scheme.KeyLength + AesVerifierLength);
        if (derived[2 * scheme.KeyLength] != prefix[saltLength] || derived[2 * scheme.KeyLength + 1] != prefix[saltLength + 1])
            throw new ArchivePasswordException(name, true, $"The password does not open ZIP member '{name}'.");
        var key = new byte[scheme.KeyLength];
        var macKey = new byte[scheme.KeyLength];
        Buffer.BlockCopy(derived, 0, key, 0, key.Length);
        Buffer.BlockCopy(derived, key.Length, macKey, 0, macKey.Length);
        long start = payload + prefix.Length, length = record.CompressedSize - prefix.Length - AesMacLength;

        // First pass: authenticate the ciphertext and, for Deflate, validate the framing of the plaintext it decrypts to.
        bool framed = true;
        using (var hmac = new HMACSHA1(macKey)) {
            using (var authenticated = new HmacReadStream(new SeekableReadWindowStream(source, start, length, cancellationToken: token), hmac)) {
                if (scheme.Method == Deflate) {
                    using var plain = new WinZipAesCtrStream(authenticated, key, ownsInner: false);
                    framed = OfficeDeflateStreamValidator.TryValidateExact(plain, record.UncompressedSize, token);
                }
                authenticated.Drain(token);
            }
            hmac.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            byte[] expected = hmac.Hash!;
            int difference = 0;
            for (int index = 0; index < AesMacLength; index++) difference |= expected[index] ^ mac[index];
            if (difference != 0) throw new ArchiveCorruptionException("ZIP", macOffset, $"the WinZip AES authentication code of '{name}'");
        }
        if (!framed) throw new InvalidDataException("ZIP member has invalid or incomplete Deflate framing.");
        Stream decrypted = new WinZipAesCtrStream(new SeekableReadWindowStream(source, start, length, cancellationToken: token), key, ownsInner: true);
        return scheme.Method == Stored ? decrypted : new DeflateStream(decrypted, CompressionMode.Decompress);
    }

    private static byte[] DeriveKeys(byte[] password, byte[] salt, int length) {
#if NET8_0_OR_GREATER
        return Rfc2898DeriveBytes.Pbkdf2(password, salt, AesIterations, HashAlgorithmName.SHA1, length);
#else
        using var kdf = new Rfc2898DeriveBytes(password, salt, AesIterations);
        return kdf.GetBytes(length);
#endif
    }

    private static bool TryReadAesExtra(byte[] extra, out int version, out int strength, out ushort method) {
        version = strength = 0;
        method = 0;
        for (int offset = 0; offset + 4 <= extra.Length;) {
            ushort id = (ushort)(extra[offset] | extra[offset + 1] << 8);
            int size = extra[offset + 2] | extra[offset + 3] << 8;
            offset += 4;
            if (size > extra.Length - offset) return false;
            if (id == AesExtraId) {
                if (size != 7 || extra[offset + 2] != (byte)'A' || extra[offset + 3] != (byte)'E') return false;
                version = extra[offset] | extra[offset + 1] << 8;
                strength = extra[offset + 4];
                method = (ushort)(extra[offset + 5] | extra[offset + 6] << 8);
                return true;
            }
            offset += size;
        }
        return false;
    }

    private static void ReadExactly(Stream stream, byte[] buffer) {
        for (int offset = 0; offset < buffer.Length;) {
            int read = stream.Read(buffer, offset, buffer.Length - offset);
            if (read == 0) throw new InvalidDataException("The encrypted ZIP member is truncated.");
            offset += read;
        }
    }

    /// <summary>The three 32-bit keys of traditional PKWARE encryption.</summary>
    private struct ZipCryptoKeys {
        private static readonly uint[] Crc = CreateTable();
        private uint _key0, _key1, _key2;

        internal static ZipCryptoKeys Create(byte[] password) {
            var keys = new ZipCryptoKeys { _key0 = 0x12345678, _key1 = 0x23456789, _key2 = 0x34567890 };
            foreach (byte value in password) keys.Update(value);
            return keys;
        }

        internal void Decrypt(byte[] buffer, int offset, int count) {
            uint key0 = _key0, key1 = _key1, key2 = _key2;
            uint[] crc = Crc;
            for (int index = offset, end = offset + count; index < end; index++) {
                uint temp = (key2 | 2) & 0xFFFF;
                byte plain = (byte)(buffer[index] ^ (byte)((temp * (temp ^ 1)) >> 8));
                buffer[index] = plain;
                key0 = crc[(key0 ^ plain) & 0xFF] ^ (key0 >> 8);
                key1 = (key1 + (key0 & 0xFF)) * 134775813 + 1;
                key2 = crc[(key2 ^ (key1 >> 24)) & 0xFF] ^ (key2 >> 8);
            }
            _key0 = key0; _key1 = key1; _key2 = key2;
        }

        private void Update(byte value) {
            _key0 = Crc[(_key0 ^ value) & 0xFF] ^ (_key0 >> 8);
            _key1 = (_key1 + (_key0 & 0xFF)) * 134775813 + 1;
            _key2 = Crc[(_key2 ^ (_key1 >> 24)) & 0xFF] ^ (_key2 >> 8);
        }

        private static uint[] CreateTable() {
            var table = new uint[256];
            for (uint index = 0; index < table.Length; index++) {
                uint value = index;
                for (int bit = 0; bit < 8; bit++) value = (value & 1) != 0 ? 0xEDB88320 ^ (value >> 1) : value >> 1;
                table[index] = value;
            }
            return table;
        }
    }

    /// <summary>A decrypting view over a seekable ciphertext window. Core's Deflate validator reads it forward and then
    /// rewinds it to the start, so positioning supports the current position and the start only.</summary>
    private abstract class DecryptingStream : Stream {
        private readonly Stream _inner;
        private readonly bool _ownsInner;
        private long _position;
        protected DecryptingStream(Stream inner, bool ownsInner) { _inner = inner; _ownsInner = ownsInner; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position {
            get => _position;
            set {
                if (value == _position) return;
                if (value != 0) throw new NotSupportedException("Decrypting streams rewind only to their start.");
                _inner.Position = 0;
                _position = 0;
                Restart();
            }
        }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = _inner.Read(buffer, offset, count);
            Transform(buffer, offset, read);
            _position += read;
            return read;
        }
        protected abstract void Transform(byte[] buffer, int offset, int count);
        protected abstract void Restart();
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) {
            if (origin != SeekOrigin.Begin) throw new NotSupportedException();
            Position = offset;
            return _position;
        }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing && _ownsInner) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class ZipCryptoStream : DecryptingStream {
        private readonly ZipCryptoKeys _initial;
        private ZipCryptoKeys _keys;
        internal ZipCryptoStream(Stream inner, ZipCryptoKeys keys) : base(inner, ownsInner: true) { _initial = _keys = keys; }
        protected override void Transform(byte[] buffer, int offset, int count) => _keys.Decrypt(buffer, offset, count);
        protected override void Restart() => _keys = _initial;
    }

    /// <summary>AES in counter mode as WinZip defines it: a 128-bit little-endian counter starting at 1 and no nonce.</summary>
    private sealed class WinZipAesCtrStream : DecryptingStream {
        private const int BlockBytes = 4096;
        private readonly Aes _aes;
        private readonly ICryptoTransform _encryptor;
        private readonly byte[] _counters = new byte[BlockBytes];
        private readonly byte[] _keystream = new byte[BlockBytes];
        private int _used = BlockBytes;
        private ulong _counter;

        internal WinZipAesCtrStream(Stream inner, byte[] key, bool ownsInner) : base(inner, ownsInner) {
            _aes = Aes.Create();
            _aes.Mode = CipherMode.ECB;
            _aes.Padding = PaddingMode.None;
            _aes.Key = key;
            _encryptor = _aes.CreateEncryptor();
        }

        protected override void Transform(byte[] buffer, int offset, int count) {
            for (int index = offset, end = offset + count; index < end; index++) {
                if (_used == BlockBytes) Refill();
                buffer[index] ^= _keystream[_used++];
            }
        }

        protected override void Restart() {
            _counter = 0;
            _used = BlockBytes;
        }

        private void Refill() {
            for (int block = 0; block < BlockBytes; block += 16) {
                ulong value = ++_counter;
                for (int index = 0; index < 8; index++) _counters[block + index] = (byte)(value >> (8 * index));
            }
            _encryptor.TransformBlock(_counters, 0, BlockBytes, _keystream, 0);
            _used = 0;
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _encryptor.Dispose();
                _aes.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>Feeds each ciphertext byte through an HMAC once, however often the reader above rewinds.</summary>
    private sealed class HmacReadStream : Stream {
        private readonly Stream _inner;
        private readonly HMAC _hmac;
        private long _hashed;
        internal HmacReadStream(Stream inner, HMAC hmac) { _inner = inner; _hmac = hmac; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override int Read(byte[] buffer, int offset, int count) {
            long start = _inner.Position;
            int read = _inner.Read(buffer, offset, count);
            long fresh = start + read - Math.Max(start, _hashed);
            if (fresh > 0) {
                _hmac.TransformBlock(buffer, offset + (int)(read - fresh), (int)fresh, null, 0);
                _hashed = start + read;
            }
            return read;
        }
        internal void Drain(CancellationToken token) {
            _inner.Position = _hashed;
            var buffer = new byte[65536];
            while (true) {
                token.ThrowIfCancellationRequested();
                if (Read(buffer, 0, buffer.Length) == 0) return;
            }
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
