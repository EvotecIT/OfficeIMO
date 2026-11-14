using System.Globalization;
using System.Text;
using Org.BouncyCastle.Bcpg;
using Org.BouncyCastle.Bcpg.Sig;
using Org.BouncyCastle.Bcpg.OpenPgp;

namespace OfficeIMO.Security;

/// <summary>
/// OpenPGP verification and decryption over BouncyCastle's RFC 4880 implementation. Keys come only from the
/// caller's <see cref="OfficeOpenPgpKeySet"/>. Every failure is reported as a status, never as an exception.
/// </summary>
internal static class OpenPgpService {
    private const int MaxKeyRings = 4096;

    internal static OfficeOpenPgpVerificationResult VerifyDetached(byte[] signature, byte[] signedContent,
        OfficeOpenPgpKeySet keys, OfficeOpenPgpOptions? options) {
        ArgumentNullException.ThrowIfNull(signature);
        ArgumentNullException.ThrowIfNull(signedContent);
        ArgumentNullException.ThrowIfNull(keys);
        options ??= new OfficeOpenPgpOptions();
        ValidateOptions(options);
        if (signature.LongLength > options.MaxInputBytes || signedContent.LongLength > options.MaxInputBytes) {
            return new OfficeOpenPgpVerificationResult(null, "The signature or signed content exceeds the configured input limit.");
        }

        if (!TryReadSignatures(signature, options, out List<PgpSignature>? signatures, out string? error)) {
            OfficeOpenPgpSignatureInfo failure = Unreadable(signature, error!);
            return new OfficeOpenPgpVerificationResult(new[] { failure }, failure.Message);
        }

        KeyIndex index = KeyIndex.Load(keys, options);
        return new OfficeOpenPgpVerificationResult(signatures!.Select(item => Evaluate(item, signedContent, index, options)),
            index.Error);
    }

    internal static OfficeOpenPgpVerificationResult VerifyCleartext(byte[] armoredMessage, OfficeOpenPgpKeySet keys,
        OfficeOpenPgpOptions? options) {
        ArgumentNullException.ThrowIfNull(armoredMessage);
        ArgumentNullException.ThrowIfNull(keys);
        options ??= new OfficeOpenPgpOptions();
        ValidateOptions(options);
        if (armoredMessage.LongLength > options.MaxInputBytes) {
            return new OfficeOpenPgpVerificationResult(null, "The message exceeds the configured input limit.");
        }

        if (!OpenPgpCleartext.TryParse(armoredMessage, out byte[]? signedText, out byte[]? signatureBlock, out string? parseError)) {
            return new OfficeOpenPgpVerificationResult(new[] { Malformed(parseError!) }, parseError);
        }

        if (!TryReadSignatures(signatureBlock!, options, out List<PgpSignature>? signatures, out string? error)) {
            OfficeOpenPgpSignatureInfo failure = Unreadable(signatureBlock!, error!);
            return new OfficeOpenPgpVerificationResult(new[] { failure }, failure.Message, signedText);
        }

        KeyIndex index = KeyIndex.Load(keys, options);
        return new OfficeOpenPgpVerificationResult(signatures!.Select(item => Evaluate(item, signedText!, index, options)),
            index.Error, signedText);
    }

    internal static OfficeOpenPgpMessageResult OpenMessage(byte[] message, OfficeOpenPgpKeySet keys, OfficeOpenPgpOptions? options) {
        ArgumentNullException.ThrowIfNull(message);
        ArgumentNullException.ThrowIfNull(keys);
        options ??= new OfficeOpenPgpOptions();
        ValidateOptions(options);
        if (message.LongLength > options.MaxInputBytes) {
            return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.LimitExceeded, "The message exceeds the configured input limit.");
        }

        try {
            using Stream decoded = PgpUtilities.GetDecoderStream(new MemoryStream(message, writable: false));
            var factory = new PgpObjectFactory(decoded);
            PgpObject? first = NextSkippingMarkers(factory);
            if (first is PgpEncryptedDataList encrypted) {
                OfficeOpenPgpMessageResult decrypted = Decrypt(encrypted, keys, options);
                return decrypted.Status == OfficeOpenPgpMessageStatus.Malformed ? Unreadable(message, decrypted) : decrypted;
            }

            if (first == null) {
                return Unreadable(message, new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Malformed,
                    "The message contains no OpenPGP packets."));
            }

            KeyIndex index = KeyIndex.Load(keys, options);
            var reader = new LiteralReader(options, index);
            OfficeOpenPgpMessageResult? failure = reader.Read(first, factory, 0);
            return failure ?? reader.ToResult(isEncrypted: false, integrityProtected: false, symmetricAlgorithm: null,
                recipients: Array.Empty<string>());
        } catch (Exception ex) when (IsParseFailure(ex)) {
            return Unreadable(message, new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Malformed,
                "The OpenPGP message could not be parsed: " + ex.Message));
        }
    }

    private static OfficeOpenPgpMessageResult Unreadable(byte[] message, OfficeOpenPgpMessageResult malformed) =>
        OpenPgpPacketVersion.TryDescribeUnsupported(message, out string unsupported)
            ? new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Unsupported, unsupported, malformed.IsEncrypted)
            : malformed;

    private static OfficeOpenPgpSignatureInfo Unreadable(byte[] signature, string error) =>
        OpenPgpPacketVersion.TryDescribeUnsupported(signature, out string unsupported)
            ? new OfficeOpenPgpSignatureInfo(OfficeOpenPgpSignatureStatus.Unsupported, unsupported)
            : Malformed(error);

    private static OfficeOpenPgpMessageResult Decrypt(PgpEncryptedDataList list, OfficeOpenPgpKeySet keys, OfficeOpenPgpOptions options) {
        var publicKeyItems = new List<PgpPublicKeyEncryptedData>();
        bool passwordOnly = false;
        foreach (PgpEncryptedData item in list.GetEncryptedDataObjects()) {
            if (item is PgpPublicKeyEncryptedData publicKeyItem) publicKeyItems.Add(publicKeyItem);
            else passwordOnly = true;
        }

        string[] recipients = publicKeyItems.Select(item => KeyIdHex(item.KeyId)).ToArray();
        if (publicKeyItems.Count == 0) {
            return new OfficeOpenPgpMessageResult(passwordOnly ? OfficeOpenPgpMessageStatus.Unsupported : OfficeOpenPgpMessageStatus.Malformed,
                passwordOnly ? "Password-encrypted OpenPGP messages are not supported." : "The message lists no recipients.",
                isEncrypted: true);
        }

        List<PgpSecretKeyRing> secretRings = LoadSecretRings(keys, options, out string? keyError);
        OfficeOpenPgpMessageStatus failure = OfficeOpenPgpMessageStatus.NoMatchingSecretKey;
        string failureMessage = keyError ?? "No supplied secret key matches a recipient of the message.";
        foreach (PgpPublicKeyEncryptedData item in publicKeyItems) {
            IEnumerable<PgpSecretKey> candidates = item.KeyId == 0
                ? secretRings.SelectMany(ring => ring.GetSecretKeys()).Where(key => key.PublicKey.IsEncryptionKey)
                : secretRings.Select(ring => ring.GetSecretKey(item.KeyId)).Where(key => key != null);
            foreach (PgpSecretKey secretKey in candidates) {
                PgpPrivateKey? privateKey = TryUnlock(secretKey, keys, out OfficeOpenPgpMessageStatus unlockStatus);
                if (privateKey == null) {
                    if (failure == OfficeOpenPgpMessageStatus.NoMatchingSecretKey || unlockStatus == OfficeOpenPgpMessageStatus.WrongPassphrase) {
                        failure = unlockStatus;
                        failureMessage = unlockStatus == OfficeOpenPgpMessageStatus.PassphraseRequired
                            ? "The secret key " + KeyIdHex(secretKey.KeyId) + " is protected and no passphrase was supplied."
                            : "The passphrase did not unlock secret key " + KeyIdHex(secretKey.KeyId) + ".";
                    }

                    continue;
                }

                return DecryptWith(item, privateKey, keys, options, recipients);
            }
        }

        return new OfficeOpenPgpMessageResult(failure, failureMessage, isEncrypted: true, recipientKeyIds: recipients);
    }

    private static OfficeOpenPgpMessageResult DecryptWith(PgpPublicKeyEncryptedData item, PgpPrivateKey privateKey,
        OfficeOpenPgpKeySet keys, OfficeOpenPgpOptions options, string[] recipients) {
        bool protectedData = item.IsIntegrityProtected();
        if (!protectedData && !options.AllowMissingIntegrityProtection) {
            return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.MissingIntegrityProtection,
                "The encrypted data has no modification detection code; the plaintext was not released.",
                isEncrypted: true, recipientKeyIds: recipients);
        }

        string? symmetric = null;
        try {
            symmetric = item.GetSymmetricAlgorithm(privateKey).ToString();
            using Stream clear = item.GetDataStream(privateKey);
            var factory = new PgpObjectFactory(clear);
            KeyIndex index = KeyIndex.Load(keys, options);
            var reader = new LiteralReader(options, index);
            PgpObject? first = NextSkippingMarkers(factory);
            if (first == null) {
                return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Malformed, "The encrypted data is empty.",
                    isEncrypted: true, recipientKeyIds: recipients);
            }

            OfficeOpenPgpMessageResult? failure = reader.Read(first, factory, 0);
            if (failure != null) {
                return new OfficeOpenPgpMessageResult(failure.Status, failure.Message, isEncrypted: true, recipientKeyIds: recipients);
            }

            // The modification detection code trails the literal data; drain before checking it.
            Drain(clear, options.MaxOutputBytes);
            if (protectedData && !item.Verify()) {
                return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.IntegrityCheckFailed,
                    "The modification detection code does not match; the message was altered and its plaintext was not released.",
                    isEncrypted: true, recipientKeyIds: recipients);
            }

            return reader.ToResult(isEncrypted: true, protectedData, symmetric, recipients);
        } catch (Exception ex) when (IsParseFailure(ex)) {
            // Altered ciphertext usually breaks packet parsing before the modification detection code is reached.
            OfficeOpenPgpMessageStatus status = protectedData
                ? OfficeOpenPgpMessageStatus.IntegrityCheckFailed
                : OfficeOpenPgpMessageStatus.Malformed;
            return new OfficeOpenPgpMessageResult(status,
                "The encrypted data could not be decrypted or failed its integrity check: " + ex.Message,
                isEncrypted: true, symmetricAlgorithm: symmetric, recipientKeyIds: recipients);
        }
    }

    private static PgpPrivateKey? TryUnlock(PgpSecretKey secretKey, OfficeOpenPgpKeySet keys, out OfficeOpenPgpMessageStatus status) {
        status = OfficeOpenPgpMessageStatus.Opened;
        if (secretKey.KeyEncryptionAlgorithm == SymmetricKeyAlgorithmTag.Null) {
            try {
                return secretKey.ExtractPrivateKey(Array.Empty<char>());
            } catch (Exception ex) when (IsParseFailure(ex)) {
                status = OfficeOpenPgpMessageStatus.Malformed;
                return null;
            }
        }

        char[]? passphrase = keys.PassphraseProvider?.Invoke(KeyIdHex(secretKey.KeyId));
        if (passphrase == null) {
            status = OfficeOpenPgpMessageStatus.PassphraseRequired;
            return null;
        }

        try {
            return secretKey.ExtractPrivateKey(passphrase);
        } catch (Exception ex) when (IsParseFailure(ex)) {
            status = OfficeOpenPgpMessageStatus.WrongPassphrase;
            return null;
        }
    }

    internal static OfficeOpenPgpSignatureInfo Evaluate(PgpSignature signature, byte[] data, KeyIndex index, OfficeOpenPgpOptions options) {
        string keyId = KeyIdHex(signature.KeyId);
        string hash = signature.HashAlgorithm.ToString();
        string algorithm = signature.KeyAlgorithm.ToString();
        DateTimeOffset created = new DateTimeOffset(DateTime.SpecifyKind(signature.CreationTime, DateTimeKind.Utc));
        bool text = signature.SignatureType == PgpSignature.CanonicalTextDocument;
        OfficeOpenPgpSignatureInfo Result(OfficeOpenPgpSignatureStatus status, string message, KeyIndex.Entry? entry = null) =>
            new OfficeOpenPgpSignatureInfo(status, message, keyId,
                entry == null ? null : Hex(entry.Key.GetFingerprint()),
                entry == null ? null : Hex(entry.Primary.GetFingerprint()),
                entry?.Primary.GetUserIds(), created, hash, algorithm, signature.Version, entry?.Key.Version ?? 0, text);

        if (signature.SignatureType != PgpSignature.BinaryDocument && signature.SignatureType != PgpSignature.CanonicalTextDocument) {
            return Result(OfficeOpenPgpSignatureStatus.Unsupported, "Only binary and text document signatures are accepted for messages.");
        }

        if (!options.AllowWeakHashAlgorithms && (signature.HashAlgorithm == HashAlgorithmTag.MD5 ||
            signature.HashAlgorithm == HashAlgorithmTag.Sha1 || signature.HashAlgorithm == HashAlgorithmTag.RipeMD160 ||
            signature.HashAlgorithm == HashAlgorithmTag.MD2)) {
            return Result(OfficeOpenPgpSignatureStatus.WeakAlgorithm, "The signature uses the weak hash algorithm " + hash + ".");
        }

        IReadOnlyList<KeyIndex.Entry> candidates = index.Find(signature.KeyId);
        if (candidates.Count == 0) {
            return Result(OfficeOpenPgpSignatureStatus.MissingPublicKey, "No supplied public key has key ID " + keyId + ".");
        }

        OfficeOpenPgpSignatureInfo? firstFailure = null;
        foreach (KeyIndex.Entry entry in candidates) {
            bool matches;
            try {
                signature.InitVerify(entry.Key);
                signature.Update(data);
                matches = signature.Verify();
            } catch (Exception ex) when (IsParseFailure(ex)) {
                firstFailure ??= Result(OfficeOpenPgpSignatureStatus.Unsupported, "The signature could not be checked: " + ex.Message, entry);
                continue;
            }

            if (!matches) {
                firstFailure ??= Result(OfficeOpenPgpSignatureStatus.Invalid, "The signature does not match the signed data.", entry);
                continue;
            }

            return CheckKeyPolicy(signature, entry, created, options, Result);
        }

        return firstFailure!;
    }

    private static OfficeOpenPgpSignatureInfo CheckKeyPolicy(PgpSignature signature, KeyIndex.Entry entry, DateTimeOffset created,
        OfficeOpenPgpOptions options, Func<OfficeOpenPgpSignatureStatus, string, KeyIndex.Entry?, OfficeOpenPgpSignatureInfo> result) {
        PgpPublicKey key = entry.Key;
        if (IsFiniteFieldKey(key.Algorithm) && key.BitStrength < options.MinimumKeyBits) {
            return result(OfficeOpenPgpSignatureStatus.WeakAlgorithm,
                "The signing key has " + key.BitStrength.ToString(CultureInfo.InvariantCulture) + " bits; at least " +
                options.MinimumKeyBits.ToString(CultureInfo.InvariantCulture) + " are required.", entry);
        }

        if (entry.Primary.HasRevocation() || HasSubkeyRevocation(key)) {
            return result(OfficeOpenPgpSignatureStatus.KeyRevoked, "The signing key carries a revocation signature.", entry);
        }

        if (!key.IsMasterKey) {
            PgpSignature? binding = FindValidBinding(entry.Primary, key);
            if (binding == null) {
                return result(OfficeOpenPgpSignatureStatus.KeyNotValidForSigning,
                    "The signing subkey has no binding signature from its primary key that verifies.", entry);
            }

            if (binding.HasSubpackets && binding.GetHashedSubPackets().HasSubpacket(SignatureSubpacketTag.KeyFlags) &&
                (binding.GetHashedSubPackets().GetKeyFlags() & KeyFlags.SignData) == 0) {
                return result(OfficeOpenPgpSignatureStatus.KeyNotValidForSigning, "The signing subkey is not flagged for signing.", entry);
            }
        }

        DateTimeOffset keyCreated = new DateTimeOffset(DateTime.SpecifyKind(key.CreationTime, DateTimeKind.Utc));
        long validSeconds = SafeValidSeconds(key);
        if (created < keyCreated || (validSeconds > 0 && created > keyCreated.AddSeconds(validSeconds))) {
            return result(OfficeOpenPgpSignatureStatus.KeyExpired, "The signing key was not valid when the signature was made.", entry);
        }

        DateTimeOffset now = options.VerificationTime ?? DateTimeOffset.UtcNow;
        if (signature.HasSubpackets) {
            long expiration = signature.GetHashedSubPackets().GetSignatureExpirationTime();
            if (expiration > 0 && now > created.AddSeconds(expiration)) {
                return result(OfficeOpenPgpSignatureStatus.SignatureExpired, "The signature has expired.", entry);
            }
        }

        return result(OfficeOpenPgpSignatureStatus.Valid, "The signature verified.", entry);
    }

    private static PgpSignature? FindValidBinding(PgpPublicKey primary, PgpPublicKey subkey) {
        foreach (PgpSignature binding in subkey.GetSignaturesOfType(PgpSignature.SubkeyBinding)) {
            try {
                binding.InitVerify(primary);
                if (binding.VerifyCertification(primary, subkey)) return binding;
            } catch (Exception ex) when (IsParseFailure(ex)) {
                // Try the next binding signature.
            }
        }

        return null;
    }

    private static bool HasSubkeyRevocation(PgpPublicKey key) =>
        !key.IsMasterKey && key.GetSignaturesOfType(PgpSignature.SubkeyRevocation).Any();

    private static long SafeValidSeconds(PgpPublicKey key) {
        try {
            return key.GetValidSeconds();
        } catch (Exception ex) when (IsParseFailure(ex)) {
            return 0;
        }
    }

    private static bool IsFiniteFieldKey(PublicKeyAlgorithmTag algorithm) =>
        algorithm == PublicKeyAlgorithmTag.RsaGeneral || algorithm == PublicKeyAlgorithmTag.RsaSign ||
        algorithm == PublicKeyAlgorithmTag.RsaEncrypt || algorithm == PublicKeyAlgorithmTag.Dsa ||
        algorithm == PublicKeyAlgorithmTag.ElGamalEncrypt || algorithm == PublicKeyAlgorithmTag.ElGamalGeneral;

    private static bool TryReadSignatures(byte[] encoded, OfficeOpenPgpOptions options, out List<PgpSignature>? signatures, out string? error) {
        signatures = null;
        error = null;
        try {
            using Stream decoded = PgpUtilities.GetDecoderStream(new MemoryStream(encoded, writable: false));
            var factory = new PgpObjectFactory(decoded);
            var found = new List<PgpSignature>();
            PgpObject? item;
            int depth = 0;
            while ((item = factory.NextPgpObject()) != null) {
                if (item is PgpCompressedData compressed) {
                    if (++depth > options.MaxNestingDepth) {
                        error = "The signature data is nested too deeply.";
                        return false;
                    }

                    factory = new PgpObjectFactory(compressed.GetDataStream());
                    continue;
                }

                if (item is PgpMarker) continue;
                if (item is not PgpSignatureList list) {
                    error = "The data is not an OpenPGP signature.";
                    return false;
                }

                for (int index = 0; index < list.Count; index++) {
                    if (found.Count >= options.MaxSignatures) {
                        error = "The signature data holds more signatures than the configured limit.";
                        return false;
                    }

                    found.Add(list[index]);
                }
            }

            if (found.Count == 0) {
                error = "No OpenPGP signature packet was found.";
                return false;
            }

            signatures = found;
            return true;
        } catch (Exception ex) when (IsParseFailure(ex)) {
            error = "The OpenPGP signature could not be parsed: " + ex.Message;
            return false;
        }
    }

    private static List<PgpSecretKeyRing> LoadSecretRings(OfficeOpenPgpKeySet keys, OfficeOpenPgpOptions options, out string? error) {
        var rings = new List<PgpSecretKeyRing>();
        error = null;
        foreach (byte[] buffer in keys.SecretKeys) {
            if (buffer.LongLength > options.MaxInputBytes) {
                error = "A supplied secret key exceeds the configured input limit.";
                continue;
            }

            try {
                using Stream decoded = PgpUtilities.GetDecoderStream(new MemoryStream(buffer, writable: false));
                foreach (PgpSecretKeyRing ring in new PgpSecretKeyRingBundle(decoded).GetKeyRings()) {
                    if (rings.Count >= MaxKeyRings) break;
                    rings.Add(ring);
                }
            } catch (Exception ex) when (IsParseFailure(ex)) {
                error = "A supplied secret key could not be parsed: " + ex.Message;
            }
        }

        return rings;
    }

    private static PgpObject? NextSkippingMarkers(PgpObjectFactory factory) {
        PgpObject? item;
        do {
            item = factory.NextPgpObject();
        } while (item is PgpMarker);
        return item;
    }

    private static void Drain(Stream stream, long limit) {
        var buffer = new byte[8192];
        long total = 0;
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            total += read;
            if (total > limit) throw new InvalidDataException("The encrypted data exceeds the configured output limit.");
        }
    }

    internal static bool IsParseFailure(Exception ex) =>
        ex is PgpException or IOException or InvalidDataException or ArgumentException or InvalidOperationException or
            InvalidCastException or NotSupportedException or NotImplementedException or FormatException or
            IndexOutOfRangeException or Org.BouncyCastle.Crypto.CryptoException or Org.BouncyCastle.Security.SecurityUtilityException;

    private static void ValidateOptions(OfficeOpenPgpOptions options) {
        if (options.MaxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "OfficeOpenPgpOptions.MaxInputBytes must be positive.");
        if (options.MaxOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "OfficeOpenPgpOptions.MaxOutputBytes must be positive.");
        if (options.MaxSignatures <= 0) throw new ArgumentOutOfRangeException(nameof(options), "OfficeOpenPgpOptions.MaxSignatures must be positive.");
        if (options.MaxNestingDepth <= 0) throw new ArgumentOutOfRangeException(nameof(options), "OfficeOpenPgpOptions.MaxNestingDepth must be positive.");
        if (options.MinimumKeyBits < 0) throw new ArgumentOutOfRangeException(nameof(options), "OfficeOpenPgpOptions.MinimumKeyBits must not be negative.");
    }

    private static OfficeOpenPgpSignatureInfo Malformed(string message) =>
        new OfficeOpenPgpSignatureInfo(OfficeOpenPgpSignatureStatus.Malformed, message);

    internal static string KeyIdHex(long keyId) => unchecked((ulong)keyId).ToString("X16", CultureInfo.InvariantCulture);

    private static string Hex(byte[] value) {
        var builder = new StringBuilder(value.Length * 2);
        foreach (byte item in value) builder.Append(item.ToString("X2", CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    /// <summary>Public keys from the caller's key set, indexed by key ID.</summary>
    internal sealed class KeyIndex {
        private readonly Dictionary<long, List<Entry>> _entries = new Dictionary<long, List<Entry>>();

        internal string? Error { get; private set; }

        internal IReadOnlyList<Entry> Find(long keyId) =>
            _entries.TryGetValue(keyId, out List<Entry>? list) ? list : (IReadOnlyList<Entry>)Array.Empty<Entry>();

        internal static KeyIndex Load(OfficeOpenPgpKeySet keys, OfficeOpenPgpOptions options) {
            var index = new KeyIndex();
            int rings = 0;
            foreach (byte[] buffer in keys.PublicKeys) {
                if (!index.TryLoad(buffer, options, ref rings, secret: false)) continue;
            }

            foreach (byte[] buffer in keys.SecretKeys) {
                index.TryLoad(buffer, options, ref rings, secret: true);
            }

            return index;
        }

        private bool TryLoad(byte[] buffer, OfficeOpenPgpOptions options, ref int rings, bool secret) {
            if (buffer.LongLength > options.MaxInputBytes) {
                Error = "A supplied key exceeds the configured input limit.";
                return false;
            }

            try {
                using Stream decoded = PgpUtilities.GetDecoderStream(new MemoryStream(buffer, writable: false));
                if (secret) {
                    foreach (PgpSecretKeyRing ring in new PgpSecretKeyRingBundle(decoded).GetKeyRings()) {
                        if (++rings > MaxKeyRings) return false;
                        PgpPublicKey primary = ring.GetPublicKey();
                        foreach (PgpSecretKey key in ring.GetSecretKeys()) Add(key.PublicKey, primary);
                    }
                } else {
                    foreach (PgpPublicKeyRing ring in new PgpPublicKeyRingBundle(decoded).GetKeyRings()) {
                        if (++rings > MaxKeyRings) return false;
                        PgpPublicKey primary = ring.GetPublicKey();
                        foreach (PgpPublicKey key in ring.GetPublicKeys()) Add(key, primary);
                    }
                }

                return true;
            } catch (Exception ex) when (IsParseFailure(ex)) {
                Error = "A supplied key could not be parsed: " + ex.Message;
                return false;
            }
        }

        private void Add(PgpPublicKey key, PgpPublicKey primary) {
            if (!_entries.TryGetValue(key.KeyId, out List<Entry>? list)) {
                list = new List<Entry>();
                _entries.Add(key.KeyId, list);
            }

            byte[] fingerprint = key.GetFingerprint();
            if (!list.Any(existing => existing.Key.GetFingerprint().SequenceEqual(fingerprint))) list.Add(new Entry(key, primary));
        }

        internal sealed class Entry {
            internal Entry(PgpPublicKey key, PgpPublicKey primary) {
                Key = key;
                Primary = primary;
            }

            internal PgpPublicKey Key { get; }

            internal PgpPublicKey Primary { get; }
        }
    }

    /// <summary>Reads compressed, one-pass-signed, literal, and signature packets of an opened message.</summary>
    private sealed class LiteralReader {
        private readonly OfficeOpenPgpOptions _options;
        private readonly KeyIndex _index;
        private readonly List<PgpSignature> _signatures = new List<PgpSignature>();
        private byte[]? _content;
        private string? _fileName;

        internal LiteralReader(OfficeOpenPgpOptions options, KeyIndex index) {
            _options = options;
            _index = index;
        }

        internal OfficeOpenPgpMessageResult? Read(PgpObject first, PgpObjectFactory factory, int depth) {
            PgpObject? item = first;
            while (item != null) {
                switch (item) {
                    case PgpMarker:
                    case PgpOnePassSignatureList:
                        break;
                    case PgpCompressedData compressed:
                        if (depth + 1 > _options.MaxNestingDepth) {
                            return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.LimitExceeded, "The message is nested too deeply.");
                        }

                        var inner = new PgpObjectFactory(compressed.GetDataStream());
                        PgpObject? innerFirst = NextSkippingMarkers(inner);
                        if (innerFirst == null) return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Malformed, "The compressed data is empty.");
                        OfficeOpenPgpMessageResult? nested = Read(innerFirst, inner, depth + 1);
                        if (nested != null) return nested;
                        break;
                    case PgpLiteralData literal:
                        if (_content != null) {
                            return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Malformed, "The message holds more than one literal data packet.");
                        }

                        _fileName = literal.FileName;
                        using (var buffer = new MemoryStream()) {
                            var chunk = new byte[8192];
                            Stream input = literal.GetInputStream();
                            int read;
                            while ((read = input.Read(chunk, 0, chunk.Length)) > 0) {
                                if (buffer.Length + read > _options.MaxOutputBytes) {
                                    return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.LimitExceeded,
                                        "The decrypted content exceeds the configured output limit.");
                                }

                                buffer.Write(chunk, 0, read);
                            }

                            _content = buffer.ToArray();
                        }

                        break;
                    case PgpSignatureList list:
                        for (int index = 0; index < list.Count; index++) {
                            if (_signatures.Count >= _options.MaxSignatures) {
                                return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.LimitExceeded,
                                    "The message holds more signatures than the configured limit.");
                            }

                            _signatures.Add(list[index]);
                        }

                        break;
                    case PgpEncryptedDataList:
                        return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Unsupported, "Nested encrypted data is not supported.");
                    default:
                        return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Unsupported,
                            "The message contains an unsupported packet: " + item.GetType().Name + ".");
                }

                item = factory.NextPgpObject();
            }

            return null;
        }

        internal OfficeOpenPgpMessageResult ToResult(bool isEncrypted, bool integrityProtected, string? symmetricAlgorithm,
            IReadOnlyList<string> recipients) {
            if (_content == null) {
                return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Malformed, "The message holds no literal data.",
                    isEncrypted, integrityProtected, recipientKeyIds: recipients);
            }

            byte[] content = _content;
            OfficeOpenPgpSignatureInfo[] signatures = _signatures.Select(signature => Evaluate(signature, content, _index, _options)).ToArray();
            return new OfficeOpenPgpMessageResult(OfficeOpenPgpMessageStatus.Opened,
                isEncrypted ? "The message was decrypted." : "The message was read.", isEncrypted, integrityProtected, content,
                _fileName, symmetricAlgorithm, recipients, signatures);
        }
    }
}
