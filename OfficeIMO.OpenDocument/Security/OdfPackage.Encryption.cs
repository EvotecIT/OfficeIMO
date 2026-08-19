namespace OfficeIMO.OpenDocument;

internal sealed partial class OdfPackage {
    private static void DecryptLoadedEntries(IReadOnlyList<OdfPackageEntry> loaded, XElement manifestRoot,
        OdfPackageEntry manifestEntry, OdfLoadOptions options) {
        List<XElement> encryptedFileEntries = manifestRoot.Elements(OdfNamespaces.Manifest + "file-entry")
            .Where(element => element.Element(OdfNamespaces.Manifest + "encryption-data") != null)
            .ToList();
        if (encryptedFileEntries.Count == 0) return;
        if (string.IsNullOrEmpty(options.Password)) {
            throw new OdfEncryptedPackageException("This OpenDocument package is encrypted; supply OdfLoadOptions.Password to open it.",
                OdfEncryptionFailureReason.PasswordRequired);
        }

        string[] encryptedPathList = encryptedFileEntries.Select(element =>
            ReadRequiredAttribute(element, "full-path", null)).ToArray();
        var encryptedPaths = new HashSet<string>(encryptedPathList, StringComparer.Ordinal);
        if (encryptedPaths.Count != encryptedPathList.Length) {
            throw InvalidMetadata("Encrypted ODF manifest contains duplicate encrypted file entries.", null);
        }
        long finalUncompressedBytes = 0;
        long totalKdfIterations = 0;
        try {
            foreach (OdfPackageEntry entry in loaded.Where(entry => !encryptedPaths.Contains(entry.Name))) {
                finalUncompressedBytes = checked(finalUncompressedBytes + entry.GetOriginalBytes().LongLength);
            }
            foreach (XElement fileEntry in encryptedFileEntries) {
                string path = ReadRequiredAttribute(fileEntry, "full-path", null);
                XElement encryptionData = fileEntry.Element(OdfNamespaces.Manifest + "encryption-data")!;
                XElement derivation = encryptionData.Element(OdfNamespaces.Manifest + "key-derivation")
                    ?? throw InvalidMetadata("Encrypted ODF entry is missing manifest:key-derivation.", path);
                int iterationCount = ReadIntAttribute(derivation, "iteration-count", path);
                try {
                    OdfPasswordEncryption.ValidateIterationCount(iterationCount);
                } catch (ArgumentOutOfRangeException ex) {
                    throw new OdfEncryptedPackageException(
                        "Encrypted ODF entry uses a PBKDF2 iteration count outside the supported security policy.",
                        OdfEncryptionFailureReason.UnsupportedProfile, path, ex);
                }
                totalKdfIterations = checked(totalKdfIterations + iterationCount);
                if (totalKdfIterations > options.MaxTotalKdfIterations) {
                    throw new OdfEncryptedPackageException(
                        $"Encrypted ODF package exceeds MaxTotalKdfIterations ({options.MaxTotalKdfIterations}).",
                        OdfEncryptionFailureReason.ResourceLimitExceeded, path);
                }
                long originalSize = ReadLongAttribute(fileEntry, "size", path);
                if (originalSize > options.MaxEntryUncompressedBytes) {
                    throw new OdfEncryptedPackageException(
                        $"Decrypted ODF entry '{path}' exceeds MaxEntryUncompressedBytes ({options.MaxEntryUncompressedBytes}).",
                        OdfEncryptionFailureReason.ResourceLimitExceeded, path);
                }
                finalUncompressedBytes = checked(finalUncompressedBytes + originalSize);
                if (finalUncompressedBytes > options.MaxTotalUncompressedBytes) {
                    throw new OdfEncryptedPackageException(
                        $"Decrypted ODF package exceeds MaxTotalUncompressedBytes ({options.MaxTotalUncompressedBytes}).",
                        OdfEncryptionFailureReason.ResourceLimitExceeded, path);
                }
            }
        } catch (OverflowException ex) {
            throw new OdfEncryptedPackageException("Decrypted ODF package size metadata exceeds supported limits.",
                OdfEncryptionFailureReason.ResourceLimitExceeded, null, ex);
        }

        byte[] startKey;
        try {
            startKey = OdfPasswordEncryption.CreateStartKey(options.Password!);
        } catch (ArgumentException ex) {
            throw new OdfEncryptedPackageException(ex.Message, OdfEncryptionFailureReason.PasswordRequired, null, ex);
        }

        try {
            foreach (XElement fileEntry in encryptedFileEntries) {
                string path = ReadRequiredAttribute(fileEntry, "full-path", null);
                OdfPackageEntry entry = loaded.FirstOrDefault(candidate => string.Equals(candidate.Name, path, StringComparison.Ordinal))
                    ?? throw InvalidMetadata("Encrypted ODF manifest entry has no matching ZIP entry.", path);
                XElement encryptionData = fileEntry.Element(OdfNamespaces.Manifest + "encryption-data")!;
                XElement algorithm = encryptionData.Element(OdfNamespaces.Manifest + "algorithm")
                    ?? throw InvalidMetadata("Encrypted ODF entry is missing manifest:algorithm.", path);
                XElement derivation = encryptionData.Element(OdfNamespaces.Manifest + "key-derivation")
                    ?? throw InvalidMetadata("Encrypted ODF entry is missing manifest:key-derivation.", path);
                XElement startKeyGeneration = encryptionData.Element(OdfNamespaces.Manifest + "start-key-generation")
                    ?? throw InvalidMetadata("Encrypted ODF entry is missing manifest:start-key-generation.", path);

                string algorithmName = ReadRequiredAttribute(algorithm, "algorithm-name", path);
                string checksumType = ReadRequiredAttribute(encryptionData, "checksum-type", path);
                string derivationName = ReadRequiredAttribute(derivation, "key-derivation-name", path);
                string startKeyName = ReadRequiredAttribute(startKeyGeneration, "start-key-generation-name", path);
                int keySize = ReadIntAttribute(derivation, "key-size", path);
                int startKeySize = ReadIntAttribute(startKeyGeneration, "key-size", path);
                if (!string.Equals(algorithmName, OdfPasswordEncryption.Aes256Cbc, StringComparison.Ordinal) ||
                    !string.Equals(checksumType, OdfPasswordEncryption.Sha256OneKilobyte, StringComparison.Ordinal) ||
                    !IsPbkdf2(derivationName) ||
                    !IsSha256StartKey(startKeyName) ||
                    keySize != 32 || startKeySize != 32) {
                    throw new OdfEncryptedPackageException(
                        $"Encrypted ODF entry '{path}' uses a profile that is not supported. Supported input is AES-256-CBC, PBKDF2-HMAC-SHA1, SHA-256 start key, and SHA-256/1K checksum.",
                        OdfEncryptionFailureReason.UnsupportedProfile, path);
                }

                byte[] salt = ReadBase64Attribute(derivation, "salt", path, 1368);
                byte[] iv = ReadBase64Attribute(algorithm, "initialisation-vector", path, 24);
                byte[] checksum = ReadBase64Attribute(encryptionData, "checksum", path, 44);
                int iterationCount = ReadIntAttribute(derivation, "iteration-count", path);
                long originalSize = ReadLongAttribute(fileEntry, "size", path);
                byte[] plaintext = OdfPasswordEncryption.Decrypt(entry.GetOriginalBytes(), startKey, salt, iv,
                    iterationCount, checksum, originalSize, options.MaxEntryUncompressedBytes, path);
                entry.ReplaceLoadedBytes(plaintext);
                encryptionData.Remove();
                fileEntry.Attribute(OdfNamespaces.Manifest + "size")?.Remove();
            }
        } finally {
            Array.Clear(startKey, 0, startKey.Length);
        }

        manifestEntry.ReplaceLoadedBytes(OdfXmlCodec.Save(manifestRoot.Document!));
    }

    private void ValidateEncryptionSaveOptions(OdfSaveOptions options) {
        if (_sourceIsEncrypted && options.Encryption == null && options.EncryptionHandling != OdfEncryptionHandling.Remove) {
            throw new OdfEncryptedPackageException(
                "This document was loaded from an encrypted package. Supply OdfSaveOptions.Encryption to preserve protection or set EncryptionHandling to Remove explicitly.",
                OdfEncryptionFailureReason.PreservationRequired);
        }
        if (options.Encryption != null && options.EncryptionHandling == OdfEncryptionHandling.Remove) {
            throw new ArgumentException("Encryption and EncryptionHandling.Remove cannot be requested together.", nameof(options));
        }
        if (options.Encryption != null) {
            byte[] validationStartKey = OdfPasswordEncryption.CreateStartKey(options.Encryption.Password);
            Array.Clear(validationStartKey, 0, validationStartKey.Length);
            OdfPasswordEncryption.ValidateIterationCount(options.Encryption.IterationCount);
        }
    }

    private List<OdfZipWriteEntry> CreateOutputEntries(OdfSaveOptions options, bool outputEncrypted) {
        var outputEntries = new List<OdfZipWriteEntry>();
        OdfPackageEntry mimetype = GetRequiredEntry("mimetype");
        outputEntries.Add(new OdfZipWriteEntry(mimetype.Name, mimetype.GetBytesForSave(), compress: false));

        IEnumerable<OdfPackageEntry> remaining = _entries.Where(entry => !entry.IsRemoved && entry.Name != "mimetype");
        if (options.Deterministic) {
            OdfPackageEntry[] original = remaining.Where(entry => !entry.IsNew).ToArray();
            OdfPackageEntry[] added = remaining.Where(entry => entry.IsNew).OrderBy(entry => entry.Name, StringComparer.Ordinal).ToArray();
            remaining = original.Concat(added);
        }
        OdfPackageEntry[] ordered = remaining.ToArray();
        if (!outputEncrypted) {
            foreach (OdfPackageEntry entry in ordered) {
                outputEntries.Add(new OdfZipWriteEntry(entry.Name, entry.GetBytesForSave(),
                    compress: !entry.Name.EndsWith("/", StringComparison.Ordinal)));
            }
            return outputEntries;
        }

        OdfEncryptionOptions encryption = options.Encryption!;
        byte[] startKey = OdfPasswordEncryption.CreateStartKey(encryption.Password);
        XDocument manifest = new XDocument(GetRequiredEntry("META-INF/manifest.xml")
            .GetXml(_loadOptions.MaxXmlCharacters, _loadOptions.MaxXmlDepth));
        XElement manifestRoot = manifest.Root ?? throw new InvalidDataException("OpenDocument manifest has no root element.");
        var encrypted = new Dictionary<string, OdfEncryptedEntry>(StringComparer.Ordinal);
        try {
            foreach (OdfPackageEntry entry in ordered.Where(IsEncryptionEligible)) {
                OdfEncryptedEntry encryptedEntry = OdfPasswordEncryption.Encrypt(entry.GetBytesForSave(), startKey,
                    encryption.IterationCount);
                encrypted.Add(entry.Name, encryptedEntry);
                AddEncryptionMetadata(manifestRoot, entry.Name, encryptedEntry);
            }
        } finally {
            Array.Clear(startKey, 0, startKey.Length);
        }
        byte[] manifestBytes = OdfXmlCodec.Save(manifest);

        foreach (OdfPackageEntry entry in ordered) {
            if (string.Equals(entry.Name, "META-INF/manifest.xml", StringComparison.Ordinal)) {
                outputEntries.Add(new OdfZipWriteEntry(entry.Name, manifestBytes, compress: true));
            } else if (encrypted.TryGetValue(entry.Name, out OdfEncryptedEntry? encryptedEntry)) {
                outputEntries.Add(new OdfZipWriteEntry(entry.Name, encryptedEntry.Ciphertext, compress: false));
            } else {
                outputEntries.Add(new OdfZipWriteEntry(entry.Name, entry.GetBytesForSave(),
                    compress: !entry.Name.EndsWith("/", StringComparison.Ordinal)));
            }
        }
        return outputEntries;
    }

    private static bool IsEncryptionEligible(OdfPackageEntry entry) =>
        !entry.Name.EndsWith("/", StringComparison.Ordinal) &&
        !entry.Name.StartsWith("META-INF/", StringComparison.Ordinal);

    private static void AddEncryptionMetadata(XElement manifestRoot, string path, OdfEncryptedEntry encrypted) {
        XElement fileEntry = manifestRoot.Elements(OdfNamespaces.Manifest + "file-entry")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Manifest + "full-path"), path,
                StringComparison.Ordinal))
            ?? throw new InvalidDataException($"OpenDocument manifest is missing file entry '{path}'.");
        fileEntry.Element(OdfNamespaces.Manifest + "encryption-data")?.Remove();
        fileEntry.SetAttributeValue(OdfNamespaces.Manifest + "size", encrypted.OriginalSize.ToString(CultureInfo.InvariantCulture));
        fileEntry.Add(new XElement(OdfNamespaces.Manifest + "encryption-data",
            new XAttribute(OdfNamespaces.Manifest + "checksum-type", OdfPasswordEncryption.Sha256OneKilobyte),
            new XAttribute(OdfNamespaces.Manifest + "checksum", Convert.ToBase64String(encrypted.Checksum)),
            new XElement(OdfNamespaces.Manifest + "algorithm",
                new XAttribute(OdfNamespaces.Manifest + "algorithm-name", OdfPasswordEncryption.Aes256Cbc),
                new XAttribute(OdfNamespaces.Manifest + "initialisation-vector", Convert.ToBase64String(encrypted.InitializationVector))),
            new XElement(OdfNamespaces.Manifest + "start-key-generation",
                new XAttribute(OdfNamespaces.Manifest + "start-key-generation-name", OdfPasswordEncryption.Sha256),
                new XAttribute(OdfNamespaces.Manifest + "key-size", 32)),
            new XElement(OdfNamespaces.Manifest + "key-derivation",
                new XAttribute(OdfNamespaces.Manifest + "key-derivation-name", OdfPasswordEncryption.Pbkdf2),
                new XAttribute(OdfNamespaces.Manifest + "iteration-count", encrypted.IterationCount),
                new XAttribute(OdfNamespaces.Manifest + "key-size", 32),
                new XAttribute(OdfNamespaces.Manifest + "salt", Convert.ToBase64String(encrypted.Salt)))));
    }

    private static bool IsPbkdf2(string value) =>
        string.Equals(value, OdfPasswordEncryption.Pbkdf2, StringComparison.Ordinal) ||
        string.Equals(value, "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#pbkdf2", StringComparison.Ordinal);

    private static bool IsSha256StartKey(string value) =>
        string.Equals(value, OdfPasswordEncryption.Sha256, StringComparison.Ordinal) ||
        string.Equals(value, OdfPasswordEncryption.Sha256XmlEncryptionAlias, StringComparison.Ordinal);

    private static string ReadRequiredAttribute(XElement element, string localName, string? entryPath) {
        string? value = (string?)element.Attribute(OdfNamespaces.Manifest + localName);
        if (string.IsNullOrEmpty(value)) throw InvalidMetadata($"Encrypted ODF metadata is missing manifest:{localName}.", entryPath);
        return value!;
    }

    private static byte[] ReadBase64Attribute(XElement element, string localName, string entryPath, int maxEncodedLength) {
        string value = ReadRequiredAttribute(element, localName, entryPath);
        if (value.Length > maxEncodedLength) {
            throw InvalidMetadata($"Encrypted ODF manifest:{localName} exceeds its encoded-length limit.", entryPath);
        }
        try { return Convert.FromBase64String(value); }
        catch (FormatException ex) { throw InvalidMetadata($"Encrypted ODF manifest:{localName} is not valid base64.", entryPath, ex); }
    }

    private static int ReadIntAttribute(XElement element, string localName, string entryPath) {
        string value = ReadRequiredAttribute(element, localName, entryPath);
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int result) || result < 0) {
            throw InvalidMetadata($"Encrypted ODF manifest:{localName} is not a valid non-negative integer.", entryPath);
        }
        return result;
    }

    private static long ReadLongAttribute(XElement element, string localName, string entryPath) {
        string value = ReadRequiredAttribute(element, localName, entryPath);
        if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long result) || result < 0) {
            throw InvalidMetadata($"Encrypted ODF manifest:{localName} is not a valid non-negative integer.", entryPath);
        }
        return result;
    }

    private static OdfEncryptedPackageException InvalidMetadata(string message, string? entryPath,
        Exception? innerException = null) => new OdfEncryptedPackageException(message,
        OdfEncryptionFailureReason.InvalidEncryptedPackage, entryPath, innerException);
}
