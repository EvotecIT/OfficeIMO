using OfficeIMO.Shared.Packaging;

namespace OfficeIMO.OpenDocument;

internal sealed class OdfPackage {
    private static readonly DateTimeOffset DeterministicTimestamp = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
    private readonly List<OdfPackageEntry> _entries = new List<OdfPackageEntry>();
    private readonly Dictionary<string, OdfPackageEntry> _entriesByName = new Dictionary<string, OdfPackageEntry>(StringComparer.Ordinal);
    private readonly List<OdfDiagnostic> _diagnostics = new List<OdfDiagnostic>();
    private readonly OdfOpenOptions _openOptions;
    private bool _entryGraphChanged;

    private OdfPackage(OdfDocumentKind kind, OdfVersion version, OdfOpenOptions openOptions) {
        Kind = kind;
        Version = version;
        _openOptions = openOptions;
    }

    internal OdfDocumentKind Kind { get; }
    internal OdfVersion Version { get; private set; }
    internal string MediaType => OdfMediaTypes.ForKind(Kind);
    internal IReadOnlyList<OdfDiagnostic> Diagnostics => _diagnostics;
    internal IReadOnlyList<OdfPackageEntry> Entries => _entries.Where(entry => !entry.IsRemoved).ToList();
    internal bool IsSigned => _entries.Any(entry => !entry.IsRemoved && IsSignaturePath(entry.Name));

    internal static OdfPackage Create(OdfDocumentKind kind, OdfVersion version = OdfVersion.V1_4) {
        var package = new OdfPackage(kind, version, new OdfOpenOptions().Normalize());
        package.AddInitialEntry("mimetype", Encoding.ASCII.GetBytes(OdfMediaTypes.ForKind(kind)), OdfMediaTypes.ForKind(kind));
        package.AddInitialXml("content.xml", OdfPackageTemplates.CreateContent(kind, version), "text/xml");
        package.AddInitialXml("styles.xml", OdfPackageTemplates.CreateStyles(version), "text/xml");
        package.AddInitialXml("meta.xml", OdfPackageTemplates.CreateMetadata(version), "text/xml");
        package.AddInitialXml("settings.xml", OdfPackageTemplates.CreateSettings(version), "text/xml");
        package.AddInitialXml("META-INF/manifest.xml", OdfPackageTemplates.CreateManifest(kind, version), "text/xml");
        package._entryGraphChanged = true;
        return package;
    }

    internal static OdfPackage Open(string path, OdfOpenOptions? options, out string fullPath) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException($"OpenDocument file '{fullPath}' does not exist.", fullPath);
        OdfOpenOptions effective = (options ?? new OdfOpenOptions()).Normalize();
        var info = new FileInfo(fullPath);
        if (info.Length > effective.MaxPackageBytes) {
            throw new InvalidDataException($"OpenDocument package size {info.Length} exceeds MaxPackageBytes ({effective.MaxPackageBytes}).");
        }
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Open(stream, effective);
    }

    internal static OdfPackage Open(Stream stream, OdfOpenOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("OpenDocument stream must be readable.", nameof(stream));
        OdfOpenOptions effective = (options ?? new OdfOpenOptions()).Normalize();
        byte[] bytes = ReadAllBytesBounded(stream, effective.MaxPackageBytes);
        return Open(bytes, effective);
    }

    internal static OdfPackage Open(byte[] packageBytes, OdfOpenOptions? options = null) {
        if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
        OdfOpenOptions effective = (options ?? new OdfOpenOptions()).Normalize();
        if (packageBytes.LongLength > effective.MaxPackageBytes) {
            throw new InvalidDataException($"OpenDocument package size {packageBytes.LongLength} exceeds MaxPackageBytes ({effective.MaxPackageBytes}).");
        }
        OdfZipHeaderInspector.ValidateMimetypeEntry(packageBytes);

        var loaded = new List<OdfPackageEntry>();
        var exactNames = new HashSet<string>(StringComparer.Ordinal);
        var foldedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long totalUncompressed = 0;

        using (var stream = new MemoryStream(packageBytes, writable: false))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false)) {
            if (archive.Entries.Count > effective.MaxEntries) {
                throw new InvalidDataException($"OpenDocument package entry count {archive.Entries.Count} exceeds MaxEntries ({effective.MaxEntries}).");
            }

            foreach (ZipArchiveEntry archiveEntry in archive.Entries) {
                string normalized = OfficeArchiveSafety.NormalizeEntryName(archiveEntry.FullName);
                if (!string.Equals(normalized, archiveEntry.FullName, StringComparison.Ordinal) || OfficeArchiveSafety.IsUnsafePath(normalized)) {
                    throw new InvalidDataException($"OpenDocument package contains unsafe or non-canonical entry path '{archiveEntry.FullName}'.");
                }
                bool isDirectory = normalized.EndsWith("/", StringComparison.Ordinal);
                if (OfficeArchiveSafety.ComputeDepth(normalized, isDirectory) > effective.MaxDepth) {
                    throw new InvalidDataException($"OpenDocument entry '{normalized}' exceeds MaxDepth ({effective.MaxDepth}).");
                }
                if (!exactNames.Add(normalized) || !foldedNames.Add(normalized)) {
                    throw new InvalidDataException($"OpenDocument package contains duplicate or case-ambiguous entry '{normalized}'.");
                }
                if (!OfficeArchiveSafety.TryGetLength(archiveEntry, out long length)) {
                    throw new InvalidDataException($"OpenDocument entry '{normalized}' has unreadable length metadata.");
                }
                if (length > effective.MaxEntryUncompressedBytes) {
                    throw new InvalidDataException($"OpenDocument entry '{normalized}' exceeds MaxEntryUncompressedBytes ({effective.MaxEntryUncompressedBytes}).");
                }
                if (OfficeArchiveSafety.IsCompressionRatioExceeded(archiveEntry, length, effective.MaxCompressionRatio)) {
                    throw new InvalidDataException($"OpenDocument entry '{normalized}' exceeds MaxCompressionRatio ({effective.MaxCompressionRatio.ToString(CultureInfo.InvariantCulture)}).");
                }
                totalUncompressed = checked(totalUncompressed + length);
                if (totalUncompressed > effective.MaxTotalUncompressedBytes) {
                    throw new InvalidDataException($"OpenDocument package exceeds MaxTotalUncompressedBytes ({effective.MaxTotalUncompressedBytes}).");
                }

                byte[] data = isDirectory ? Array.Empty<byte>() : ReadEntryBytes(archiveEntry, length);
                DateTimeOffset timestamp;
                try { timestamp = archiveEntry.LastWriteTime; } catch { timestamp = DeterministicTimestamp; }
                loaded.Add(new OdfPackageEntry(normalized, data, null, timestamp, isNew: false));
            }
        }

        OdfPackageEntry mimetypeEntry = loaded.FirstOrDefault(entry => entry.Name == "mimetype")
            ?? throw new InvalidDataException("OpenDocument package is missing 'mimetype'.");
        string mediaType = Encoding.ASCII.GetString(mimetypeEntry.GetOriginalBytes());
        if (!OdfMediaTypes.TryGetKind(mediaType, out OdfDocumentKind kind)) {
            throw new InvalidDataException($"Unsupported OpenDocument media type '{mediaType}'.");
        }

        OdfPackageEntry manifestEntry = loaded.FirstOrDefault(entry => entry.Name == "META-INF/manifest.xml")
            ?? throw new InvalidDataException("OpenDocument package is missing 'META-INF/manifest.xml'.");
        XDocument manifest = manifestEntry.GetXml(effective.MaxXmlCharacters);
        XElement manifestRoot = manifest.Root ?? throw new InvalidDataException("OpenDocument manifest has no root element.");
        if (manifestRoot.Name != OdfNamespaces.Manifest + "manifest") {
            throw new InvalidDataException("OpenDocument manifest root must be 'manifest:manifest'.");
        }

        XElement? packageRootEntry = manifestRoot.Elements(OdfNamespaces.Manifest + "file-entry")
            .FirstOrDefault(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path") == "/");
        string? manifestMediaType = (string?)packageRootEntry?.Attribute(OdfNamespaces.Manifest + "media-type");
        if (!string.Equals(mediaType, manifestMediaType, StringComparison.Ordinal)) {
            throw new InvalidDataException("OpenDocument mimetype does not match the root manifest media type.");
        }

        string? versionToken = (string?)manifestRoot.Attribute(OdfNamespaces.Manifest + "version")
            ?? (string?)packageRootEntry?.Attribute(OdfNamespaces.Manifest + "version");
        if (!OdfVersionExtensions.TryParse(versionToken, out OdfVersion version)) {
            version = OdfVersion.V1_4;
        }

        var package = new OdfPackage(kind, version, effective);
        foreach (OdfPackageEntry entry in loaded) package.AddLoadedEntry(entry);
        package.ApplyManifestMediaTypes(manifestRoot);
        if (!package.ContainsEntry("content.xml")) {
            throw new InvalidDataException("OpenDocument package is missing 'content.xml'.");
        }
        if (manifestRoot.Descendants(OdfNamespaces.Manifest + "encryption-data").Any()) {
            throw new OdfEncryptedPackageException("Encrypted OpenDocument packages are detected but not yet supported for native editing.");
        }
        if (!OdfVersionExtensions.TryParse(versionToken, out _)) {
            package._diagnostics.Add(new OdfDiagnostic("ODF003", OdfDiagnosticSeverity.Warning,
                $"OpenDocument version '{versionToken ?? "<missing>"}' is not recognized; ODF 1.4 compatibility rules are used.", "META-INF/manifest.xml"));
        }
        return package;
    }

    internal bool ContainsEntry(string name) => _entriesByName.TryGetValue(name, out OdfPackageEntry? entry) && !entry.IsRemoved;

    internal OdfPackageEntry GetRequiredEntry(string name) {
        if (!_entriesByName.TryGetValue(name, out OdfPackageEntry? entry) || entry.IsRemoved) {
            throw new InvalidDataException($"OpenDocument package entry '{name}' is missing.");
        }
        return entry;
    }

    internal XDocument GetXml(string name) => GetRequiredEntry(name).GetXml(_openOptions.MaxXmlCharacters);

    internal XDocument EnsureXml(string name, XDocument template, string mediaType) {
        if (!ContainsEntry(name)) {
            AddOrReplaceEntry(name, OdfXmlCodec.Save(template), mediaType);
        }
        return GetXml(name);
    }

    internal void MarkXmlDirty(string name) => GetRequiredEntry(name).MarkDirty();

    internal void AddDiagnostic(OdfDiagnostic diagnostic) {
        if (diagnostic == null) throw new ArgumentNullException(nameof(diagnostic));
        _diagnostics.Add(diagnostic);
    }

    internal void AddOrReplaceEntry(string name, byte[] data, string mediaType) {
        ValidateNewEntryName(name);
        if (_entriesByName.TryGetValue(name, out OdfPackageEntry? existing)) {
            existing.ReplaceBytes(data, mediaType);
        } else {
            var entry = new OdfPackageEntry(name, data, mediaType, DeterministicTimestamp, isNew: true);
            _entries.Add(entry);
            _entriesByName.Add(name, entry);
        }
        _entryGraphChanged = true;
    }

    internal void RemoveEntry(string name) {
        if (_entriesByName.TryGetValue(name, out OdfPackageEntry? entry) && !entry.IsRemoved) {
            entry.Remove();
            _entryGraphChanged = true;
        }
    }

    internal byte[] Write(OdfSaveOptions? options = null) {
        OdfSaveOptions effective = options ?? new OdfSaveOptions();
        bool hasChanges = _entryGraphChanged || _entries.Any(entry => entry.IsDirty);
        if (IsSigned && hasChanges) {
            if (effective.SignatureHandling == OdfSignatureHandling.RejectInvalidation) {
                throw new InvalidOperationException("Saving this changed document would invalidate its signatures. Set SignatureHandling to RemoveInvalidated to continue.");
            }
            foreach (OdfPackageEntry signature in _entries.Where(entry => IsSignaturePath(entry.Name)).ToList()) {
                signature.Remove();
            }
            _entryGraphChanged = true;
        }

        OdfVersion outputVersion = ResolveOutputVersion(effective.CompatibilityProfile);
        if (outputVersion != Version) {
            UpdateXmlVersions(outputVersion);
            _entryGraphChanged = true;
        }
        if (_entryGraphChanged) {
            RebuildManifest(outputVersion);
        }

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            OdfPackageEntry mimetype = GetRequiredEntry("mimetype");
            WriteEntry(archive, mimetype, CompressionLevel.NoCompression, effective.Deterministic);

            IEnumerable<OdfPackageEntry> remaining = _entries.Where(entry => !entry.IsRemoved && entry.Name != "mimetype");
            if (effective.Deterministic) {
                OdfPackageEntry[] original = remaining.Where(entry => !entry.IsNew).ToArray();
                OdfPackageEntry[] added = remaining.Where(entry => entry.IsNew).OrderBy(entry => entry.Name, StringComparer.Ordinal).ToArray();
                remaining = original.Concat(added);
            }
            foreach (OdfPackageEntry entry in remaining) {
                CompressionLevel level = entry.Name.EndsWith("/", StringComparison.Ordinal) ? CompressionLevel.NoCompression : CompressionLevel.Optimal;
                WriteEntry(archive, entry, level, effective.Deterministic);
            }
        }
        Version = outputVersion;
        return output.ToArray();
    }

    internal OdfSaveReport CreateSaveReport() {
        return new OdfSaveReport(
            _entries.Where(entry => !entry.IsRemoved && entry.IsDirty).Select(entry => entry.Name).ToArray(),
            _entries.Where(entry => !entry.IsRemoved && !entry.IsDirty).Select(entry => entry.Name).ToArray(),
            _entries.Where(entry => entry.IsRemoved).Select(entry => entry.Name).ToArray());
    }

    private void AddInitialEntry(string name, byte[] data, string? mediaType) {
        var entry = new OdfPackageEntry(name, data, mediaType, DeterministicTimestamp, isNew: true);
        _entries.Add(entry);
        _entriesByName.Add(name, entry);
    }

    private void AddInitialXml(string name, XDocument document, string mediaType) {
        AddInitialEntry(name, OdfXmlCodec.Save(document), mediaType);
        GetRequiredEntry(name).GetXml(_openOptions.MaxXmlCharacters);
    }

    private void AddLoadedEntry(OdfPackageEntry entry) {
        _entries.Add(entry);
        _entriesByName.Add(entry.Name, entry);
    }

    private void ApplyManifestMediaTypes(XElement manifestRoot) {
        foreach (XElement fileEntry in manifestRoot.Elements(OdfNamespaces.Manifest + "file-entry")) {
            string? path = (string?)fileEntry.Attribute(OdfNamespaces.Manifest + "full-path");
            if (string.IsNullOrEmpty(path) || path == "/") continue;
            if (_entriesByName.TryGetValue(path!, out OdfPackageEntry? entry)) {
                entry.MediaType = (string?)fileEntry.Attribute(OdfNamespaces.Manifest + "media-type");
            }
        }
    }

    private void RebuildManifest(OdfVersion outputVersion) {
        OdfPackageEntry manifestEntry = GetRequiredEntry("META-INF/manifest.xml");
        XDocument manifest = manifestEntry.GetXml(_openOptions.MaxXmlCharacters);
        XElement root = manifest.Root ?? throw new InvalidDataException("OpenDocument manifest has no root element.");
        root.SetAttributeValue(OdfNamespaces.Manifest + "version", outputVersion.ToToken());

        List<XElement> fileEntries = root.Elements(OdfNamespaces.Manifest + "file-entry").ToList();
        XElement? rootEntry = fileEntries.FirstOrDefault(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path") == "/");
        if (rootEntry == null) {
            rootEntry = OdfPackageTemplates.FileEntry("/", MediaType, outputVersion.ToToken());
            root.AddFirst(rootEntry);
        }
        rootEntry.SetAttributeValue(OdfNamespaces.Manifest + "media-type", MediaType);
        rootEntry.SetAttributeValue(OdfNamespaces.Manifest + "version", outputVersion.ToToken());

        var actualPaths = new HashSet<string>(_entries.Where(entry => !entry.IsRemoved).Select(entry => entry.Name), StringComparer.Ordinal);
        foreach (XElement fileEntry in fileEntries) {
            string? path = (string?)fileEntry.Attribute(OdfNamespaces.Manifest + "full-path");
            if (string.IsNullOrEmpty(path) || path == "/") continue;
            if (path == "mimetype" || path == "META-INF/manifest.xml" || !actualPaths.Contains(path!)) {
                fileEntry.Remove();
            }
        }

        var listed = new HashSet<string>(root.Elements(OdfNamespaces.Manifest + "file-entry")
            .Select(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path"))
            .Where(path => !string.IsNullOrEmpty(path))
            .Select(path => path!), StringComparer.Ordinal);
        foreach (OdfPackageEntry entry in _entries.Where(entry => !entry.IsRemoved && entry.Name != "mimetype" && entry.Name != "META-INF/manifest.xml")) {
            if (entry.Name.StartsWith("META-INF/", StringComparison.Ordinal)) continue;
            if (listed.Add(entry.Name)) {
                root.Add(OdfPackageTemplates.FileEntry(entry.Name, entry.MediaType ?? GuessMediaType(entry.Name), null));
            } else {
                XElement existing = root.Elements(OdfNamespaces.Manifest + "file-entry")
                    .First(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path") == entry.Name);
                if (!string.IsNullOrEmpty(entry.MediaType)) {
                    existing.SetAttributeValue(OdfNamespaces.Manifest + "media-type", entry.MediaType);
                }
            }
        }
        manifestEntry.MarkDirty();
    }

    private void UpdateXmlVersions(OdfVersion outputVersion) {
        foreach (string path in new[] { "content.xml", "styles.xml", "meta.xml", "settings.xml" }) {
            if (!ContainsEntry(path)) continue;
            XDocument xml = GetXml(path);
            if (xml.Root != null) {
                xml.Root.SetAttributeValue(OdfNamespaces.Office + "version", outputVersion.ToToken());
                MarkXmlDirty(path);
            }
        }
    }

    private OdfVersion ResolveOutputVersion(OdfCompatibilityProfile profile) {
        switch (profile) {
            case OdfCompatibilityProfile.Odf13: return OdfVersion.V1_3;
            case OdfCompatibilityProfile.PreserveSource: return Version;
            default: return OdfVersion.V1_4;
        }
    }

    private static void WriteEntry(ZipArchive archive, OdfPackageEntry entry, CompressionLevel compressionLevel, bool deterministic) {
        ZipArchiveEntry outputEntry = archive.CreateEntry(entry.Name, compressionLevel);
        if (deterministic) outputEntry.LastWriteTime = DeterministicTimestamp;
        byte[] data = entry.GetBytesForSave();
        using Stream destination = outputEntry.Open();
        destination.Write(data, 0, data.Length);
    }

    private static byte[] ReadAllBytesBounded(Stream stream, long maxBytes) {
        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            if (remaining > maxBytes) throw new InvalidDataException($"OpenDocument stream exceeds MaxPackageBytes ({maxBytes}).");
        }
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            total += read;
            if (total > maxBytes) throw new InvalidDataException($"OpenDocument stream exceeds MaxPackageBytes ({maxBytes}).");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static byte[] ReadEntryBytes(ZipArchiveEntry entry, long length) {
        if (length > int.MaxValue) throw new InvalidDataException($"OpenDocument entry '{entry.FullName}' is too large for the in-memory package store.");
        using Stream source = entry.Open();
        using var output = new MemoryStream(length > 0 ? (int)length : 0);
        source.CopyTo(output);
        if (output.Length != length) throw new InvalidDataException($"OpenDocument entry '{entry.FullName}' length changed while reading.");
        return output.ToArray();
    }

    private static bool IsSignaturePath(string path) {
        return path.StartsWith("META-INF/", StringComparison.Ordinal) &&
            path.EndsWith("signatures.xml", StringComparison.OrdinalIgnoreCase);
    }

    private static string GuessMediaType(string path) {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        switch (extension) {
            case ".xml": return "text/xml";
            case ".png": return "image/png";
            case ".jpg":
            case ".jpeg": return "image/jpeg";
            case ".gif": return "image/gif";
            case ".svg": return "image/svg+xml";
            default: return string.Empty;
        }
    }

    private static void ValidateNewEntryName(string name) {
        string normalized = OfficeArchiveSafety.NormalizeEntryName(name);
        if (!string.Equals(name, normalized, StringComparison.Ordinal) || OfficeArchiveSafety.IsUnsafePath(normalized)) {
            throw new ArgumentException("Package entry name must be a safe, canonical relative path.", nameof(name));
        }
    }
}
