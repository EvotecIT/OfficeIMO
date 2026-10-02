using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using System.Threading;

namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    /// <summary>Reuses package discovery and archive identity checks for loss-aware editing.</summary>
    internal static (string OpfPath, Dictionary<string, byte[]> Entries, IReadOnlyList<EpubEncryptionInfo> Encryption)
        ReadEditablePackage(byte[] bytes, EpubPublicationLoadOptions options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        OfficeProvenanceZip.ValidateMimetypeEntry(bytes, "application/epub+zip", options.MaxEntries);
        token.ThrowIfCancellationRequested();
        var readOptions = new EpubReadOptions {
            MaxPackageBytes = options.MaxInputBytes, MaxArchiveEntries = options.MaxEntries,
            MaxTotalUncompressedBytes = options.MaxExpandedBytes, MaxPackageMetadataBytes = options.MaxMetadataBytes
        };
        using var input = new MemoryStream(bytes, false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        OfficeProvenanceZip.GetValidatedEntryNames(bytes, archive);
        var diagnostics = new EpubDiagnosticCollector();
        Dictionary<string, ZipArchiveEntry> index = BuildEntryIndex(archive, readOptions, diagnostics, token);
        if (diagnostics.Items.Count != 0) throw new InvalidDataException("Editing requires unique, safe archive entry paths.");
        foreach (var pair in index) {
            if (pair.Value.FullName != pair.Key) throw new InvalidDataException("Editing requires canonical archive paths.");
        }
        EpubPackage? package = TryReadPackage(index, readOptions, diagnostics, out IReadOnlyList<EpubRootfile> rootfiles, token);
        if (package == null || !rootfiles.Any(rootfile => rootfile.IsSelected)) {
            throw new InvalidDataException("Editing requires a readable, declared package document.");
        }
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        long retained = 0;
        foreach (var pair in index) {
            token.ThrowIfCancellationRequested();
            if (pair.Value.Length > options.MaxEntryBytes) throw new InvalidDataException("Entry exceeds MaxEntryBytes: " + pair.Key);
            using Stream entry = pair.Value.Open();
            byte[] payload = OfficeStreamReader.ReadAllBytes(entry, token, Math.Max(1, Math.Min(options.MaxEntryBytes, options.MaxExpandedBytes - retained)));
            if (payload.LongLength > options.MaxExpandedBytes - retained) throw new InvalidDataException("Package exceeds MaxExpandedBytes.");
            retained += payload.LongLength;
            entries.Add(pair.Key, payload);
        }
        return (package.OpfPath, entries, ReadEditableEncryption(index, readOptions, token));
    }

    // Inspection may skip unreadable declarations. Editing must identify every protected payload.
    private static IReadOnlyList<EpubEncryptionInfo> ReadEditableEncryption(
        IReadOnlyDictionary<string, ZipArchiveEntry> index, EpubReadOptions options, CancellationToken token) {
        if (!index.TryGetValue("META-INF/encryption.xml", out ZipArchiveEntry? entry)) return Array.Empty<EpubEncryptionInfo>();
        if (entry.Length > options.MaxPackageMetadataBytes ||
            !TryParseEntryXml(entry, options.MaxPackageMetadataBytes, token, out XDocument? document) ||
            document?.Root?.Name != XName.Get("encryption", ContainerNamespaceUri))
            throw new InvalidDataException("Editing requires a readable encryption.xml within MaxMetadataBytes.");
        XElement root = document.Root;
        XElement[] declarations = root.Elements(XName.Get("EncryptedData", XmlEncryptionNamespaceUri)).ToArray();
        if (declarations.Length == 0 || root.Elements().Any(element => !IsXmlEncryptionName(element, "EncryptedData") && !IsXmlEncryptionName(element, "EncryptedKey")) ||
            root.Descendants(XName.Get("EncryptedData", XmlEncryptionNamespaceUri)).Count() != declarations.Length)
            throw new InvalidDataException("Editing requires unambiguous encryption declarations.");
        var results = new List<EpubEncryptionInfo>();
        var paths = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement declaration in declarations) {
            token.ThrowIfCancellationRequested();
            XElement[] methods = declaration.Elements(XName.Get("EncryptionMethod", XmlEncryptionNamespaceUri)).ToArray();
            XElement[] ciphers = declaration.Elements(XName.Get("CipherData", XmlEncryptionNamespaceUri)).ToArray();
            if (methods.Length != 1 || ciphers.Length != 1 || ciphers[0].Elements().Count() != 1 ||
                ciphers[0].Elements().Single().Name != XName.Get("CipherReference", XmlEncryptionNamespaceUri))
                throw new InvalidDataException("Editing requires one encryption algorithm and resource reference per declaration.");
            XElement reference = ciphers[0].Elements().Single();
            EpubReference target = EpubReference.Resolve("package.opf", GetUnqualifiedAttribute(reference, "URI"));
            if (reference.HasElements || target.Kind != EpubReferenceKind.Container || target.ContainerPath == null ||
                target.Query != null || target.Fragment != null || !index.ContainsKey(target.ContainerPath) || !paths.Add(target.ContainerPath))
                throw new InvalidDataException("Editing requires distinct, existing container encryption targets without transforms.");
            string? algorithm = NullIfWhiteSpace(GetUnqualifiedAttribute(methods[0], "Algorithm"));
            results.Add(new EpubEncryptionInfo { Path = target.ContainerPath, Algorithm = algorithm, Kind = ClassifyEncryption(algorithm) });
        }
        return results;
    }
}
