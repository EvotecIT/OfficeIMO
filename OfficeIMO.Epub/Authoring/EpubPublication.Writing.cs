using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>Produces a complete bounded ZIP before publication. Unedited imports return their exact original bytes.</summary>
    public EpubWriteResult Write(EpubWriteOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new EpubWriteOptions();
        long maxOutput = options.MaxOutputBytes, maxExpanded = options.MaxExpandedBytes;
        int maxEntries = options.MaxEntries;
        bool removeSignatures = options.RemoveInvalidatedSignatures;
        bool compress = options.CompressEntries;
        DateTimeOffset? modifiedAt = options.ModifiedAt;
        if (maxOutput < 1 || maxExpanded < 1 || maxEntries < 1) throw new ArgumentOutOfRangeException(nameof(options), "Output bounds must be positive.");
        cancellationToken.ThrowIfCancellationRequested();
        var entries = new Dictionary<string, byte[]>(_entries, StringComparer.Ordinal);
        bool changed = _changed || modifiedAt.HasValue;
        bool rewriteArchive = changed || !compress;
        var diagnostics = new List<OfficeConversionFidelityDiagnostic>();
        if (rewriteArchive && _originalHasZipSignature) {
            if (!removeSignatures) throw new InvalidOperationException("Edits invalidate the ZIP directory signature. Explicit signature removal is required.");
            diagnostics.Add(new OfficeConversionFidelityDiagnostic("EPUB_WRITE_ZIP_SIGNATURE_REMOVED",
                "Invalidated ZIP central-directory signature was removed by explicit policy.", OfficeConversionLossKind.Omission, "OfficeIMO.Epub"));
        }
        if (changed && _originalEntries.ContainsKey("META-INF/signatures.xml")) {
            if (!removeSignatures) throw new InvalidOperationException("Edits invalidate package signatures. Explicit signature removal is required.");
            entries.Remove("META-INF/signatures.xml");
            diagnostics.Add(new OfficeConversionFidelityDiagnostic("EPUB_WRITE_SIGNATURE_REMOVED",
                "Invalidated META-INF/signatures.xml was removed by explicit policy.", OfficeConversionLossKind.Omission, "OfficeIMO.Epub", "META-INF/signatures.xml"));
        }
        if (changed && _encryption.Any(item => item.RequiresDecryption)) throw new NotSupportedException("Editing packages with unsupported resource encryption is unavailable.");
        if (_encryption.Count != 0 && Identifier != _originalIdentifier) throw new NotSupportedException("Encrypted/obfuscated package identity must remain unchanged.");
        if (changed && PackageVersion == "2.0" && Identifier != _originalIdentifier) {
            string path = NavigationPath();
            XDocument ncx = ParseXml(entries[path], _maximumEntryBytes);
            XElement head = ncx.Root?.Element(Ncx + "head") ?? throw new InvalidDataException("NCX has no metadata head.");
            XElement? uid = head.Elements(Ncx + "meta").FirstOrDefault(meta => (string?)meta.Attribute("name") == "dtb:uid");
            if (uid == null) head.Add(new XElement(Ncx + "meta", new XAttribute("name", "dtb:uid"), new XAttribute("content", Identifier)));
            else uid.SetAttributeValue("content", Identifier);
            entries[path] = SerializeXml(ncx, _maximumEntryBytes);
        }
        XDocument package = new XDocument(_package);
        if (changed && PackageVersion == "3.0") {
            XElement metadata = package.Root!.Element(Opf + "metadata")!;
            if (EpubVocabulary.Expand(package.Root, "dcterms:modified") != "http://purl.org/dc/terms/modified")
                throw new InvalidDataException("The dcterms prefix cannot be reassigned when writing modification metadata.");
            XElement[] stamps = metadata.Elements(Opf + "meta").Where(element =>
                EpubVocabulary.Expand(package.Root, (string?)element.Attribute("property") ?? string.Empty) == "http://purl.org/dc/terms/modified" && element.Attribute("refines") == null).ToArray();
            if (stamps.Length > 1) throw new InvalidDataException("Package contains multiple unrefined modification stamps.");
            string value = (modifiedAt ?? _modifiedAt).UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
            if (stamps.Length == 0) metadata.Add(new XElement(Opf + "meta", new XAttribute("property", "dcterms:modified"), value));
            else stamps[0].Value = value;
        }
        if (changed || !entries.ContainsKey(PackagePath)) entries[PackagePath] = SerializeXml(package, Math.Min(_maximumMetadataBytes, _maximumEntryBytes));
        if (!entries.TryGetValue("mimetype", out byte[]? mimetype) || !mimetype.SequenceEqual(Encoding.ASCII.GetBytes("application/epub+zip")))
            throw new InvalidDataException("Package mimetype must contain exactly application/epub+zip.");
        int outputEntryCount = !rewriteArchive && _originalBytes != null ? _originalEntryCount : entries.Count;
        if (outputEntryCount > Math.Min(maxEntries, _maximumEntries)) throw new InvalidDataException("Output exceeds MaxEntries or the retained entry-count limit.");
        ValidatePublication(package, entries, diagnostics, cancellationToken, changed);
        if (changed) {
            EnsurePackageBudget(package, entries.Where(entry => entry.Key != PackagePath).Sum(entry => entry.Value.LongLength) - RetainedPayloadBytes);
            entries[PackagePath] = SerializeXml(package, Math.Min(_maximumMetadataBytes, _maximumEntryBytes));
        }
        long expanded = 0;
        foreach (var entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Value.LongLength > _maximumEntryBytes) throw new InvalidDataException("Output resource exceeds the retained entry-byte limit.");
            if (entry.Value.LongLength > maxExpanded - expanded) throw new InvalidDataException("Output exceeds MaxExpandedBytes.");
            expanded += entry.Value.LongLength;
        }
        if (!rewriteArchive && _originalBytes != null) {
            if (_originalExpandedBytes > maxExpanded) throw new InvalidDataException("Output exceeds MaxExpandedBytes.");
            if (_originalBytes.LongLength > maxOutput) throw new InvalidDataException("Output exceeds MaxOutputBytes.");
            return new EpubWriteResult((byte[])_originalBytes.Clone(), new EpubWriteReport(true,
                entries.Keys.OrderBy(path => path, StringComparer.Ordinal), Array.Empty<string>(), Array.Empty<string>(), diagnostics));
        }
        OfficeProvenanceZipWriteEntry[] writeEntries = new[] { "mimetype" }
            .Concat(entries.Keys.Where(path => path != "mimetype").OrderBy(path => path, StringComparer.Ordinal))
            .Select(path => new OfficeProvenanceZipWriteEntry(path, entries[path].LongLength, path != "mimetype" && compress,
                new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero), 0, 0,
                Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<byte>(), () => new MemoryStream(entries[path], false))).ToArray();
        byte[] output = OfficeProvenanceZipWriter.Write(writeEntries, maxExpanded, maximumOutputBytes: maxOutput, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        string[] preserved = entries.Where(pair => _originalEntries.TryGetValue(pair.Key, out byte[]? original) && original.SequenceEqual(pair.Value))
            .Select(pair => pair.Key).OrderBy(path => path, StringComparer.Ordinal).ToArray();
        string[] regenerated = entries.Keys.Except(preserved, StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal).ToArray();
        string[] removed = _originalEntries.Keys.Except(entries.Keys, StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal).ToArray();
        foreach (string path in removed.Where(path => path != "META-INF/signatures.xml")) diagnostics.Add(new OfficeConversionFidelityDiagnostic(
            "EPUB_WRITE_ENTRY_REMOVED", "Original entry was explicitly removed by editing.", OfficeConversionLossKind.Omission, "OfficeIMO.Epub", path));
        return new EpubWriteResult(output, new EpubWriteReport(false, preserved, regenerated, removed, diagnostics));
    }

    /// <summary>Atomically saves a completed artifact to a file, preserving an existing destination on validation/cancellation failure.</summary>
    public EpubWriteReport Save(string path, EpubWriteOptions? options = null, CancellationToken cancellationToken = default) {
        EpubWriteResult result = Write(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        string temporary = OfficeFileCommit.StageAllBytes(path, result.Bytes);
        try {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeFileCommit.CommitTemporaryFileAtomically(temporary, path);
        } finally { OfficeFileCommit.DeleteIfExists(temporary); }
        return result.Report;
    }
    /// <summary>Writes a staged complete artifact to a caller-owned stream. Seekable streams are replaced and rewound.</summary>
    public EpubWriteReport Save(Stream stream, EpubWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        EpubWriteResult result = Write(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeStreamWriter.Write(stream, destination => WritePayload(destination, result.Bytes, cancellationToken));
        return result.Report;
    }
    /// <summary>Asynchronously atomically saves a completed artifact to a file.</summary>
    public async Task<EpubWriteReport> SaveAsync(string path, EpubWriteOptions? options = null, CancellationToken cancellationToken = default) {
        EpubWriteResult result = Write(options, cancellationToken);
        string temporary = await OfficeFileCommit.StageAllBytesAsync(path, result.Bytes, cancellationToken).ConfigureAwait(false);
        try {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeFileCommit.CommitTemporaryFileAtomically(temporary, path);
        } finally { OfficeFileCommit.DeleteIfExists(temporary); }
        return result.Report;
    }
    /// <summary>Asynchronously writes a staged artifact without closing the destination stream.</summary>
    public async Task<EpubWriteReport> SaveAsync(Stream stream, EpubWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        EpubWriteResult result = Write(options, cancellationToken);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, result.Bytes, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }
    private static void WritePayload(Stream destination, byte[] bytes, CancellationToken token) {
        for (int offset = 0; offset < bytes.Length; offset += 81920) {
            token.ThrowIfCancellationRequested();
            destination.Write(bytes, offset, Math.Min(81920, bytes.Length - offset));
        }
        token.ThrowIfCancellationRequested();
    }
}
