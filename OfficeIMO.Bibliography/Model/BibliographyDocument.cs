namespace OfficeIMO.Bibliography;

/// <summary>An editable, format-neutral, source-backed bibliography.</summary>
public sealed partial class BibliographyDocument {
    private readonly string? _originalText;
    private readonly byte[]? _originalBytes;
    private readonly string _baselineFingerprint;

    /// <summary>Creates an empty bibliography for a destination format.</summary>
    public BibliographyDocument(BibliographyFormat format) {
        ValidateFormat(format);
        SourceFormat = format;
        _baselineFingerprint = BibliographyFingerprint.Create(this);
    }

    internal BibliographyDocument(BibliographyFormat format, IList<BibliographyItem> items, IList<BibliographyNativeEntry> nativeEntries, string originalText, byte[]? originalBytes, IReadOnlyList<BibliographyDiagnostic> diagnostics, bool cslJsonSingleObjectRoot = false, bool endNoteRecordsRoot = false) {
        ValidateFormat(format);
        SourceFormat = format;
        Items = items ?? throw new ArgumentNullException(nameof(items));
        NativeEntries = nativeEntries ?? throw new ArgumentNullException(nameof(nativeEntries));
        _originalText = originalText ?? throw new ArgumentNullException(nameof(originalText));
        _originalBytes = originalBytes;
        Diagnostics = diagnostics ?? Array.Empty<BibliographyDiagnostic>();
        CslJsonSingleObjectRoot = cslJsonSingleObjectRoot;
        EndNoteRecordsRoot = endNoteRecordsRoot;
        _baselineFingerprint = BibliographyFingerprint.Create(this);
    }

    /// <summary>Source format, or intended format for a new document.</summary>
    public BibliographyFormat SourceFormat { get; }
    /// <summary>Editable citation records in source order.</summary>
    public IList<BibliographyItem> Items { get; private set; } = new List<BibliographyItem>();
    /// <summary>Parser and recovery diagnostics.</summary>
    public IReadOnlyList<BibliographyDiagnostic> Diagnostics { get; private set; } = Array.Empty<BibliographyDiagnostic>();
    /// <summary>Document-level native directives retained in source order.</summary>
    public IList<BibliographyNativeEntry> NativeEntries { get; private set; } = new List<BibliographyNativeEntry>();
    /// <summary>True when the semantic model differs from the parsed baseline.</summary>
    public bool IsModified => !string.Equals(_baselineFingerprint, BibliographyFingerprint.Create(this), StringComparison.Ordinal);
    /// <summary>True when exact original text is available.</summary>
    public bool HasOriginalSource => _originalText != null;

    /// <summary>Original decoded source text, when the document was parsed or loaded.</summary>
    public string? OriginalSourceText => _originalText;

    /// <summary>Returns a copy of the original loaded bytes, or null when the document was parsed from text or created in memory.</summary>
    public byte[]? GetOriginalBytes() => _originalBytes == null ? null : (byte[])_originalBytes.Clone();

    internal string? OriginalText => _originalText;
    internal byte[]? OriginalBytes => _originalBytes;
    internal bool CslJsonSingleObjectRoot { get; }
    internal bool EndNoteRecordsRoot { get; }

    /// <summary>Parses bibliography text using an explicit format.</summary>
    public static BibliographyReadResult Parse(string source, BibliographyFormat format, BibliographyReadOptions? options = null, CancellationToken cancellationToken = default) =>
        BibliographyReader.Parse(source, format, options, null, cancellationToken);

    /// <summary>Parses bibliography text after bounded content detection.</summary>
    public static BibliographyReadResult Parse(string source, BibliographyReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return Parse(source, BibliographyFormatDetector.Detect(source, options), options, cancellationToken);
    }

    /// <summary>Writes or converts the document.</summary>
    public BibliographyWriteResult Write(BibliographyWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        BibliographyWriter.Write(this, options, cancellationToken);

    /// <summary>Writes or converts the document to bytes.</summary>
    public byte[] ToBytes(BibliographyWriteOptions? options = null, CancellationToken cancellationToken = default) => Write(options, cancellationToken).Bytes;

    /// <summary>Writes or converts the document to a new memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(BibliographyWriteOptions? options = null, CancellationToken cancellationToken = default) => new MemoryStream(ToBytes(options, cancellationToken));

    /// <summary>Writes text in preserve mode using the source format.</summary>
    public override string ToString() => Write().Content;

    private static void ValidateFormat(BibliographyFormat format) {
        if (!Enum.IsDefined(typeof(BibliographyFormat), format)) throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown bibliography format.");
    }
}

internal static class BibliographyFingerprint {
    internal static string Create(BibliographyDocument document) {
        var builder = new StringBuilder();
        Add(builder, document.Items.Count.ToString(CultureInfo.InvariantCulture));
        foreach (BibliographyItem item in document.Items) {
            Add(builder, item.Key); Add(builder, ((int)item.Type).ToString(CultureInfo.InvariantCulture)); Add(builder, item.NativeType);
            Add(builder, item.Title); Add(builder, item.ContainerTitle); Add(builder, item.CollectionTitle); Add(builder, item.Publisher);
            Add(builder, item.PublisherPlace); Add(builder, item.Edition); Add(builder, item.Volume); Add(builder, item.Issue);
            Add(builder, item.Pages); Add(builder, item.Abstract); Add(builder, item.Language); Add(builder, item.Url);
            Add(builder, item.Contributors.Count.ToString(CultureInfo.InvariantCulture));
            foreach (BibliographyContributor contributor in item.Contributors) {
                Add(builder, ((int)contributor.Role).ToString(CultureInfo.InvariantCulture)); Add(builder, contributor.Name.Given);
                Add(builder, contributor.Name.Family); Add(builder, contributor.Name.Literal); Add(builder, contributor.Name.Suffix);
                Add(builder, contributor.Name.DroppingParticle); Add(builder, contributor.Name.NonDroppingParticle);
                Add(builder, contributor.Name.NativeFields.Count.ToString(CultureInfo.InvariantCulture));
                foreach (BibliographyNativeField field in contributor.Name.NativeFields) { Add(builder, ((int)field.Format).ToString(CultureInfo.InvariantCulture)); Add(builder, field.Name); Add(builder, field.Value); Add(builder, field.RawValue); }
            }
            Add(builder, item.Dates.Count.ToString(CultureInfo.InvariantCulture));
            foreach (BibliographyDate date in item.Dates) {
                Add(builder, ((int)date.Role).ToString(CultureInfo.InvariantCulture)); Add(builder, date.Year?.ToString(CultureInfo.InvariantCulture));
                Add(builder, date.Month?.ToString(CultureInfo.InvariantCulture)); Add(builder, date.Day?.ToString(CultureInfo.InvariantCulture)); Add(builder, date.Literal);
                Add(builder, date.EndYear?.ToString(CultureInfo.InvariantCulture)); Add(builder, date.EndMonth?.ToString(CultureInfo.InvariantCulture)); Add(builder, date.EndDay?.ToString(CultureInfo.InvariantCulture));
                Add(builder, date.NativeFields.Count.ToString(CultureInfo.InvariantCulture));
                foreach (BibliographyNativeField field in date.NativeFields) { Add(builder, ((int)field.Format).ToString(CultureInfo.InvariantCulture)); Add(builder, field.Name); Add(builder, field.Value); Add(builder, field.RawValue); }
            }
            Add(builder, item.Identifiers.Count.ToString(CultureInfo.InvariantCulture));
            foreach (BibliographyIdentifier identifier in item.Identifiers) { Add(builder, identifier.Scheme); Add(builder, identifier.Value); }
            Add(builder, item.Keywords.Count.ToString(CultureInfo.InvariantCulture));
            foreach (string keyword in item.Keywords) Add(builder, keyword);
            Add(builder, item.Notes.Count.ToString(CultureInfo.InvariantCulture));
            foreach (string note in item.Notes) Add(builder, note);
            Add(builder, item.NativeFields.Count.ToString(CultureInfo.InvariantCulture));
            foreach (BibliographyNativeField field in item.NativeFields) {
                Add(builder, ((int)field.Format).ToString(CultureInfo.InvariantCulture)); Add(builder, field.Name); Add(builder, field.Value); Add(builder, field.RawValue);
            }
        }
        Add(builder, document.NativeEntries.Count.ToString(CultureInfo.InvariantCulture));
        foreach (BibliographyNativeEntry entry in document.NativeEntries) {
            Add(builder, ((int)entry.Format).ToString(CultureInfo.InvariantCulture)); Add(builder, entry.Kind); Add(builder, entry.Name); Add(builder, entry.Value);
        }
        return builder.ToString();
    }

    private static void Add(StringBuilder builder, string? value) {
        if (value == null) { builder.Append("-1:;"); return; }
        builder.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value).Append(';');
    }
}
