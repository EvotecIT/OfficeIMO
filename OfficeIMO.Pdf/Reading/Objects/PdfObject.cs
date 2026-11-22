using System.Runtime.InteropServices;

namespace OfficeIMO.Pdf;

internal abstract class PdfObject {
    // Retain lenient read compatibility while exposing discarded or incomplete syntax
    // to mutation preflight and conservative feature detection.
    internal bool HasIncompleteSyntax { get; set; }
}

/// <summary>PDF numeric value (integer or real).</summary>
internal sealed class PdfNumber : PdfObject {
    public double Value { get; }
    public PdfNumber(double value) { Value = value; }
    public override string ToString() => Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
}

/// <summary>PDF boolean value.</summary>
internal sealed class PdfBoolean : PdfObject {
    public bool Value { get; }
    public PdfBoolean(bool value) { Value = value; }
    public override string ToString() => Value ? "true" : "false";
}

/// <summary>PDF name object (e.g. /Type, /Font).</summary>
internal sealed class PdfName : PdfObject {
    public string Name { get; }
    public PdfName(string name) { Name = name; }
    public override string ToString() => "/" + Name;
}

/// <summary>PDF literal string object (..).</summary>
internal sealed class PdfStringObj : PdfObject {
    public string Value { get; }
    public bool UseTextStringEncoding { get; }
    public byte[] RawBytes { get; }
    /// <summary>Byte-constructed and parsed strings retain their payload without re-encoding the decoded view.</summary>
    internal bool PreserveRawBytes { get; }
    internal int? EncodedTokenLength { get; }

    public PdfStringObj(string value, bool useTextStringEncoding = false) {
        Value = value;
        UseTextStringEncoding = useTextStringEncoding;
        RawBytes = useTextStringEncoding
            ? PdfTextString.Encode(value)
            : PdfEncoding.Latin1GetBytes(value);
    }

    public PdfStringObj(
        byte[] rawBytes,
        bool useTextStringEncoding = false,
        int? encodedTokenLength = null) {
        RawBytes = (byte[])rawBytes.Clone();
        PreserveRawBytes = true;
        Value = PdfTextString.Decode(rawBytes);
        UseTextStringEncoding = useTextStringEncoding;
        EncodedTokenLength = encodedTokenLength;
    }

    /// <summary>Adopts bytes allocated exclusively for a parsed string token.</summary>
    internal static PdfStringObj FromParsedBytes(
        byte[] ownedRawBytes,
        string decodedValue,
        bool useTextStringEncoding,
        int? encodedTokenLength) => new PdfStringObj(
            ownedRawBytes,
            decodedValue,
            useTextStringEncoding,
            encodedTokenLength);

    private PdfStringObj(
        byte[] ownedRawBytes,
        string decodedValue,
        bool useTextStringEncoding,
        int? encodedTokenLength) {
        RawBytes = ownedRawBytes;
        PreserveRawBytes = true;
        Value = decodedValue;
        UseTextStringEncoding = useTextStringEncoding;
        EncodedTokenLength = encodedTokenLength;
    }

    public override string ToString() => Value;
}

/// <summary>PDF array object.</summary>
internal sealed class PdfArray : PdfObject {
    public PdfArray() { }
    internal PdfArray(int capacity) {
        if (capacity > 0) Items.Capacity = capacity;
    }

    public System.Collections.Generic.List<PdfObject> Items { get; } = new();
}

/// <summary>PDF null object.</summary>
internal sealed class PdfNull : PdfObject {
    private PdfNull() { }
    public static PdfNull Instance { get; } = new PdfNull();
    public override string ToString() => "null";
}

/// <summary>PDF dictionary object.</summary>
internal sealed class PdfDictionary : PdfObject {
    public PdfDictionary() => Items = new();
    internal PdfDictionary(int capacity) => Items = new(capacity > 0 ? capacity : 0);

    public System.Collections.Generic.Dictionary<string, PdfObject> Items { get; }
    public T? Get<T>(string key) where T : PdfObject => Items.TryGetValue(key, out var v) ? v as T : null;
}

/// <summary>PDF indirect reference (e.g. 5 0 R).</summary>
internal sealed class PdfReference : PdfObject {
    public int ObjectNumber { get; }
    public int Generation { get; }
    public PdfReference(int obj, int gen) { ObjectNumber = obj; Generation = gen; }
    public override string ToString() => $"{ObjectNumber} {Generation} R";
}

/// <summary>PDF stream object (dictionary + bytes).</summary>
internal sealed class PdfStream : PdfObject {
    private byte[]? _data;
    private readonly ReadOnlyMemory<byte> _source;

    public PdfDictionary Dictionary { get; }
    public byte[] Data => Volatile.Read(ref _data) ?? Materialize(_source.ToArray());

    internal byte[] GetData(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[]? data = Volatile.Read(ref _data);
        if (data is not null) return data;
        data = new byte[_source.Length];
        CopyData(_source.Span, data, cancellationToken);
        return Materialize(data);
    }

    private byte[] Materialize(byte[] data) => Interlocked.CompareExchange(ref _data, data, null) ?? data;

    internal int DataLength => _data?.Length ?? _source.Length;
    internal long DataLongLength => DataLength;
    /// <summary>True when a decode filter failed; <see cref="Data"/> contains original undecoded bytes.</summary>
    public bool DecodingFailed { get; }
    /// <summary>Error message from decode failure, when available.</summary>
    public string? DecodingError { get; }
    public PdfStream(PdfDictionary dict, byte[] data, bool decodingFailed = false, string? error = null) {
        Dictionary = dict;
        _data = data;
        DecodingFailed = decodingFailed;
        DecodingError = error;
    }

    private PdfStream(PdfDictionary dictionary, ReadOnlyMemory<byte> source, bool decodingFailed, string? error) {
        Dictionary = dictionary;
        _source = source;
        DecodingFailed = decodingFailed;
        DecodingError = error;
    }

    internal static PdfStream FromOwnedSource(
        PdfDictionary dictionary,
        ReadOnlyMemory<byte> source,
        int sourceOffset,
        int sourceLength,
        bool decodingFailed = false,
        string? error = null) {
        if (sourceOffset < 0 || sourceLength < 0 || sourceOffset > source.Length - sourceLength) {
            throw new ArgumentOutOfRangeException(nameof(sourceLength));
        }
        return new PdfStream(dictionary, source.Slice(sourceOffset, sourceLength), decodingFailed, error);
    }

    internal void CopyDataTo(byte[] destination, int destinationOffset) =>
        GetDataMemory().Span.CopyTo(destination.AsSpan(destinationOffset));

    internal void CopyDataTo(byte[] destination, int destinationOffset, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        CopyData(GetDataMemory().Span, destination.AsSpan(destinationOffset), cancellationToken);
    }

    private static void CopyData(ReadOnlySpan<byte> source, Span<byte> destination, CancellationToken cancellationToken) {
        for (int offset = 0; offset < source.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(65536, source.Length - offset);
            source.Slice(offset, count).CopyTo(destination.Slice(offset));
            offset += count;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private ReadOnlyMemory<byte> GetDataMemory() {
        byte[]? data = Volatile.Read(ref _data);
        return data is not null ? data : _source;
    }

    internal void GetDataSegment(out byte[] buffer, out int offset, out int length, CancellationToken cancellationToken = default) {
        if (MemoryMarshal.TryGetArray(GetDataMemory(), out ArraySegment<byte> segment)) {
            buffer = segment.Array!;
            offset = segment.Offset;
            length = segment.Count;
            return;
        }

        buffer = GetData(cancellationToken);
        offset = 0;
        length = buffer.Length;
    }
}

/// <summary>PDF indirect object wrapper.</summary>
internal sealed class PdfIndirectObject : PdfObject {
    /// <summary>Byte offset of the selected definition, or its containing object stream, in the input PDF.</summary>
    internal int SourceOffset { get; set; } = -1;
    /// <summary>Chronological update containing this definition, derived from the active xref chain.</summary>
    internal int SourceRevision { get; set; } = -1;
    public int ObjectNumber { get; }
    public int Generation { get; }

    private PdfObject? _value;
    // Holds the source of a value that has been validated but not built yet. Cleared once the value exists.
    private PdfDeferredObjectSource? _deferred;
    private readonly PdfDeferredObjectSummary? _summary;

    /// <summary>The object's value. A deferred object is parsed on first access, once, and the same instance is returned afterwards.</summary>
    public PdfObject Value => Volatile.Read(ref _value) ?? Materialize();

    public PdfIndirectObject(int number, int generation, PdfObject value) { ObjectNumber = number; Generation = generation; _value = value; }

    /// <summary>Creates an object whose value is parsed from <paramref name="source"/> when first read.</summary>
    internal PdfIndirectObject(int number, int generation, PdfDeferredObjectSource source, PdfDeferredObjectSummary summary) {
        ObjectNumber = number;
        Generation = generation;
        _deferred = source;
        _summary = summary;
    }

    /// <summary>
    /// Returns the load-time summary of a value that has not been built yet. Whole-document passes use it
    /// in place of <see cref="Value"/> so they do not parse objects nobody reads.
    /// </summary>
    internal bool TryGetUnbuiltSummary(out PdfDeferredObjectSummary summary) {
        summary = _summary!;
        return _summary is not null && Volatile.Read(ref _value) is null;
    }

    /// <summary>True when <see cref="Value"/> is already built, so reading it cannot parse anything.</summary>
    internal bool IsValueBuilt => Volatile.Read(ref _value) is not null;

    private PdfObject Materialize() {
        // Parsing is a pure function of the retained bytes, so racing readers may each parse, but only the first
        // published value is kept and every caller returns it. No lock is taken: these objects also live in
        // identity-hashed sets, and locking an object whose header holds a hash code inflates a sync block.
        PdfDeferredObjectSource? deferred = Volatile.Read(ref _deferred);
        if (deferred is null) return Volatile.Read(ref _value)!;
        PdfObject parsed = deferred.Parse(_summary!.Start, _summary.Length);
        PdfObject value = Interlocked.CompareExchange(ref _value, parsed, null) ?? parsed;
        Volatile.Write(ref _deferred, null);
        return value;
    }
}

/// <summary>
/// Decoded bytes of an object stream whose objects were validated at load and are parsed only when their values are read.
/// One instance is shared by all deferred objects of the stream.
/// </summary>
internal sealed class PdfDeferredObjectSource {
    private readonly PdfByteContentText _text;
    private readonly PdfReadLimits _limits;
    private readonly bool _classicDictionaries;

    internal PdfDeferredObjectSource(byte[] data, PdfReadLimits limits)
        : this(new PdfByteContentText(data), limits, classicDictionaries: false) {
    }

    private PdfDeferredObjectSource(PdfByteContentText text, PdfReadLimits limits, bool classicDictionaries) {
        _text = text;
        _limits = limits;
        _classicDictionaries = classicDictionaries;
    }

    /// <summary>
    /// Dictionaries of top-level indirect objects in a stable input buffer. Their summaries span the text between
    /// the outer delimiters, which is what the eager load parses for such an object.
    /// </summary>
    internal static PdfDeferredObjectSource ForClassicDictionaries(PdfByteContentText text, PdfReadLimits limits) =>
        new(text, limits, classicDictionaries: true);

    internal bool Holds(byte[] data) => ReferenceEquals(_text.Source, data);

    // The load-time scan accepted these bytes under the same limits, so this parse cannot fail or exceed them.
    internal PdfObject Parse(int start, int length) => _classicDictionaries
        ? PdfSyntax.ParseDeferredClassicDictionary(_text, start, length, _limits)
        : PdfSyntax.ParseDeferredObject(_text, start, length, _limits);
}

/// <summary>
/// What load-time passes need from an object that has not been built: its outgoing references and a few
/// top-level facts. It is recorded only for objects whose syntax is complete and unambiguous.
/// </summary>
internal sealed class PdfDeferredObjectSummary {
    /// <summary>Marks <see cref="ParentReference"/> when the value has no top-level /Parent given by reference.</summary>
    internal const long NoReference = -1;

    internal PdfDeferredObjectSummary(int start, int length, long[] references, string? knownType, string? knownSubtype, bool hasEncryptKey, bool mayCarrySignatureData, ulong markerNames, long parentReference = NoReference) {
        ParentReference = parentReference;
        Start = start;
        Length = length;
        References = references;
        KnownType = knownType;
        KnownSubtype = knownSubtype;
        HasEncryptKey = hasEncryptKey;
        MayCarrySignatureData = mayCarrySignatureData;
        MarkerNames = markerNames;
    }

    /// <summary>
    /// One bit for each whole-document feature-marker name that occurs as a key or name value at any depth of the value.
    /// </summary>
    internal ulong MarkerNames { get; }

    /// <summary>
    /// The top-level /Parent when it is a reference, as object number (high 32 bits) and generation (low 32 bits);
    /// otherwise <see cref="NoReference"/>. The page-tree walk checks a leaf page's parent link with it.
    /// </summary>
    internal long ParentReference { get; }

    /// <summary>Offset of the object in its source's decoded bytes.</summary>
    internal int Start { get; }

    /// <summary>Length of the object in its source's decoded bytes.</summary>
    internal int Length { get; }

    /// <summary>Every reference in the value at any depth, as object number (high 32 bits) and generation (low 32 bits).</summary>
    internal long[] References { get; }

    /// <summary>
    /// The direct top-level /Type name when it is one that whole-document passes look for, such as Catalog, XRef or Page.
    /// Null for any other type, or when /Type is absent or not a direct name.
    /// </summary>
    internal string? KnownType { get; }

    /// <summary>The direct top-level /Subtype name when it is Widget, Link or FileAttachment; otherwise null.</summary>
    internal string? KnownSubtype { get; }

    /// <summary>The value is a dictionary with a top-level /Encrypt key.</summary>
    internal bool HasEncryptKey { get; }

    /// <summary>
    /// The value holds a /ByteRange, /Sig or /SigFlags name at any depth, or a top-level /FT or /Type given by reference.
    /// Objects without these cannot be signature fields, signature values or signature markers.
    /// </summary>
    internal bool MayCarrySignatureData { get; }
}

internal static class PdfObjectLookup {
    public static bool TryGet(
        System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects,
        PdfReference reference,
        out PdfIndirectObject indirect) {
        if (objects.TryGetValue(reference.ObjectNumber, out indirect!) &&
            indirect.Generation == reference.Generation) {
            return true;
        }

        indirect = null!;
        return false;
    }

    public static PdfObject? Resolve(
        System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value) {
        if (value is PdfReference reference &&
            TryGet(objects, reference, out var indirect)) {
            return indirect.Value;
        }

        return value;
    }

    public static PdfObject? ResolveChain(
        System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value) => ResolveChainCancellable(objects, value, default);

    public static PdfObject? ResolveChainCancellable(
        System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        System.Threading.CancellationToken cancellationToken) =>
        TryResolveReferenceChainCancellable(objects, value, out PdfObject? resolved, cancellationToken) ? resolved : null;

    public static bool TryResolveReferenceChain(
        System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        out PdfObject? resolved) => TryResolveReferenceChainCancellable(objects, value, out resolved, default);

    public static bool TryResolveReferenceChainCancellable(
        System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        out PdfObject? resolved,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var visited = new System.Collections.Generic.HashSet<(int ObjectNumber, int Generation)>();
        resolved = value;
        while (resolved is PdfReference reference) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !TryGet(objects, reference, out PdfIndirectObject indirect)) {
                resolved = null;
                return false;
            }
            resolved = indirect.Value;
        }
        return true;
    }
}
