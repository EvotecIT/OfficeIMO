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
    public System.Collections.Generic.Dictionary<string, PdfObject> Items { get; } = new();
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
    private readonly byte[]? _source;
    private readonly int _sourceOffset;
    private readonly int _sourceLength;

    public PdfDictionary Dictionary { get; }
    public byte[] Data {
        get {
            byte[]? data = System.Threading.Volatile.Read(ref _data);
            if (data is not null) return data;
            data = new byte[_sourceLength];
            Buffer.BlockCopy(_source!, _sourceOffset, data, 0, data.Length);
            return System.Threading.Interlocked.CompareExchange(ref _data, data, null) ?? data;
        }
    }
    internal int DataLength => _data?.Length ?? _sourceLength;
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

    private PdfStream(
        PdfDictionary dictionary,
        byte[] ownedSource,
        int sourceOffset,
        int sourceLength) {
        Dictionary = dictionary;
        _source = ownedSource;
        _sourceOffset = sourceOffset;
        _sourceLength = sourceLength;
    }

    internal static PdfStream FromOwnedSource(
        PdfDictionary dictionary,
        byte[] ownedSource,
        int sourceOffset,
        int sourceLength) {
        if (sourceOffset < 0 || sourceLength < 0 || sourceOffset > ownedSource.Length - sourceLength) {
            throw new ArgumentOutOfRangeException(nameof(sourceLength));
        }
        return new PdfStream(dictionary, ownedSource, sourceOffset, sourceLength);
    }

    internal void CopyDataTo(byte[] destination, int destinationOffset) {
        byte[]? data = System.Threading.Volatile.Read(ref _data);
        if (data is not null) {
            Buffer.BlockCopy(data, 0, destination, destinationOffset, data.Length);
            return;
        }
        Buffer.BlockCopy(_source!, _sourceOffset, destination, destinationOffset, _sourceLength);
    }

    internal void GetDataSegment(out byte[] buffer, out int offset, out int length) {
        byte[]? data = System.Threading.Volatile.Read(ref _data);
        if (data is not null) {
            buffer = data;
            offset = 0;
            length = data.Length;
            return;
        }

        buffer = _source!;
        offset = _sourceOffset;
        length = _sourceLength;
    }
}

/// <summary>PDF indirect object wrapper.</summary>
internal sealed class PdfIndirectObject : PdfObject {
    public int ObjectNumber { get; }
    public int Generation { get; }
    public PdfObject Value { get; }
    public PdfIndirectObject(int number, int generation, PdfObject value) { ObjectNumber = number; Generation = generation; Value = value; }
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
        PdfObject? value) =>
        TryResolveReferenceChain(objects, value, out PdfObject? resolved) ? resolved : null;

    public static bool TryResolveReferenceChain(
        System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        out PdfObject? resolved) {
        var visited = new System.Collections.Generic.HashSet<(int ObjectNumber, int Generation)>();
        resolved = value;
        while (resolved is PdfReference reference) {
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
