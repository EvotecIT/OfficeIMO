namespace OfficeIMO.Pdf;

internal abstract class PdfObject { }

/// <summary>PDF numeric value (integer or real).</summary>
internal sealed class PdfNumber : PdfObject {
    public double Value { get; }
    public PdfNumber(double value) { Value = value; }
    public override string ToString() => Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
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
    public PdfStringObj(string value) { Value = value; }
    public override string ToString() => Value;
}

/// <summary>PDF array object.</summary>
internal sealed class PdfArray : PdfObject {
    public System.Collections.Generic.List<PdfObject> Items { get; } = new();
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
    public PdfDictionary Dictionary { get; }
    public byte[] Data { get; }
    public PdfStream(PdfDictionary dict, byte[] data) { Dictionary = dict; Data = data; }
}

/// <summary>PDF indirect object wrapper.</summary>
internal sealed class PdfIndirectObject : PdfObject {
    public int ObjectNumber { get; }
    public int Generation { get; }
    public PdfObject Value { get; }
    public PdfIndirectObject(int number, int generation, PdfObject value) { ObjectNumber = number; Generation = generation; Value = value; }
}
