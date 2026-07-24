namespace OfficeIMO.Pdf;

/// <summary>One named scalar or multi-value AcroForm data entry.</summary>
public sealed class PdfFormDataField {
    private readonly string[] _values;
    /// <summary>Creates one form-data entry.</summary>
    public PdfFormDataField(string name, IEnumerable<string> values) {
        Guard.NotNullOrWhiteSpace(name, nameof(name)); Guard.NotNull(values, nameof(values));
        _values = values.ToArray();
        if (_values.Length == 0 || _values.Any(static value => value is null)) throw new ArgumentException("Form data fields require at least one non-null value.", nameof(values));
        Name = name;
    }
    /// <summary>Fully qualified field name.</summary>
    public string Name { get; }
    /// <summary>Field values in source order.</summary>
    public IReadOnlyList<string> Values => Array.AsReadOnly(_values);
}

/// <summary>Dependency-free AcroForm data set with typed and XFDF interchange.</summary>
public sealed class PdfFormDataSet {
    /// <summary>Default maximum XFDF source length accepted before XML materialization.</summary>
    public const int DefaultMaxXfdfDocumentCharacters = 8_000_000;
    /// <summary>Default maximum number of XFDF fields.</summary>
    public const int DefaultMaxXfdfFields = 100_000;
    /// <summary>Default maximum aggregate XFDF value length.</summary>
    public const int DefaultMaxXfdfValueCharacters = 4_000_000;

    private readonly PdfFormDataField[] _fields;
    /// <summary>Creates a unique-name data set.</summary>
    public PdfFormDataSet(IEnumerable<PdfFormDataField> fields) {
        Guard.NotNull(fields, nameof(fields)); _fields = fields.ToArray();
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfFormDataField field in _fields) { Guard.NotNull(field, nameof(fields)); if (!names.Add(field.Name)) throw new ArgumentException("Form data field names must be unique: " + field.Name, nameof(fields)); }
    }
    /// <summary>Data fields in deterministic source order.</summary>
    public IReadOnlyList<PdfFormDataField> Fields => Array.AsReadOnly(_fields);
    /// <summary>Converts values to the shared form-filler contract.</summary>
    public IReadOnlyDictionary<string, PdfFormFieldValue> ToFieldValues() {
        var values = new Dictionary<string, PdfFormFieldValue>(StringComparer.Ordinal);
        foreach (PdfFormDataField field in _fields) values[field.Name] = PdfFormFieldValue.FromValues(field.Values);
        return new System.Collections.ObjectModel.ReadOnlyDictionary<string, PdfFormFieldValue>(values);
    }
    /// <summary>Serializes the data set as XFDF 1.0 XML.</summary>
    public string ToXfdf() {
        var builder = new StringBuilder();
        var settings = new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true, OmitXmlDeclaration = false };
        using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(builder, settings)) {
            writer.WriteStartDocument(); writer.WriteStartElement("xfdf", "http://ns.adobe.com/xfdf/"); writer.WriteStartElement("fields");
            foreach (PdfFormDataField field in _fields) { writer.WriteStartElement("field"); writer.WriteAttributeString("name", field.Name); foreach (string value in field.Values) writer.WriteElementString("value", value); writer.WriteEndElement(); }
            writer.WriteEndElement(); writer.WriteEndElement(); writer.WriteEndDocument();
        }
        return builder.ToString();
    }
    /// <summary>Parses bounded, DTD-free XFDF field data.</summary>
    public static PdfFormDataSet ParseXfdf(
        string xfdf,
        int maxFields = DefaultMaxXfdfFields,
        int maxValueCharacters = DefaultMaxXfdfValueCharacters,
        int maxDocumentCharacters = DefaultMaxXfdfDocumentCharacters) {
        Guard.NotNull(xfdf, nameof(xfdf));
#pragma warning disable CA1512 // ThrowIfNegativeOrZero is unavailable on every target framework.
        if (maxFields <= 0) throw new ArgumentOutOfRangeException(nameof(maxFields));
        if (maxValueCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxValueCharacters));
        if (maxDocumentCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxDocumentCharacters));
#pragma warning restore CA1512
        if (xfdf.Length > maxDocumentCharacters) throw new InvalidOperationException("XFDF document character limit exceeded.");
        var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = maxDocumentCharacters };
        var document = new System.Xml.XmlDocument { XmlResolver = null };
        using (var reader = System.Xml.XmlReader.Create(new StringReader(xfdf), settings)) {
            document.Load(reader);
        }
        var fields = new List<PdfFormDataField>(); int valueCharacters = 0;
        foreach (System.Xml.XmlNode node in document.GetElementsByTagName("field", "*")) {
            if (fields.Count >= maxFields) throw new InvalidOperationException("XFDF field count limit exceeded.");
            string? name = node.Attributes?["name"]?.Value; var values = new List<string>();
            foreach (System.Xml.XmlNode child in node.ChildNodes) {
                if (child.LocalName != "value") continue; string value = child.InnerText; valueCharacters = checked(valueCharacters + value.Length); if (valueCharacters > maxValueCharacters) throw new InvalidOperationException("XFDF value character limit exceeded."); values.Add(value);
            }
            if (values.Count == 0) values.Add(string.Empty); fields.Add(new PdfFormDataField(name ?? throw new FormatException("XFDF field name is missing."), values));
        }
        return new PdfFormDataSet(fields);
    }
}
