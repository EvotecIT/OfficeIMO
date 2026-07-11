namespace OfficeIMO.OpenDocument;

/// <summary>Logical value kinds supported by the native ODS cell model.</summary>
public enum OdsCellValueKind {
    /// <summary>No value.</summary>
    Empty,
    /// <summary>Text value.</summary>
    String,
    /// <summary>Floating-point or decimal value.</summary>
    Number,
    /// <summary>Boolean value.</summary>
    Boolean,
    /// <summary>Date or date-time value.</summary>
    Date,
    /// <summary>Time-of-day or duration value represented by an ODF duration.</summary>
    Time,
    /// <summary>Percentage stored as a fractional number.</summary>
    Percentage,
    /// <summary>Currency value with an ISO currency code.</summary>
    Currency
}

/// <summary>An immutable typed ODS cell value retaining its invariant lexical form.</summary>
public readonly struct OdsCellValue {
    internal OdsCellValue(OdsCellValueKind kind, string lexicalValue, string displayText, string? currencyCode = null) {
        Kind = kind; LexicalValue = lexicalValue; DisplayText = displayText; CurrencyCode = currencyCode;
    }
    /// <summary>Empty value.</summary>
    public static OdsCellValue Empty => new OdsCellValue(OdsCellValueKind.Empty, string.Empty, string.Empty);
    /// <summary>Logical value kind.</summary>
    public OdsCellValueKind Kind { get; }
    /// <summary>Invariant value as stored in the ODF attribute.</summary>
    public string LexicalValue { get; }
    /// <summary>Displayed text stored in cell paragraphs, when available.</summary>
    public string DisplayText { get; }
    /// <summary>Currency code for currency values.</summary>
    public string? CurrencyCode { get; }
    /// <summary>Parses the value as a decimal.</summary>
    public decimal AsDecimal() => decimal.Parse(LexicalValue, NumberStyles.Float, CultureInfo.InvariantCulture);
    /// <summary>Parses the value as a double.</summary>
    public double AsDouble() => double.Parse(LexicalValue, NumberStyles.Float, CultureInfo.InvariantCulture);
    /// <summary>Parses the value as a boolean.</summary>
    public bool AsBoolean() => bool.Parse(LexicalValue);
    /// <summary>Parses the value as an ODF date or date-time.</summary>
    public DateTimeOffset AsDateTimeOffset() => DateTimeOffset.Parse(LexicalValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
    /// <summary>Parses the ODF duration used for time and duration cells.</summary>
    public TimeSpan AsTimeSpan() => XmlConvert.ToTimeSpan(LexicalValue);
    /// <inheritdoc />
    public override string ToString() => DisplayText.Length > 0 ? DisplayText : LexicalValue;
}
