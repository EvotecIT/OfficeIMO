namespace OfficeIMO.Data.Conversions;

/// <summary>The structural events every source produces and every sink consumes.</summary>
internal enum DataTokenType : byte {
    None = 0,
    StartDocument,
    EndDocument,
    StartObject,
    EndObject,
    StartArray,
    EndArray,
    PropertyName,
    Scalar,
    Comment
}

/// <summary>
/// The type of a scalar. Numbers and date-times travel with their source text, so a sink writes <c>1.10</c> or a
/// 23-digit integer exactly when its format can hold it.
/// </summary>
internal enum DataScalarKind : byte {
    Null = 0,
    Boolean,
    /// <summary>Text is a canonical decimal integer: optional <c>-</c>, digits, no leading zeros.</summary>
    Integer,
    /// <summary>Text is a JSON number, or <c>NaN</c>, <c>Infinity</c> or <c>-Infinity</c>.</summary>
    Float,
    String,
    /// <summary>RFC 3339 date-time with an offset (<c>Z</c> or <c>±hh:mm</c>).</summary>
    OffsetDateTime,
    /// <summary>RFC 3339 date-time without an offset.</summary>
    LocalDateTime,
    /// <summary>RFC 3339 full-date.</summary>
    LocalDate,
    /// <summary>RFC 3339 partial-time.</summary>
    LocalTime
}

internal static class DataScalarKinds {
    public static bool IsDateTime(DataScalarKind kind) => kind >= DataScalarKind.OffsetDateTime;

    public static bool IsNonFinite(ReadOnlySpan<char> text) =>
        text.Length > 0 && (text[0] == 'N' || text[0] == 'I' || (text.Length > 1 && text[0] == '-' && text[1] == 'I'));
}
