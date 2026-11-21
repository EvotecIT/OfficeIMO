namespace OfficeIMO.Toml;

/// <summary>The syntax element <see cref="TomlReader"/> stopped on.</summary>
public enum TomlTokenType : byte {
    /// <summary>Before the first <see cref="TomlReader.Read"/> or after the end of the input.</summary>
    None = 0,
    /// <summary>The <c>[</c> of a table header. Its keys follow as <see cref="Key"/> tokens, then <see cref="TableHeaderEnd"/>.</summary>
    TableHeaderStart,
    /// <summary>The <c>]</c> that closes a table header.</summary>
    TableHeaderEnd,
    /// <summary>The <c>[[</c> of an array-of-tables header. Its keys follow as <see cref="Key"/> tokens, then <see cref="ArrayOfTablesHeaderEnd"/>.</summary>
    ArrayOfTablesHeaderStart,
    /// <summary>The <c>]]</c> that closes an array-of-tables header.</summary>
    ArrayOfTablesHeaderEnd,
    /// <summary>One segment of a key: <c>a.b.c = 1</c> reads as three keys followed by the value.</summary>
    Key,
    /// <summary>A basic, literal or multi-line string.</summary>
    String,
    /// <summary>A decimal, hexadecimal, octal or binary integer.</summary>
    Integer,
    /// <summary>A float, including <c>inf</c> and <c>nan</c>.</summary>
    Float,
    /// <summary><c>true</c> or <c>false</c>.</summary>
    Boolean,
    /// <summary>A date-time with an offset, such as <c>1979-05-27T07:32:00Z</c>.</summary>
    OffsetDateTime,
    /// <summary>A date-time without an offset.</summary>
    LocalDateTime,
    /// <summary>A date without a time.</summary>
    LocalDate,
    /// <summary>A time of day without a date.</summary>
    LocalTime,
    /// <summary>The <c>[</c> of an array value.</summary>
    StartArray,
    /// <summary>The <c>]</c> of an array value.</summary>
    EndArray,
    /// <summary>The <c>{</c> of an inline table.</summary>
    StartInlineTable,
    /// <summary>The <c>}</c> of an inline table.</summary>
    EndInlineTable,
    /// <summary>A comment, reported only when <see cref="TomlReaderOptions.IncludeComments"/> is set.</summary>
    Comment
}

/// <summary>How a string or key is written in the source.</summary>
public enum TomlStringStyle : byte {
    /// <summary>A bare key such as <c>name</c>. Strings are never bare.</summary>
    Bare = 0,
    /// <summary>A basic string in double quotes, with escapes.</summary>
    Basic,
    /// <summary>A literal string in single quotes, without escapes.</summary>
    Literal,
    /// <summary>A multi-line basic string in triple double quotes.</summary>
    MultilineBasic,
    /// <summary>A multi-line literal string in triple single quotes.</summary>
    MultilineLiteral
}

/// <summary>The kind of a <see cref="TomlNode"/>.</summary>
public enum TomlNodeKind : byte {
    /// <summary>A <see cref="TomlTable"/>: the root, a <c>[table]</c>, an inline table or a table made by dotted keys.</summary>
    Table,
    /// <summary>A <see cref="TomlTableArray"/>: the tables of <c>[[name]]</c> headers.</summary>
    TableArray,
    /// <summary>A <see cref="TomlArray"/>: an array value in brackets.</summary>
    Array,
    /// <summary>A string <see cref="TomlValue"/>.</summary>
    String,
    /// <summary>An integer <see cref="TomlValue"/>.</summary>
    Integer,
    /// <summary>A float <see cref="TomlValue"/>.</summary>
    Float,
    /// <summary>A boolean <see cref="TomlValue"/>.</summary>
    Boolean,
    /// <summary>A date, time or date-time <see cref="TomlValue"/>; <see cref="TomlDateTime.Kind"/> tells which.</summary>
    DateTime
}

/// <summary>How a <see cref="TomlTable"/> is defined in the source.</summary>
public enum TomlTableKind : byte {
    /// <summary>The document's root table: the key/value pairs before the first header.</summary>
    Root,
    /// <summary>A table defined by a <c>[name]</c> header.</summary>
    Header,
    /// <summary>A table that exists only because a header names a table inside it, as <c>[a.b]</c> makes <c>a</c>.</summary>
    Implicit,
    /// <summary>A table made by dotted keys, as <c>a.b = 1</c> makes <c>a</c>.</summary>
    Dotted,
    /// <summary>An inline table in braces. It is complete as written: nothing outside the braces adds to it.</summary>
    Inline,
    /// <summary>One table of an array of tables, defined by a <c>[[name]]</c> header.</summary>
    ArrayItem
}

/// <summary>Which date and time fields a <see cref="TomlDateTime"/> has.</summary>
public enum TomlDateTimeKind : byte {
    /// <summary>A date, time and UTC offset: <c>1979-05-27T07:32:00-07:00</c>.</summary>
    OffsetDateTime,
    /// <summary>A date and time without an offset: <c>1979-05-27T07:32:00</c>.</summary>
    LocalDateTime,
    /// <summary>A date: <c>1979-05-27</c>.</summary>
    LocalDate,
    /// <summary>A time of day: <c>07:32:00</c>.</summary>
    LocalTime
}
