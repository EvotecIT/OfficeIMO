namespace OfficeIMO.Rtf;

/// <summary>
/// Built-in generated text controls represented directly by RTF control words.
/// </summary>
public enum RtfGeneratedTextKind {
    /// <summary>Current page number, emitted as <c>\chpgn</c>.</summary>
    PageNumber,

    /// <summary>Current section number, emitted as <c>\sectnum</c>.</summary>
    SectionNumber,

    /// <summary>Current date, emitted as <c>\chdate</c>.</summary>
    CurrentDate,

    /// <summary>Current date in long format, emitted as <c>\chdpl</c>.</summary>
    CurrentDateLong,

    /// <summary>Current date in abbreviated format, emitted as <c>\chdpa</c>.</summary>
    CurrentDateAbbreviated,

    /// <summary>Current time, emitted as <c>\chtime</c>.</summary>
    CurrentTime,

    /// <summary>Automatic note reference marker, emitted as <c>\chftn</c>.</summary>
    NoteReference
}
