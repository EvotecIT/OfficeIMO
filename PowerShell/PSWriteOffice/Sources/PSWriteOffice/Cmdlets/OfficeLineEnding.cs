namespace PSWriteOffice.Cmdlets;

/// <summary>Canonical text-file line endings.</summary>
public enum OfficeLineEnding
{
    /// <summary>Line feed used by Unix-like systems.</summary>
    LF,
    /// <summary>Carriage return followed by line feed, used by Windows.</summary>
    CRLF,
    /// <summary>Carriage return used by classic Mac text files.</summary>
    CR
}

internal static class OfficeLineEndingUtilities
{
    internal static string ToText(OfficeLineEnding lineEnding) => lineEnding switch
    {
        OfficeLineEnding.LF => "\n",
        OfficeLineEnding.CRLF => "\r\n",
        OfficeLineEnding.CR => "\r",
        _ => throw new System.ArgumentOutOfRangeException(nameof(lineEnding), lineEnding, null)
    };
}
