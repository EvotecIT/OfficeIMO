#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Represents a parse error collected while reading CSV data.
/// </summary>
public sealed class CsvParseError
{
    /// <summary>
    /// Initializes a collected parse error.
    /// </summary>
    public CsvParseError(int lineNumber, string message, Exception exception)
    {
        LineNumber = lineNumber;
        Message = message;
        Exception = exception;
    }

    /// <summary>
    /// Gets the line where parsing failed.
    /// </summary>
    public int LineNumber { get; }

    /// <summary>
    /// Gets the parse error message.
    /// </summary>
    public string Message { get; }

    /// <summary>
    /// Gets the original exception.
    /// </summary>
    public Exception Exception { get; }
}
