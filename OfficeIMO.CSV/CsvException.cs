#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Base exception for CSV related failures.
/// </summary>
public class CsvException : Exception
{
    /// <summary>
    /// Initializes a new instance of <see cref="CsvException"/>.
    /// </summary>
    public CsvException(string message, Exception? innerException = null) : base(message, innerException)
    {
    }
}

/// <summary>
/// Raised when the CSV content cannot be parsed.
/// </summary>
public sealed class CsvParseException : CsvException
{
    /// <summary>
    /// Initializes a new instance of <see cref="CsvParseException"/>.
    /// </summary>
    public CsvParseException(string message, int? lineNumber = null, Exception? innerException = null)
        : base(lineNumber.HasValue ? $"Line {lineNumber.Value}: {message}" : message, innerException)
    {
        LineNumber = lineNumber;
    }

    /// <summary>
    /// Gets the 1-based line number where the parsing error occurred, when available.
    /// </summary>
    public int? LineNumber { get; }
}

/// <summary>
/// Raised when a CSV document fails schema validation.
/// </summary>
public sealed class CsvValidationException : CsvException
{
    /// <summary>
    /// Initializes a new instance of <see cref="CsvValidationException"/>.
    /// </summary>
    public CsvValidationException(string message, IReadOnlyList<CsvValidationError> errors)
        : base(message)
    {
        Errors = errors;
    }

    /// <summary>
    /// Gets the collection of validation errors that triggered the exception.
    /// </summary>
    public IReadOnlyList<CsvValidationError> Errors { get; }
}
