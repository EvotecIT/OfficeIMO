namespace OfficeIMO.Word;

/// <summary>
/// Exception thrown when an unsupported image format is encountered.
/// </summary>
public class ImageFormatNotSupportedException : OfficeIMOException {
    /// <summary>
    /// Initializes a new instance of the <see cref="ImageFormatNotSupportedException"/> class.
    /// </summary>
    /// <param name="message">Exception message.</param>
    public ImageFormatNotSupportedException(string message) : base(message) {

    }
}
