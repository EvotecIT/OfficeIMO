#pragma warning disable RS0026 // Path and stream destinations share the save name and its optional parameters, as PdfDocument.Save does.
namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Converts this PDF to PDF/A-2b, 2u, 3b, 3u, 4 or 4f, as <see cref="PdfAConversionOptions.Profile"/> selects (PDF/A-2b by
    /// default). Conversion fixes what it can without moving anything on the page and reports every change. When something
    /// cannot be made compliant without changing the content, nothing is produced and the result lists every violation by
    /// object. An encrypted PDF must have been loaded with its owner password in <see cref="PdfLoadOptions.Password"/>.
    /// </summary>
    public PdfAConversionResult ConvertToPdfA(PdfAConversionOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return PdfAConverter.Convert(GetBytesForOperation(cancellationToken), ReadOptions, options ?? new PdfAConversionOptions(), cancellationToken);
    }

    /// <summary>
    /// Converts this PDF to PDF/A and writes it to <paramref name="path"/>. Throws <see cref="PdfAConversionException"/>
    /// listing every violation, without creating the file, when the source cannot be converted without changing its content.
    /// </summary>
    public PdfAConversionReport SaveAsPdfA(string path, PdfAConversionOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        PdfAConversionResult result = ConvertToPdfA(options, cancellationToken);
        result.Save(path, cancellationToken);
        return result.Report;
    }

    /// <summary>
    /// Converts this PDF to PDF/A and writes it to a caller-owned stream. Throws <see cref="PdfAConversionException"/> listing
    /// every violation, without writing to the stream, when the source cannot be converted without changing its content.
    /// </summary>
    public PdfAConversionReport SaveAsPdfA(Stream stream, PdfAConversionOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        PdfAConversionResult result = ConvertToPdfA(options, cancellationToken);
        result.Save(stream, cancellationToken);
        return result.Report;
    }
}