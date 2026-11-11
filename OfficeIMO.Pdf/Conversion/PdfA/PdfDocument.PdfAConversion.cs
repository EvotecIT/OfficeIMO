namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Converts this PDF to PDF/A. Conversion fixes what it can without moving anything on the page and reports every change.
    /// When something cannot be made compliant without changing the content, nothing is produced and the result lists every
    /// violation by object.
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
}
