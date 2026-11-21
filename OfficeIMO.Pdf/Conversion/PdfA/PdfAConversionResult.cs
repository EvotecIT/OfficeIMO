#pragma warning disable RS0026 // Path and stream destinations share the save name and its optional parameters, as PdfDocument.Save does.
namespace OfficeIMO.Pdf;

/// <summary>The PDF/A file produced by conversion, or the report of why the source was refused.</summary>
public sealed class PdfAConversionResult : IOfficeResult<PdfDocument> {
    private readonly byte[]? _pdf;
    private PdfDocument? _document;

    internal PdfAConversionResult(byte[]? pdf, PdfAConversionReport report) {
        _pdf = pdf;
        Report = report;
    }

    /// <summary>Gets whether conversion produced a file. When <see langword="false"/>, <see cref="PdfAConversionReport.Violations"/> lists every reason.</summary>
    public bool Succeeded => _pdf is not null;

    /// <summary>Gets the changes made, or the reasons the source was refused.</summary>
    public PdfAConversionReport Report { get; }

    /// <summary>Gets the converted document, or <see langword="null"/> when conversion was refused.</summary>
    public PdfDocument? Value => _pdf is null ? null : _document ??= PdfDocument.Load(_pdf);

    /// <summary>Returns the converted document, or throws <see cref="PdfAConversionException"/> listing every violation.</summary>
    public PdfDocument RequireValue() => Value ?? throw new PdfAConversionException(Report);

    /// <summary>Returns a copy of the converted bytes, or throws <see cref="PdfAConversionException"/> listing every violation.</summary>
    public byte[] ToBytes() => _pdf is null ? throw new PdfAConversionException(Report) : (byte[])_pdf.Clone();

    /// <summary>Writes the converted file to a caller-owned stream, or throws <see cref="PdfAConversionException"/> without writing to it.</summary>
    public void Save(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        byte[] pdf = _pdf ?? throw new PdfAConversionException(Report);
        OfficeIMO.Core.Internal.OfficeStreamWriter.CopyBytes(stream, pdf, cancellationToken);
    }

    /// <summary>Writes the converted file, or throws <see cref="PdfAConversionException"/> without creating <paramref name="path"/>.</summary>
    public void Save(string path, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        byte[] pdf = _pdf ?? throw new PdfAConversionException(Report);
        OfficeIMO.Core.Internal.OfficeFileCommit.Write(path,
            stream => OfficeIMO.Core.Internal.OfficeStreamWriter.CopyBytes(stream, pdf, cancellationToken), cancellationToken);
    }
}