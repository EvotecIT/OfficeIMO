using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Reader.Epub;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static OfficeConversionReport ConvertEpubPdf(Stream source, Stream destination, OfficeConversionOptions options,
        CancellationToken cancellationToken) {
        EpubLoadOptions reading = OfficeDocumentReaderBuilderEpubExtensions.Clone(options.ReaderEpub) ?? new EpubLoadOptions();
        reading.MaxPackageBytes = Math.Min(reading.MaxPackageBytes, options.MaximumInputBytes);
        reading.IncludeRawHtml = true;
        reading.IncludeResourceData = EpubPdfExport.AllowsPackageResources(options.Html);
        EpubDocument book = EpubDocument.Load(source, reading, cancellationToken);
        return EpubPdfExport.Write(book, destination, options.Html, options.FailOnLoss, cancellationToken);
    }
}
