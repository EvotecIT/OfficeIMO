using System;
using System.IO;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;

namespace PSWriteOffice.Services;

/// <summary>Admits the six native image source families without saving caller-owned documents.</summary>
internal static class OfficeDocumentImageSource {
    internal static OfficeDocumentFormat Format(object document) => document switch {
        WordDocument => OfficeDocumentFormat.Word, ExcelDocument => OfficeDocumentFormat.Excel,
        PowerPointDocument => OfficeDocumentFormat.PowerPoint, PdfDocument => OfficeDocumentFormat.Pdf,
        VisioDocument => OfficeDocumentFormat.Visio, HtmlConversionDocument => OfficeDocumentFormat.Html,
        _ => throw new PSArgumentException("Document requires a native Word, Excel, PowerPoint, PDF, Visio or HTML document.")
    };

    internal static OfficeDocumentFormat Format(string path) {
        OfficeDocumentFormat format = OfficeConverter.FormatForExtension(Path.GetExtension(path));
        if (format is not (OfficeDocumentFormat.Word or OfficeDocumentFormat.Excel or OfficeDocumentFormat.PowerPoint
            or OfficeDocumentFormat.Pdf or OfficeDocumentFormat.Visio or OfficeDocumentFormat.Html))
            throw new PSArgumentException("The source has no native document image exporter: " + path);
        return format;
    }

    internal static async Task<object> LoadAsync(string path, OfficeDocumentFormat format,
        OfficeDocumentImageOptions options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        switch (format) {
            case OfficeDocumentFormat.Word:
                options.WordLoad.AccessMode = DocumentAccessMode.ReadOnly;
                options.WordLoad.PersistenceMode = DocumentPersistenceMode.Explicit;
                return await WordDocument.LoadAsync(path, options.WordLoad, token).ConfigureAwait(false);
            case OfficeDocumentFormat.Excel:
                options.ExcelLoad.AccessMode = DocumentAccessMode.ReadOnly;
                options.ExcelLoad.PersistenceMode = DocumentPersistenceMode.Explicit;
                return await ExcelDocument.LoadAsync(path, options.ExcelLoad, token).ConfigureAwait(false);
            case OfficeDocumentFormat.PowerPoint:
                options.PowerPointLoad.AccessMode = DocumentAccessMode.ReadOnly;
                options.PowerPointLoad.PersistenceMode = DocumentPersistenceMode.Explicit;
                return await PowerPointDocument.LoadAsync(path, options.PowerPointLoad, token).ConfigureAwait(false);
            case OfficeDocumentFormat.Pdf: return await PdfDocument.LoadAsync(path, options.PdfLoad, token).ConfigureAwait(false);
            case OfficeDocumentFormat.Visio: return await VisioDocument.LoadAsync(path, options.VisioLoad, token).ConfigureAwait(false);
            case OfficeDocumentFormat.Html:
                return await HtmlConversionDocument.LoadAsync(path, options.HtmlLoad, cancellationToken: token).ConfigureAwait(false);
            default: throw new PSArgumentException("The source has no native document image exporter.");
        }
    }
}
