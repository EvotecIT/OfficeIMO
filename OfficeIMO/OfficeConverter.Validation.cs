using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task ValidateRegisteredOutputAsync(Stream output, OfficeConversionOptions settings, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        output.Position = 0;
        string extension = Path.GetExtension(settings.TargetName ?? string.Empty).ToLowerInvariant();
        switch (FormatForExtension(extension)) {
            case OfficeDocumentFormat.Pdf: {
                PdfCreateOptions? creation = settings.Word?.PdfOptions ?? settings.Excel?.PdfOptions ?? settings.PowerPoint?.PdfOptions
                    ?? settings.Markdown?.PdfOptions ?? settings.Html?.PdfOptions ?? settings.PlainText?.PdfOptions ?? settings.ReaderPdf?.PdfOptions;
                var options = new PdfLoadOptions { Password = creation?.Encryption?.OwnerPassword ?? creation?.Encryption?.UserPassword ?? settings.SourcePassword };
                if (creation?.Encryption?.AesCryptographyProvider is { } cryptography)
                    options = PdfLoadOptions.WithAesCryptographyProvider(options, cryptography);
                if (Registrations.TryGetValue(SelectRoute(settings).Id, out IOfficeConversionRoute? route) && route is IOfficeConversionOutputValidation owner)
                    options = owner.CreateOutputPdfLoadOptions(settings) ?? options;
                using PdfDocument document = await PdfDocument.LoadAsync(output, options, token).ConfigureAwait(false);
                if (document.Inspect(options, token).PageCount == 0) throw new InvalidDataException("The registered converter produced a PDF without pages.");
                break;
            }
            case OfficeDocumentFormat.Word: {
                WordLoadOptions options = WordInputOptions(settings);
                options.AccessMode = DocumentAccessMode.ReadOnly;
                options.MaxInputBytes = settings.MaximumOutputBytes;
                options.PackageSecurity!.MaxPackageBytes = settings.MaximumOutputBytes;
                options.Password = settings.WordSave?.ResolvePassword(null);
                using WordDocument document = await WordDocument.LoadAsync(output, options, token).ConfigureAwait(false);
                break;
            }
            case OfficeDocumentFormat.Excel: {
                ExcelLoadOptions options = ExcelInputOptions(settings);
                options.AccessMode = DocumentAccessMode.ReadOnly;
                options.MaxInputBytes = settings.MaximumOutputBytes;
                options.PackageSecurity!.MaxPackageBytes = settings.MaximumOutputBytes;
                options.Password = settings.ExcelSave?.ResolvePassword(null);
                using ExcelDocument document = await ExcelDocument.LoadAsync(output, options, token).ConfigureAwait(false);
                break;
            }
            case OfficeDocumentFormat.PowerPoint: {
                PowerPointLoadOptions options = PowerPointInputOptions(settings);
                options.AccessMode = DocumentAccessMode.ReadOnly;
                options.MaxInputBytes = settings.MaximumOutputBytes;
                options.PackageSecurity!.MaxPackageBytes = settings.MaximumOutputBytes;
                options.Password = settings.PowerPointSave?.ResolvePassword(null);
                using PowerPointDocument document = await PowerPointDocument.LoadAsync(output, options, token).ConfigureAwait(false);
                break;
            }
            case OfficeDocumentFormat.Html:
                _ = await HtmlConversionDocument.LoadAsync(output, cancellationToken: token).ConfigureAwait(false);
                break;
            case OfficeDocumentFormat.Markdown: {
                using var reader = new StreamReader(output, new System.Text.UTF8Encoding(false, true), true, 8192, leaveOpen: true);
                using var document = Markdown.MarkdownDocument.Parse(await reader.ReadToEndAsync(token).ConfigureAwait(false));
                break;
            }
            case OfficeDocumentFormat.Epub: {
                var document = Epub.EpubPublication.Load(output, cancellationToken: token);
                _ = document.Write(cancellationToken: token);
                break;
            }
            default: throw new NotSupportedException("The registered conversion target has no artifact validator.");
        }
        token.ThrowIfCancellationRequested();
        output.Position = output.Length;
    }
}
