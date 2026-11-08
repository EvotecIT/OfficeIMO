using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        private async Task<OfficeConversionReport> ConvertRtfAsync(Stream input, Stream output, OfficeConversionOptions options,
            OfficeWorkflowRouteSettings settings, CancellationToken token) {
            if (Id.StartsWith("rtf-", StringComparison.Ordinal)) {
                RtfReadResult read = await RtfDocument.LoadResultAsync(input, settings.RtfLoad, cancellationToken: token).ConfigureAwait(false);
                using RtfDocument document = read.Document;
                if (options.Theme != null && Id != "rtf-markdown") document.ApplyTheme(options.Theme);
                var source = new RtfConversionReport();
                source.AddReadDiagnostics(read.Diagnostics, options.SourceName);
                if (Id == "rtf-html") {
                    RtfHtmlExportOptions conversion = settings.RtfHtml ?? new();
                    conversion.FailOnLoss |= options.FailOnLoss;
                    return await RetainSourceAsync(source, () => document.SaveAsHtmlAsync(output, conversion, token), conversion.FailOnLoss).ConfigureAwait(false);
                }
                if (Id == "rtf-markdown") {
                    RtfMarkdownExportOptions conversion = settings.RtfMarkdown ?? new();
                    conversion.FailOnLoss |= options.FailOnLoss;
                    return await RetainSourceAsync(source, () => document.SaveAsMarkdownAsync(output, conversion, token), conversion.FailOnLoss).ConfigureAwait(false);
                }
                RtfWordExportOptions word = settings.RtfWord ?? new();
                word.FailOnLoss |= options.FailOnLoss;
                return await SaveModelAsync(await read.ToWordAsync(word, token).ConfigureAwait(false),
                    model => model.SaveAsync(output, null, token), token).ConfigureAwait(false);
            }
            if (Id == "docx-rtf") {
                WordLoadOptions load = OfficeConverter.WordInputOptions(options);
                using WordDocument document = await WordDocument.LoadAsync(input, load, token).ConfigureAwait(false);
                if (options.Theme != null) document.ApplyTheme(options.Theme);
                WordRtfExportOptions conversion = settings.WordRtf ?? new();
                conversion.FailOnLoss |= options.FailOnLoss;
                conversion.SaveOptions = settings.RtfSave ?? conversion.SaveOptions;
                return await RetainSourceAsync(document.CreateLoadReport(options.SourceName!, TimeSpan.Zero),
                    () => document.SaveAsRtfAsync(output, conversion, token), conversion.FailOnLoss).ConfigureAwait(false);
            }
            if (Id == "pdf-rtf") {
                PdfLoadOptions load = (options.PdfLoad ?? new()).Clone(options.MaximumInputBytes,
                    options.PdfLoad?.Password ?? options.SourcePassword);
                using PdfDocument document = await PdfDocument.LoadAsync(input, load, token).ConfigureAwait(false);
                PdfRtfExportOptions conversion = settings.PdfRtf ?? new();
                conversion.ReadOptions ??= load;
                conversion.FailOnLoss |= options.FailOnLoss;
                return await SaveModelAsync(await document.ToRtfAsync(conversion, token).ConfigureAwait(false),
                    model => model.SaveAsync(output, settings.RtfSave, cancellationToken: token), token).ConfigureAwait(false);
            }
            if (Id == "html-rtf") {
                HtmlConversionDocument document = await HtmlConversionDocument.LoadAsync(input, options.HtmlLoad, cancellationToken: token).ConfigureAwait(false);
                if (settings.HtmlRtfImport != null || settings.WordRtf != null) {
                    HtmlWordExportOptions import = settings.HtmlRtfImport ?? new();
                    if (settings.HtmlRtf != null) throw new ArgumentException("Choose semantic HtmlRtf options or the HtmlRtfImport Word import path.");
                    import.BasePath ??= options.ResourceDirectory;
                    import.FailOnLoss |= options.FailOnLoss;
                    OfficeConversionResult<WordDocument> converted = await document.ToWordAsync(import, token).ConfigureAwait(false);
                    using WordDocument word = converted.Value;
                    WordRtfExportOptions export = settings.WordRtf ?? new();
                    export.FailOnLoss |= options.FailOnLoss;
                    export.SaveOptions = settings.RtfSave ?? export.SaveOptions;
                    return await RetainSourceAsync(converted.Report, () => word.SaveAsRtfAsync(output, export, token),
                        options.FailOnLoss).ConfigureAwait(false);
                }
                HtmlRtfExportOptions conversion = settings.HtmlRtf ?? new();
                conversion.FailOnLoss |= options.FailOnLoss;
                conversion.SaveOptions = settings.RtfSave ?? conversion.SaveOptions;
                return await document.SaveAsRtfAsync(output, conversion, token).ConfigureAwait(false);
            }
            using MarkdownDocument markdown = await LoadMarkdownAsync(input, options, settings, token).ConfigureAwait(false);
            MarkdownRtfExportOptions rtf = settings.MarkdownRtf ?? new();
            rtf.FailOnLoss |= options.FailOnLoss;
            rtf.SaveOptions = settings.RtfSave ?? rtf.SaveOptions;
            return await markdown.SaveAsRtfAsync(output, rtf, token).ConfigureAwait(false);
        }
    }
}
