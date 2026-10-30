using System.Diagnostics;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf.Word;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static readonly HashSet<string> BuiltinRouteIds = new HashSet<string>(new[] {
        "email-pdf", "epub-pdf", "reader-visio-pdf", "doc-pdf", "docx-pdf", "xls-pdf", "xlsb-pdf", "xlsx-pdf", "ppt-pdf", "pptx-pdf", "markdown-pdf", "html-pdf", "txt-pdf",
        "pdf-docx", "pdf-xlsx", "pdf-pptx", "pdf-html", "docx-html", "docx-markdown", "html-docx", "html-markdown", "markdown-docx", "markdown-html"
    }.Concat(OfficeConversionCapabilityCatalog.All.Where(route => route.Id.StartsWith("native-", StringComparison.Ordinal)).Select(route => route.Id)), StringComparer.OrdinalIgnoreCase);

    private static async Task<OfficeConversionReport> ConvertCapturedCoreAsync(Stream input, Stream output,
        OfficeConversionOptions settings, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (IsNativePair(settings)) {
            if (settings.RouteId != null) _ = SelectRoute(settings);
            if (settings.OutputProfile != OfficeConversionProfile.Faithful) throw new ArgumentException("Native Office conversion supports the Faithful output profile.");
            return await ConvertNativeAsync(input, output, settings, token).ConfigureAwait(false);
        }
        ValidateDestinationSaveFormats(settings);
        string id = SelectRoute(settings).Id;
        if (settings.OutputProfile != OfficeConversionProfile.Faithful && id is not ("doc-pdf" or "docx-pdf" or "xls-pdf" or "xlsb-pdf" or "xlsx-pdf" or "ppt-pdf" or "pptx-pdf"))
            throw new ArgumentException("The selected route supports only the Faithful output profile.", nameof(settings));
        if (Registrations.TryGetValue(id, out IOfficeConversionRoute? registration)) {
            OfficeConversionReport registeredReport = await registration.ConvertAsync(input, output, settings, token).ConfigureAwait(false);
            if (settings.FailOnLoss) registeredReport.ThrowIfLoss();
            await ApplyRegisteredSaveOptionsAsync(output, settings, registeredReport, token).ConfigureAwait(false);
            await OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(registeredReport,
                () => ValidateRegisteredOutputAsync(output, settings, token)).ConfigureAwait(false);
            return registeredReport;
        }
        var timer = Stopwatch.StartNew();
        OfficeConversionReport result;
        OfficeConversionReport? sourceReport = null;
        try {
        switch (id) {
            case "email-pdf":
            case "epub-pdf":
            case "reader-visio-pdf":
                result = await ConvertReaderPdfAsync(input, output, settings, token).ConfigureAwait(false);
                break;
            case "doc-pdf":
            case "docx-pdf":
            case "docx-html":
            case "docx-markdown": {
                using WordDocument document = await WordDocument.LoadAsync(input, WordInputOptions(settings), token).ConfigureAwait(false);
                sourceReport = document.CreateLoadReport(settings.SourceName!, timer.Elapsed);
                if (settings.FailOnLoss) sourceReport.ThrowIfLoss();
                if (id is "doc-pdf" or "docx-pdf") {
                    WordPdfExportOptions options = settings.Word ?? new WordPdfExportOptions().UseProfile(ExportProfile(settings.OutputProfile)); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
                } else if (id == "docx-html") {
                    WordHtmlExportOptions options = settings.WordHtml ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsHtmlAsync(output, options, token).ConfigureAwait(false);
                } else {
                    WordMarkdownExportOptions options = settings.WordMarkdown ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsMarkdownAsync(output, options, token).ConfigureAwait(false);
                }
                break;
            }
            case "xls-pdf":
            case "xlsb-pdf":
            case "xlsx-pdf": {
                using ExcelDocument document = await ExcelDocument.LoadAsync(input, ExcelInputOptions(settings), token).ConfigureAwait(false);
                sourceReport = document.CreateLoadReport(settings.SourceName!, timer.Elapsed);
                if (settings.FailOnLoss) sourceReport.ThrowIfLoss();
                ExcelPdfExportOptions options = settings.Excel ?? new ExcelPdfExportOptions().UseProfile(ExportProfile(settings.OutputProfile)); options.FailOnLoss |= settings.FailOnLoss;
                result = await document.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
                break;
            }
            case "ppt-pdf":
            case "pptx-pdf": {
                using PowerPointPresentation document = await PowerPointPresentation.LoadAsync(input, PowerPointInputOptions(settings), token).ConfigureAwait(false);
                sourceReport = document.CreateLoadReport(settings.SourceName!, timer.Elapsed);
                if (settings.FailOnLoss) sourceReport.ThrowIfLoss();
                PowerPointPdfExportOptions options = settings.PowerPoint ?? new PowerPointPdfExportOptions().UseProfile(ExportProfile(settings.OutputProfile)); options.FailOnLoss |= settings.FailOnLoss;
                result = await document.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
                break;
            }
            case "pdf-docx":
            case "pdf-xlsx":
            case "pdf-pptx":
            case "pdf-html": {
                using PdfDocument document = PdfDocument.Load(input, (settings.PdfLoad ?? new()).Clone(settings.MaximumInputBytes, settings.PdfLoad?.Password ?? settings.SourcePassword), token);
                if (id == "pdf-docx") {
                    WordPdfImportOptions options = settings.PdfWord ?? new(); options.ReadOptions ??= settings.PdfLoad; options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToWordAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.WordSave, token)).ConfigureAwait(false);
                } else if (id == "pdf-xlsx") {
                    ExcelPdfImportOptions options = settings.PdfExcel ?? new(); options.ReadOptions ??= settings.PdfLoad; options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToExcelAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.ExcelSave, token)).ConfigureAwait(false);
                } else if (id == "pdf-pptx") {
                    PowerPointPdfImportOptions options = settings.PdfPowerPoint ?? new(); options.ReadOptions ??= settings.PdfLoad; options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToPowerPointAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.PowerPointSave, token)).ConfigureAwait(false);
                } else {
                    PdfToHtmlOptions options = settings.PdfHtml ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    options.MaximumOutputCharacters = (int)Math.Min(options.MaximumOutputCharacters ?? int.MaxValue, settings.MaximumOutputBytes);
                    result = await document.SaveAsHtmlAsync(output, options, token).ConfigureAwait(false);
                }
                break;
            }
            case "markdown-pdf":
            case "markdown-docx":
            case "markdown-html": {
                using MarkdownDocument document = MarkdownDocument.Load(input, new MarkdownLoadOptions { MaxInputBytes = settings.MaximumInputBytes });
                if (id == "markdown-pdf") {
                    MarkdownPdfExportOptions options = settings.Markdown ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    options.BaseDirectory ??= settings.ResourceDirectory;
                    result = await document.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
                } else if (id == "markdown-docx") {
                    WordMarkdownImportOptions options = ConfigureMarkdownHtml(settings.MarkdownWord, token); options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToWordAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.WordSave, token)).ConfigureAwait(false);
                } else {
                    MarkdownHtmlExportOptions options = settings.MarkdownHtml ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsHtmlAsync(output, options, token).ConfigureAwait(false);
                }
                break;
            }
            case "html-docx":
            case "html-markdown":
            case "html-pdf": {
                HtmlConversionDocumentOptions sourceOptions = settings.HtmlLoad ?? new();
                if (sourceOptions.BaseUri == null && settings.ResourceDirectory != null)
                    sourceOptions.BaseUri = new Uri(settings.ResourceSourcePath ?? Path.Combine(settings.ResourceDirectory, settings.SourceName!));
                if (settings.ResourceDirectory != null && settings.HtmlLoad == null) sourceOptions.ResourceUrlPolicy = OfficeHtmlResourceResolver.CreateResourcePolicy();
                var document = await HtmlConversionDocument.LoadAsync(input, sourceOptions, cancellationToken: token).ConfigureAwait(false);
                if (id == "html-docx") {
                    WordHtmlImportOptions options = settings.HtmlWord ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToWordAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.WordSave, token)).ConfigureAwait(false);
                } else if (id == "html-markdown") {
                    MarkdownHtmlImportOptions options = settings.HtmlMarkdown ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsMarkdownAsync(output, options, token).ConfigureAwait(false);
                } else {
                    HtmlToPdfOptions? rendering = settings.ResourceDirectory == null ? settings.Html
                        : OfficeHtmlResourceResolver.CreateOptions(settings.ResourceSourcePath ?? Path.Combine(settings.ResourceDirectory, settings.SourceName!),
                            Math.Max(0L, settings.MaximumInputBytes - input.Length), settings.ResourceSnapshots, settings.Html);
                    PdfDocumentConversionResult converted = await document.ConvertToPdfWithEvidenceAsync(rendering, token).ConfigureAwait(false);
                    using PdfDocument model = converted.Value;
                    PdfSaveResult saved = settings.FailOnLoss
                        ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Html", output, timer, token).ConfigureAwait(false)
                        : await PdfConversionEvidence.RunOutputAsync(converted, "Html", () => converted.SaveAsync(output, token)).ConfigureAwait(false);
                    result = PdfConversionEvidence.CreateReport("Html", saved.ConversionReports, elapsed: timer.Elapsed,
                        itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
                }
                break;
            }
            case "txt-pdf": {
                PdfDocumentConversionResult converted = PdfPlainTextConverter.ToPdfDocumentResult(input, settings.PlainText, settings.MaximumInputBytes, token);
                using PdfDocument model = converted.Value;
                PdfSaveResult saved = settings.FailOnLoss
                    ? await PdfConversionEvidence.SaveLosslessAsync(converted, "Text", output, timer, token).ConfigureAwait(false)
                    : await PdfConversionEvidence.RunOutputAsync(converted, "Text", () => converted.SaveAsync(output, token)).ConfigureAwait(false);
                result = PdfConversionEvidence.CreateReport("Text", saved.ConversionReports, elapsed: timer.Elapsed,
                    itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
                break;
            }
            default: throw new NotSupportedException("Register the owning adapter before executing this route.");
        }
        } catch (OfficeConversionException error) when (sourceReport != null) {
            var report = new OfficeConversionReport(settings.SourceName!, settings.TargetName!,
                elapsed: timer.Elapsed, diagnostics: sourceReport.Diagnostics.Concat(ReferenceEquals(sourceReport, error.Report) ? Array.Empty<OfficeDiagnostic>() :
                    error.Report is OfficeConversionReport shared ? shared.Diagnostics : error.Report.FidelityDiagnostics.Select(
                        d => new OfficeDiagnostic(d.Code, OfficeDiagnosticSeverity.Warning, d.Message, d.Location, d.Source, d.LossKind))));
            throw new OfficeConversionException(error.Message, report, error);
        }
        return new OfficeConversionReport(settings.SourceName!, result.Target, itemsWritten: result.ItemsWritten,
            elapsed: timer.Elapsed, diagnostics: (sourceReport?.Diagnostics ?? Array.Empty<OfficeDiagnostic>()).Concat(result.Diagnostics));
    }

    private static OfficeConversionCapability SelectRoute(OfficeConversionOptions settings) {
        if (settings.RouteId is { } requested) {
            OfficeConversionCapability? exact = Routes.SingleOrDefault(route => string.Equals(route.Id, requested, StringComparison.OrdinalIgnoreCase));
            if (exact == null || !exact.SourceExtensions.Contains(Path.GetExtension(settings.SourceName), StringComparer.OrdinalIgnoreCase)
                || !string.Equals(exact.TargetExtension, Path.GetExtension(settings.TargetName), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The explicit conversion route must match the source and target filenames.");
            return exact;
        }
        string source = Path.GetExtension(settings.SourceName ?? string.Empty), target = Path.GetExtension(settings.TargetName ?? string.Empty);
        OfficeConversionCapability[] matches = Routes.Where(route => route.SourceExtensions.Contains(source, StringComparer.OrdinalIgnoreCase)
            && string.Equals(route.TargetExtension, target, StringComparison.OrdinalIgnoreCase)).ToArray();
        string? preferred = OfficeConversionCapabilityCatalog.PreferredRouteId(source, target);
        if (preferred != null && matches.FirstOrDefault(route => route.Id == preferred) is { } defaultRoute) return defaultRoute;
        return matches.Length == 1 ? matches[0]
            : throw new NotSupportedException("No unique executable route matches the input and output. Register the owning adapter or choose RouteId for ambiguous formats.");
    }

    private static PdfExportProfile ExportProfile(OfficeConversionProfile profile) => profile switch {
        OfficeConversionProfile.Faithful => PdfExportProfile.Faithful,
        OfficeConversionProfile.Lightweight => PdfExportProfile.Lightweight,
        OfficeConversionProfile.PrintReady => PdfExportProfile.PrintReady,
        OfficeConversionProfile.TextOnly => PdfExportProfile.TextOnly,
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    private static Task WriteTextAsync(Stream stream, string text, CancellationToken token) =>
        OfficeRuntimeCompatibility.WriteAllAsync(stream, Encoding.UTF8.GetBytes(text), token);
}
