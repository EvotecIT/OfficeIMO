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
        "word-xml-docx", "spreadsheet-xml-xlsx", "pdf-docx", "pdf-xlsx", "pdf-pptx", "pdf-html", "docx-html", "pptx-html", "docx-markdown", "html-docx", "html-markdown", "markdown-docx", "markdown-html"
    }.Concat(OfficeConversionCapabilityCatalog.All.Where(route => route.Id.StartsWith("native-", StringComparison.Ordinal)).Select(route => route.Id))
        .Concat(DataRouteIds), StringComparer.OrdinalIgnoreCase);

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
        ValidateThemeRoute(settings, id);
        if (settings.OutputProfile != OfficeConversionProfile.Faithful && id is not ("doc-pdf" or "docx-pdf" or "xls-pdf" or "xlsb-pdf" or "xlsx-pdf" or "ppt-pdf" or "pptx-pdf")
            && !(id == "email-pdf" && settings.OutputProfile == OfficeConversionProfile.TextOnly))
            throw new ArgumentException("The selected route supports only the Faithful output profile.", nameof(settings));
        if (Registrations.TryGetValue(id, out IOfficeConversionRoute? registration)) {
            OfficeConversionReport registeredReport = await registration.ConvertAsync(input, output, settings, token).ConfigureAwait(false);
            if (settings.FailOnLoss) registeredReport.ThrowIfLoss();
            await ApplyRegisteredSaveOptionsAsync(output, settings, registeredReport, token).ConfigureAwait(false);
            await OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(registeredReport,
                () => ValidateRegisteredArtifactAsync(registration, output, settings, token)).ConfigureAwait(false);
            return registeredReport;
        }
        var timer = Stopwatch.StartNew();
        OfficeConversionReport result;
        OfficeConversionReport? sourceReport = null;
        try {
        switch (id) {
            case "email-pdf" when settings.OutputProfile != OfficeConversionProfile.TextOnly:
                result = await ConvertEmailPdfAsync(input, output, settings, token).ConfigureAwait(false);
                break;
            case "email-pdf":
            case "epub-pdf":
            case "reader-visio-pdf":
                result = await ConvertReaderPdfAsync(input, output, settings, token).ConfigureAwait(false);
                break;
            case "word-xml-docx": {
                using WordDocument document = await WordDocument.LoadAsync(input, WordInputOptions(settings), token).ConfigureAwait(false);
                if (document.SourceFormat is not (WordFileFormat.FlatOpc or WordFileFormat.WordMl2003) ||
                    document.OpenXmlDocument.DocumentType != DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
                    throw new NotSupportedException("This route requires Word 2003 XML or a plain Word Flat OPC document.");
                sourceReport = document.CreateLoadReport(settings.SourceName!, timer.Elapsed);
                if (settings.FailOnLoss) sourceReport.ThrowIfLoss();
                if (settings.Theme != null) document.ApplyTheme(settings.Theme);
                await document.SaveAsync(output, settings.WordSave, token).ConfigureAwait(false);
                result = new OfficeConversionReport(settings.SourceName!, settings.TargetName!, itemsWritten: 1);
                break;
            }
            case "spreadsheet-xml-xlsx": {
                var xmlOptions = new ExcelSpreadsheetXmlOptions { CancellationToken = token };
                xmlOptions.MaxInputBytes = (int)Math.Min(xmlOptions.MaxInputBytes, settings.MaximumInputBytes);
                var converted = ExcelDocument.ImportSpreadsheetXml(input, xmlOptions);
                using ExcelDocument document = converted.Value;
                sourceReport = converted.Report;
                if (settings.FailOnLoss) sourceReport.ThrowIfLoss();
                await document.SaveAsync(output, settings.ExcelSave, token).ConfigureAwait(false);
                result = new OfficeConversionReport(settings.SourceName!, settings.TargetName!, itemsWritten: 1);
                break;
            }
            case "doc-pdf":
            case "docx-pdf":
            case "docx-html":
            case "docx-markdown": {
                using WordDocument document = await WordDocument.LoadAsync(input, WordInputOptions(settings), token).ConfigureAwait(false);
                sourceReport = document.CreateLoadReport(settings.SourceName!, timer.Elapsed);
                if (settings.FailOnLoss) sourceReport.ThrowIfLoss();
                if (settings.Theme != null) document.ApplyTheme(settings.Theme);
                if (id is "doc-pdf" or "docx-pdf") {
                    WordPdfExportOptions options = settings.Word ?? new WordPdfExportOptions().UseProfile(ExportProfile(settings.OutputProfile)); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
                } else if (id == "docx-html") {
                    WordHtmlExportOptions options = settings.WordHtml ?? (settings.Theme == null ? new() : WordHtmlExportOptions.CreateSemanticDocumentProfile()); options.FailOnLoss |= settings.FailOnLoss;
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
            case "pptx-pdf":
            case "pptx-html": {
                PowerPointLoadOptions loadOptions = PowerPointInputOptions(settings);
                // A PPTX to PDF route only reads the presentation. Opening it read-only keeps
                // disposal from rewriting every package entry the conversion read. Caller-supplied
                // load options keep the access they asked for.
                if (id == "pptx-pdf" && settings.PowerPointLoad == null) loadOptions.AccessMode = DocumentAccessMode.ReadOnly;
                using PowerPointDocument document = await PowerPointDocument.LoadAsync(input, loadOptions, token).ConfigureAwait(false);
                sourceReport = document.CreateLoadReport(settings.SourceName!, timer.Elapsed);
                if (settings.FailOnLoss) sourceReport.ThrowIfLoss();
                if (id == "pptx-html") result = await ConvertPowerPointHtmlAsync(document, output, settings, token).ConfigureAwait(false);
                else {
                    PowerPointPdfExportOptions options = settings.PowerPoint ?? new PowerPointPdfExportOptions().UseProfile(ExportProfile(settings.OutputProfile)); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
                }
                break;
            }
            case "pdf-docx":
            case "pdf-xlsx":
            case "pdf-pptx":
            case "pdf-html": {
                using PdfDocument document = PdfDocument.Load(input, (settings.PdfLoad ?? new()).Clone(settings.MaximumInputBytes, settings.PdfLoad?.Password ?? settings.SourcePassword), token);
                if (id == "pdf-docx") {
                    PdfWordExportOptions options = settings.PdfWord ?? new(); options.ReadOptions ??= settings.PdfLoad; options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToWordAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.WordSave, token)).ConfigureAwait(false);
                } else if (id == "pdf-xlsx") {
                    PdfExcelExportOptions options = settings.PdfExcel ?? new(); options.ReadOptions ??= settings.PdfLoad; options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToExcelAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.ExcelSave, token)).ConfigureAwait(false);
                } else if (id == "pdf-pptx") {
                    PdfPowerPointExportOptions options = settings.PdfPowerPoint ?? new(); options.ReadOptions ??= settings.PdfLoad; options.FailOnLoss |= settings.FailOnLoss;
                    result = await SaveConvertedAsync(await document.ToPowerPointAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.PowerPointSave, token)).ConfigureAwait(false);
                } else {
                    PdfHtmlExportOptions options = settings.PdfHtml ?? new(); options.ReadOptions ??= settings.PdfLoad; options.FailOnLoss |= settings.FailOnLoss;
                    options.MaximumOutputCharacters = (int)Math.Min(options.MaximumOutputCharacters ?? int.MaxValue, settings.MaximumOutputBytes);
                    result = await document.SaveAsHtmlAsync(output, options, token).ConfigureAwait(false);
                }
                break;
            }
            case "markdown-pdf":
            case "markdown-docx":
            case "markdown-html": {
                MarkdownLoadOptions load = id == "markdown-html" ? settings.MarkdownHtml?.ReaderOptions?.Clone() ?? new()
                    : id == "markdown-docx" ? (settings.MarkdownWord ?? new()).CreateReaderOptions() : new();
                load.MaxInputBytes = Math.Min(load.MaxInputBytes ?? settings.MaximumInputBytes, settings.MaximumInputBytes);
                // The document is converted once and disposed. HTML renderers or extensions supplied by the
                // caller receive parsed blocks with their source spans, so those get the syntax tree.
                load.CaptureSyntaxTree = id == "markdown-html" && settings.MarkdownHtml?.HasCallerRenderHooks == true;
                using MarkdownDocument document = await MarkdownDocument.LoadAsync(input, load, token).ConfigureAwait(false);
                if (id == "markdown-pdf") {
                    MarkdownPdfExportOptions options = settings.Markdown ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    if (settings.Theme != null) options.Theme ??= MarkdownVisualTheme.FromOfficeTheme(settings.Theme);
                    options.BaseDirectory ??= settings.ResourceDirectory;
                    result = await document.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
                } else if (id == "markdown-docx") {
                    MarkdownWordExportOptions options = ConfigureMarkdownHtml(settings.MarkdownWord, token); options.FailOnLoss |= settings.FailOnLoss;
                    if (settings.Theme != null) options.Theme ??= MarkdownVisualTheme.FromOfficeTheme(settings.Theme);
                    options.SourceDirectory ??= settings.ResourceDirectory;
                    result = await SaveConvertedAsync(await document.ToWordAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.WordSave, token), options.Template).ConfigureAwait(false);
                } else {
                    MarkdownHtmlExportOptions options = ConfigureThemedMarkdownHtml(settings); options.FailOnLoss |= settings.FailOnLoss;
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
                if (settings.Theme != null && id == "html-pdf") document = document.WithTheme(settings.Theme);
                if (id == "html-docx") {
                    HtmlWordExportOptions options = settings.HtmlWord ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    options.Theme ??= settings.Theme;
                    result = await SaveConvertedAsync(await document.ToWordAsync(options, token).ConfigureAwait(false),
                        model => model.SaveAsync(output, settings.WordSave, token)).ConfigureAwait(false);
                } else if (id == "html-markdown") {
                    HtmlMarkdownExportOptions options = settings.HtmlMarkdown ?? new(); options.FailOnLoss |= settings.FailOnLoss;
                    result = await document.SaveAsMarkdownAsync(output, options, token).ConfigureAwait(false);
                } else {
                    HtmlPdfExportOptions? rendering = settings.ResourceDirectory == null ? settings.Html
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
            default:
                if (!IsDataRoute(id)) throw new NotSupportedException("Register the owning adapter before executing this route.");
                result = ConvertData(id, input, output, settings, token);
                break;
        }
        } catch (OfficeConversionException error) when (sourceReport != null) {
            var report = new OfficeConversionReport(settings.SourceName!, settings.TargetName!,
                elapsed: timer.Elapsed, diagnostics: sourceReport.Diagnostics.Concat(ReferenceEquals(sourceReport, error.Report) ? Array.Empty<OfficeDiagnostic>() :
                    error.Report.Diagnostics));
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
