using OfficeIMO.Core.Internal;
using OfficeIMO.Markdown;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Workflows;

/// <summary>Projects Markdown outlines to native presentations through the existing Markup conversion engine.</summary>
public static class MarkdownPowerPointConverterExtensions {
    /// <summary>Converts an outline to a detached native presentation.</summary>
    public static PowerPointDocument ToPowerPoint(this MarkdownDocument document, MarkdownToPowerPointOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToPowerPoint(options, out _, cancellationToken);

    /// <summary>Converts an outline and returns its complete projection and native renderer findings.</summary>
    public static PowerPointDocument ToPowerPoint(this MarkdownDocument document, MarkdownToPowerPointOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        OfficeConversionResult<PowerPointDocument> converted = document.ToPowerPointAsync(options, cancellationToken).GetAwaiter().GetResult();
        report = converted.Report;
        return converted.Value;
    }

    /// <summary>Converts an outline, asynchronously capturing permitted images before native presentation rendering.</summary>
    public static async Task<OfficeConversionResult<PowerPointDocument>> ToPowerPointAsync(this MarkdownDocument document,
        MarkdownToPowerPointOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        MarkdownToPowerPointOptions settings = (options ?? new()).Clone();
        var findings = new List<OfficeDiagnostic>();
        OfficeMarkupDocument markup = MarkdownPowerPointProjection.Build(document, settings, findings, cancellationToken);
        if (findings.Any(finding => finding.Severity == OfficeDiagnosticSeverity.Error))
            throw new OfficeConversionException("The Markdown outline contains invalid presentation markup.",
                new OfficeConversionReport("Markdown", "PowerPoint", diagnostics: findings));
        string root = OfficeTemporaryDirectory.Create("officeimo-markdown-deck-");
        try {
            await MarkdownPowerPointProjection.PrepareImagesAsync(markup.Blocks, settings, root, new Dictionary<string, string>(), findings,
                cancellationToken).ConfigureAwait(false);
            var projection = new OfficeConversionReport("Markdown", "PowerPoint", diagnostics: findings);
            if (settings.FailOnLoss) projection.ThrowIfLoss();
            OfficeConversionResult<PowerPointDocument> converted;
            var export = MarkdownPowerPointProjection.ExportOptions(settings, root);
            export.Theme = settings.Theme;
            try { converted = await markup.ToPowerPointAsync(export, cancellationToken).ConfigureAwait(false); }
            catch (OfficeConversionException error) {
                throw new OfficeConversionException(error.Message, new OfficeConversionReport("Markdown", "PowerPoint",
                    diagnostics: OfficeConversionFidelityDiagnostics.Flatten(new[] { projection, error.Report })), error);
            }
            var report = new OfficeConversionReport("Markdown", "PowerPoint", itemsWritten: converted.Value.Slides.Count,
                diagnostics: OfficeConversionFidelityDiagnostics.Flatten(new[] { projection, converted.Report }));
            try { cancellationToken.ThrowIfCancellationRequested(); if (settings.FailOnLoss) report.ThrowIfLoss(); }
            catch { converted.Value.Dispose(); throw; }
            return new OfficeConversionResult<PowerPointDocument>(converted.Value, report);
        } finally { Directory.Delete(root, recursive: true); }
    }

    /// <summary>Converts an outline and saves a native presentation without replacing existing output.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this MarkdownDocument document, string path,
        MarkdownToPowerPointOptions? options = null, CancellationToken cancellationToken = default) =>
        document.SaveAsPowerPointAsync(path, options, cancellationToken).GetAwaiter().GetResult();

    /// <summary>Converts an outline and writes native presentation bytes to a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this MarkdownDocument document, Stream stream,
        MarkdownToPowerPointOptions? options = null, CancellationToken cancellationToken = default) =>
        document.SaveAsPowerPointAsync(stream, options, cancellationToken).GetAwaiter().GetResult();

    /// <summary>Converts an outline and atomically saves it without replacing existing output.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this MarkdownDocument document, string path,
        MarkdownToPowerPointOptions? options = null, CancellationToken cancellationToken = default) {
        var converted = await document.ToPowerPointAsync(options, cancellationToken).ConfigureAwait(false);
        using PowerPointDocument value = converted.Value;
        await OfficeConversionFailures.RunAsync(converted.Report, () => OfficeFileCommit.WriteAsync(path,
            (output, token) => value.SaveAsync(output, token), OfficeFileCommit.ConflictPolicy.FailIfExists, cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(converted.Report.Source, converted.Report.Target, path, converted.Report.ItemsWritten,
            converted.Report.Elapsed, converted.Report.Diagnostics);
    }

    /// <summary>Converts an outline and writes native presentation bytes, leaving the caller stream open.</summary>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this MarkdownDocument document, Stream stream,
        MarkdownToPowerPointOptions? options = null, CancellationToken cancellationToken = default) {
        var converted = await document.ToPowerPointAsync(options, cancellationToken).ConfigureAwait(false);
        using PowerPointDocument value = converted.Value;
        await OfficeConversionFailures.RunAsync(converted.Report, () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return converted.Report;
    }
}
