using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;

namespace OfficeIMO.OneNote.Html;

public static partial class HtmlOneNoteConverterExtensions {
    /// <summary>Converts prepared HTML into a OneNote section.</summary>
    public static OneNoteSection ToOneNoteSection(this HtmlConversionDocument document, HtmlOneNoteExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToOneNoteSection(options, out _, cancellationToken);

    /// <summary>Converts prepared HTML into a OneNote section and returns the completed conversion report.</summary>
    public static OneNoteSection ToOneNoteSection(this HtmlConversionDocument document, HtmlOneNoteExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var result = document.ConvertToOneNoteSectionWithEvidence(options);
        OneNoteSection value = Require(result);
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("Html", "OneNote", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts prepared HTML into a OneNote section, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<OneNoteSection>> ToOneNoteSectionAsync(this HtmlConversionDocument document, HtmlOneNoteExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToOneNoteSection(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Converts prepared HTML into a single-section OneNote notebook.</summary>
    public static OneNoteNotebook ToOneNoteNotebook(this HtmlConversionDocument document, HtmlOneNoteExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToOneNoteNotebook(options, out _, cancellationToken);

    /// <summary>Converts prepared HTML into a single-section OneNote notebook and returns the completed conversion report.</summary>
    public static OneNoteNotebook ToOneNoteNotebook(this HtmlConversionDocument document, HtmlOneNoteExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var result = document.ConvertToOneNoteNotebookWithEvidence(options);
        OneNoteNotebook value = Require(result);
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("Html", "OneNote", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts prepared HTML into a single-section OneNote notebook, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<OneNoteNotebook>> ToOneNoteNotebookAsync(this HtmlConversionDocument document, HtmlOneNoteExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToOneNoteNotebook(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

}
