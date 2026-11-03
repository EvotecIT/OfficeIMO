using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <summary>
/// Provides extension methods for converting between <see cref="RtfDocument"/> and semantic HTML.
/// </summary>
public static partial class HtmlRtfConverterExtensions {

    private static string ToHtmlCore(RtfDocument document, RtfHtmlExportOptions effectiveOptions) {
        ArgumentNullException.ThrowIfNull(document);

        HtmlConversionProfileContracts.Get(effectiveOptions.SharedProfile);
        effectiveOptions.DocumentOutput.Validate();

        if (effectiveOptions.PreferEncapsulatedHtml &&
            document.IsHtmlEncapsulationCurrent &&
            document.HtmlEncapsulation != null &&
            !string.IsNullOrWhiteSpace(document.HtmlEncapsulation.Html)) {
            var importOptions = HtmlRtfExportOptions.UntrustedDefaults;
            importOptions.UrlPolicy = effectiveOptions.GetUrlPolicy().Clone();
            HtmlConversionDocument encapsulatedHtml = HtmlConversionDocument.Parse(
                document.HtmlEncapsulation.Html,
                new HtmlConversionDocumentOptions {
                    Profile = HtmlConversionProfile.Document,
                    Trust = HtmlInputTrust.Untrusted
                });
            effectiveOptions.HtmlDiagnostics.AddRange(encapsulatedHtml.ResourceManifest.Diagnostics);
            HtmlToRtfResult imported = encapsulatedHtml.ConvertToRtfWithEvidence(importOptions);
            foreach (HtmlRtfConversionDiagnostic diagnostic in imported.RtfDiagnostics) {
                effectiveOptions.AddDiagnostic(diagnostic.Code, diagnostic.Message, diagnostic.Source, severity: diagnostic.Severity, action: diagnostic.Action);
            }

            effectiveOptions.AddDiagnostic(
                "RtfHtmlEncapsulatedHtmlUsed",
                "Outlook/Exchange encapsulated HTML was used instead of the RTF plain-text fallback.",
                severity: HtmlRtfConversionDiagnosticSeverity.Info);
            return RtfHtmlWriter.Write(imported.RequireValue(), effectiveOptions);
        }

        if (effectiveOptions.PreferEncapsulatedHtml && document.HtmlEncapsulation != null) {
            effectiveOptions.AddDiagnostic("RtfHtmlEncapsulationInvalidated",
                "Semantic edits invalidated the original encapsulated HTML; current RTF content was converted.",
                severity: HtmlRtfConversionDiagnosticSeverity.Warning, action: RtfConversionAction.Omitted);
        }
        effectiveOptions.ConversionReport.Merge(document.GetStyleConversionDiagnostics());
        return RtfHtmlWriter.Write(document, effectiveOptions);
    }
}
