using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Markdown;

internal partial class WordToMarkdownConverter {
    private static void ReportOmission(WordMarkdownExportOptions options, string message, string location) =>
        options.DiagnosticSink?.Invoke(new WordMarkdownConversionDiagnostic(OfficeDiagnosticCodes.WordToMarkdownWarning,
            message, OfficeConversionLossKind.Omission, location: location, format: "Word"));

    private static void ReportOmittedRunningContent(WordSection section, int index, WordMarkdownExportOptions options, CancellationToken token) {
        Check(section.Header.Default, "Header/default"); Check(section.Header.First, "Header/first"); Check(section.Header.Even, "Header/even");
        Check(section.Footer.Default, "Footer/default"); Check(section.Footer.First, "Footer/first"); Check(section.Footer.Even, "Footer/even");
        void Check(WordHeaderFooter? story, string slot) {
            OpenXmlCompositeElement? root = (OpenXmlCompositeElement?)story?._header ?? story?._footer;
            if (root == null) return;
            foreach (OpenXmlElement item in root.Descendants()) {
                token.ThrowIfCancellationRequested();
                if (item is Text text && !string.IsNullOrWhiteSpace(text.Text)
                    || item is DocumentFormat.OpenXml.Wordprocessing.Drawing or Picture or FieldCode or FieldChar or SimpleField or Table or AltChunk
                    || item is DocumentFormat.OpenXml.Math.OfficeMath) {
                    ReportOmission(options, "Word running content was omitted from Markdown because semantic headers and footers are disabled.",
                        "Section " + (index + 1) + "/" + slot);
                    return;
                }
            }
        }
    }
}
