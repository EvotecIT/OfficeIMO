using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Markdown;

internal partial class WordToMarkdownConverter {
    private static void ReportOmission(WordMarkdownExportOptions options, string message, string location) =>
        options.DiagnosticSink?.Invoke(new WordMarkdownConversionDiagnostic(OfficeDiagnosticCodes.WordToMarkdownWarning,
            message, OfficeConversionLossKind.Omission, location: location, format: "Word"));

    private static void ReportOmittedRunningContent(WordSection section, int index, WordMarkdownExportOptions options, CancellationToken token) {
        Check(section.PageSetup.GetHeader(WordHeaderFooterType.Default), "Header/default"); Check(section.PageSetup.GetHeader(WordHeaderFooterType.First), "Header/first"); Check(section.PageSetup.GetHeader(WordHeaderFooterType.Even), "Header/even");
        Check(section.PageSetup.GetFooter(WordHeaderFooterType.Default), "Footer/default"); Check(section.PageSetup.GetFooter(WordHeaderFooterType.First), "Footer/first"); Check(section.PageSetup.GetFooter(WordHeaderFooterType.Even), "Footer/even");
        void Check(WordHeaderFooter? story, string slot) {
            OfficeXmlNode root = story?.Node() ?? default;
            if (root.IsNone) return;
            foreach (OfficeXmlNode item in root.DescendantElements()) {
                token.ThrowIfCancellationRequested();
                if (item.IsW("t") && !string.IsNullOrWhiteSpace(item.InnerText)
                    || IsOmittedRunningContent(item)) {
                    ReportOmission(options, "Word running content was omitted from Markdown because semantic headers and footers are disabled.",
                        "Section " + (index + 1) + "/" + slot);
                    return;
                }
            }
        }
    }

    private static bool IsOmittedRunningContent(OfficeXmlNode item) {
        if (item.Is(WordXml.M, "oMath")) return true;
        if (item.NamespaceId != WordXml.W) return false;
        switch (item.LocalName) {
            case "drawing": case "pict": case "instrText": case "fldChar": case "fldSimple": case "tbl": case "altChunk":
                return true;
            default:
                return false;
        }
    }
}
