using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static void CheckDescendantLanguages(XElement root, string path, List<EpubDiagnostic> findings, CancellationToken token) {
        foreach (XElement element in root.Descendants()) {
            token.ThrowIfCancellationRequested();
            // Bare lang belongs to XHTML/SVG. xml:lang applies across embedded XML vocabularies.
            string? language = element.Name.Namespace == Html || element.Name.NamespaceName == "http://www.w3.org/2000/svg" ?
                (string?)element.Attribute("lang") : null;
            string? xmlLanguage = (string?)element.Attribute(XNamespace.Xml + "lang");
            if (language == null && xmlLanguage == null) continue;
            bool invalid = language != null && language.Length != 0 && !EpubLanguageTag.IsWellFormed(language) ||
                xmlLanguage != null && xmlLanguage.Length != 0 && !EpubLanguageTag.IsWellFormed(xmlLanguage);
            bool conflict = language != null && xmlLanguage != null && !string.Equals(language, xmlLanguage, StringComparison.OrdinalIgnoreCase);
            if (!invalid && !conflict) continue;
            string source = (string?)element.Attribute("id") ?? (string?)element.Attribute(XNamespace.Xml + "id") ?? element.Name.LocalName;
            if (source.Length > 128) source = source.Substring(0, 128) + "…";
            if (invalid) Add(findings, "EPUB_PREFLIGHT_CONTENT_LANGUAGE_INVALID", EpubDiagnosticSeverity.Error,
                "Element '" + source + "' needs a well-formed BCP 47 language tag or an empty language declaration for unknown language.", path);
            if (conflict) Add(findings, "EPUB_PREFLIGHT_CONTENT_LANGUAGE_CONFLICT", EpubDiagnosticSeverity.Error,
                "Element '" + source + "' has conflicting lang and xml:lang declarations; use the same language or empty value in both.", path);
        }
    }
}
