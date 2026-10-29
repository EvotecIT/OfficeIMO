using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    // EPUB 3.3 metadata link vocabulary: https://www.w3.org/TR/epub-33/#sec-link-vocab
    private EpubPreflightCheck CheckMetadataLinks(IReadOnlyDictionary<string, string> prefixes,
        CancellationToken cancellationToken) {
        const string vocabulary = "http://idpf.org/epub/vocab/package/link/#";
        var findings = new List<EpubDiagnostic>();
        foreach (XElement metadata in Root.Descendants(Opf + "metadata").Where(element =>
            element.Parent == Root || element.Parent?.Name == Opf + "collection")) {
            foreach (XElement link in metadata.Elements(Opf + "link")) {
                cancellationToken.ThrowIfCancellationRequested();
                string[] tokens = Tokens((string?)link.Attribute("rel"));
                bool alternate = false, record = false, voicing = false, combined = false;
                string? first = tokens.Length == 0 ? null : tokens[0];
                foreach (string token in tokens) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (token != first) combined = true;
                    if (token.IndexOf(':') < 0) {
                        alternate |= token == "alternate";
                        record |= token == "record";
                        voicing |= token == "voicing";
                    } else {
                        alternate |= EpubVocabulary.Matches(prefixes, token, vocabulary, "alternate");
                        record |= EpubVocabulary.Matches(prefixes, token, vocabulary, "record");
                        voicing |= EpubVocabulary.Matches(prefixes, token, vocabulary, "voicing");
                    }
                }
                string context = (string?)link.Attribute("id") ?? (string?)link.Attribute("href") ?? "(unnamed)";
                if (context.Length > 128) context = context.Substring(0, 128);
                string label = "Metadata link '" + context + "': ";
                if ((alternate || record) && link.Attribute("refines") != null)
                    Add(findings, "EPUB_PREFLIGHT_LINK_REFINES_FORBIDDEN", EpubDiagnosticSeverity.Error,
                        label + "alternate and record relationships apply to the publication or collection and must not have refines.", PackagePath);
                if (voicing && string.IsNullOrWhiteSpace((string?)link.Attribute("refines")))
                    Add(findings, "EPUB_PREFLIGHT_LINK_REFINES_REQUIRED", EpubDiagnosticSeverity.Error,
                        label + "a voicing relationship requires refines to identify its expression or resource.", PackagePath);
                if ((record || voicing) && string.IsNullOrWhiteSpace((string?)link.Attribute("media-type")))
                    Add(findings, "EPUB_PREFLIGHT_LINK_MEDIA_TYPE_MISSING", EpubDiagnosticSeverity.Error,
                        label + "record and voicing relationships require media-type, including external resources.", PackagePath);
                if (alternate && combined)
                    Add(findings, "EPUB_PREFLIGHT_LINK_ALTERNATE_COMBINED", EpubDiagnosticSeverity.Error,
                        label + "alternate must not be combined with another relationship keyword.", PackagePath);
            }
        }
        return Result("metadata-links", findings);
    }
}
