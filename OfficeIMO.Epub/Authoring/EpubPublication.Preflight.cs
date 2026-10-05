using System.Threading;
using System.Xml;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Checks the current publication using the selected save policy and inspects all retained XHTML/SVG
    /// identifiers and image alternatives. Schema conformance, comprehensive accessibility and reader
    /// presentation remain explicitly unchecked. No file is written and no external process is invoked.
    /// </summary>
    public EpubPreflightReport Preflight(EpubWriteOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var checks = new List<EpubPreflightCheck>();
        var package = new List<EpubDiagnostic>();
        try {
            EpubWriteReport report = Write(options, cancellationToken).Report;
            foreach (var diagnostic in report.FidelityDiagnostics) Add(package, diagnostic.Code,
                EpubDiagnosticSeverity.Warning, diagnostic.Message, PackagePath);
        } catch (Exception error) when (IsPreflightFailure(error)) {
            Add(package, "EPUB_PREFLIGHT_WRITE_FAILED", EpubDiagnosticSeverity.Error, error.Message, PackagePath);
        }
        checks.Add(Result("native-save", package));
        checks.Add(CheckAccessibilityDiscovery());

        var contentFindings = new List<EpubDiagnostic>();
        var imageFindings = new List<EpubDiagnostic>();
        foreach (var group in Manifest.Where(item => HasMediaType(item.MediaType, "application/xhtml+xml") || HasMediaType(item.MediaType, "image/svg+xml"))
            .GroupBy(item => item.Reference.ContainerPath ?? item.Href, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            EpubManifestItem item = group.First();
            string path = group.Key;
            try {
                if (item.Reference.Kind != EpubReferenceKind.Container || _encryption.Any(entry => entry.Path == path && entry.RequiresDecryption))
                    throw new NotSupportedException("Content requires external access or decryption; this preflight cannot inspect it.");
                XDocument document = GetContentXml(item.Id);
                HashSet<string> ids = EpubContentIdentifiers.Collect(document.Root!, path, true, cancellationToken);
                EpubContentIdentifiers.ValidateReferences(document.Root!, ids, path, cancellationToken);
                foreach (XElement img in document.Descendants(Html + "img")) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (img.Attribute("alt") == null) Add(imageFindings, "EPUB_PREFLIGHT_IMAGE_ALT_MISSING",
                        EpubDiagnosticSeverity.Error, "Image " + ((string?)img.Attribute("id") ?? (string?)img.Attribute("src") ?? "(unnamed)") +
                        " needs an alt attribute. Supply a meaningful alternative, or an empty alternative for a decorative image.", path);
                }
            } catch (Exception error) when (IsPreflightFailure(error)) {
                Add(contentFindings, "EPUB_PREFLIGHT_CONTENT_INVALID", EpubDiagnosticSeverity.Error, error.Message, path);
                Add(imageFindings, "EPUB_PREFLIGHT_IMAGE_CHECK_INCOMPLETE", EpubDiagnosticSeverity.Error,
                    "Image alternatives could not be fully checked because content inspection failed.", path);
            }
        }
        checks.Add(Result("content-identifiers", contentFindings));
        checks.Add(Result("image-alternative-presence", imageFindings));
        foreach (string scope in new[] { "epub-schema-conformance", "accessibility-assessment", "reading-system-presentation" })
            checks.Add(new EpubPreflightCheck(scope, EpubPreflightStatus.NotChecked, Array.Empty<EpubDiagnostic>()));
        return new EpubPreflightReport(checks);
    }

    private static bool IsPreflightFailure(Exception error) => error is InvalidDataException || error is NotSupportedException ||
        error is InvalidOperationException || error is XmlException || error is ArgumentException;

    private static EpubPreflightCheck Result(string code, List<EpubDiagnostic> findings) => new EpubPreflightCheck(code,
        findings.Any(item => item.Severity == EpubDiagnosticSeverity.Error) ? EpubPreflightStatus.Failed : EpubPreflightStatus.Passed, findings);

    private static void Add(List<EpubDiagnostic> findings, string code, EpubDiagnosticSeverity severity, string message, string? path) {
        const int maximum = 10_000;
        if (findings.Count >= maximum) return;
        if (findings.Count == maximum - 1) findings.Add(new EpubDiagnostic { Code = "EPUB_PREFLIGHT_FINDING_LIMIT",
            Severity = EpubDiagnosticSeverity.Error, Message = "Preflight reached its 10,000-finding bound. Repair findings and rerun to inspect the remainder." });
        else findings.Add(new EpubDiagnostic { Code = code, Severity = severity, Message = message, Path = path });
    }
}
