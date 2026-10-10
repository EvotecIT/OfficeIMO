using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

public static partial class PdfProvenance {
    private static int ManifestRevision(PdfReadDocument document, int streamNumber) {
        return document.Objects.TryGetValue(streamNumber, out var stream) ? stream.SourceRevision : -1;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> SelectActiveManifest(
        List<(OfficeProvenanceEvidence Evidence, int Revision, int Stream, OfficeC2paManifestSummary? Summary)> stores) {
        var diagnostics = new List<string>();
        var eligible = new List<(OfficeProvenanceEvidence Evidence, OfficeC2paManifestSummary? Summary)>();
        foreach (var revision in stores.GroupBy(store => store.Revision).OrderByDescending(group => group.Key)) {
            if (revision.Key < 0) {
                diagnostics.Add("A C2PA store's PDF update order could not be established; no active summary was selected from it.");
                continue;
            }
            var unique = revision.GroupBy(store => store.Stream).Select(group => group.First()).ToArray();
            if (unique.Length != 1) {
                diagnostics.Add($"PDF revision {revision.Key} contains multiple C2PA stores; no active summary was selected from that revision.");
                continue;
            }
            eligible.Add((unique[0].Evidence, unique[0].Summary));
        }
        if (eligible.Count > 0 && eligible[0].Summary is { } active) {
            int count = eligible.Sum(store => store.Summary?.ManifestCount ?? 0);
            eligible[0].Evidence.WithManifest(new OfficeC2paManifestSummary(active.Label, active.ClaimGenerator,
                active.Title, active.Format, active.Actions, active.Ingredients, active.SignedBy,
                active.CertificateIssuer, count, active.DeclaresGenerativeAi));
        }
        return diagnostics.AsReadOnly();
    }
}
