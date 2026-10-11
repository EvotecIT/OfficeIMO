using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

public static partial class PdfProvenance {
    private sealed class ManifestOccurrence(OfficeProvenanceEvidence evidence, int revision, int stream, OfficeC2paManifestSummary? summary) {
        internal OfficeProvenanceEvidence Evidence { get; } = evidence;
        internal int Revision { get; set; } = revision;
        internal int LastSnapshot { get; set; } = revision;
        internal int Stream { get; } = stream;
        internal OfficeC2paManifestSummary? Summary { get; } = summary;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> SelectActiveManifest(
        List<ManifestOccurrence> stores) {
        var diagnostics = new List<string>();
        var eligible = new List<(OfficeProvenanceEvidence Evidence, OfficeC2paManifestSummary? Summary)>();
        // An undated store may be newer than every dated store. Do not assert an active claim.
        bool canSelect = !stores.Any(store => store.Revision < 0);
        foreach (var revision in stores.GroupBy(store => store.Revision).OrderByDescending(group => group.Key)) {
            if (revision.Key < 0) {
                diagnostics.Add("A C2PA store's PDF update order could not be established; no active summary was selected from it.");
                continue;
            }
            var unique = revision.GroupBy(store => store.Stream)
                .Select(group => group.OrderByDescending(store => store.Summary != null).First()).ToArray();
            if (unique.Length != 1) {
                diagnostics.Add($"PDF revision {revision.Key} contains multiple C2PA stores; no active summary was selected from that revision.");
                if (eligible.Count == 0) canSelect = false;
                continue;
            }
            eligible.Add((unique[0].Evidence, unique[0].Summary));
            if (eligible.Count == 1 && unique[0].Summary == null) {
                diagnostics.Add($"PDF revision {revision.Key} contains an unreadable C2PA store; no active summary was selected.");
                canSelect = false;
            }
        }
        if (canSelect && eligible.Count > 0 && eligible[0].Summary is { } active) {
            // Known historical counts are independent of which revision can be active.
            string[] identities = stores.Where(store => store.Summary != null)
                .SelectMany(store => store.Summary!.ManifestIdentities).Distinct(StringComparer.Ordinal).ToArray();
            int count = identities.Length;
            eligible[0].Evidence.WithManifest(new OfficeC2paManifestSummary(active.Label, active.ClaimGenerator,
                active.Title, active.Format, active.Actions, active.Ingredients, active.SignedBy,
                active.CertificateIssuer, count, active.DeclaresGenerativeAi, identities));
        }
        return diagnostics.AsReadOnly();
    }
}
