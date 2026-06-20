namespace OfficeIMO.Html;

/// <summary>
/// Writes stable markdown manifests for OfficeIMO HTML capability gallery runs.
/// </summary>
public static class HtmlCapabilityGalleryManifestWriter {
    /// <summary>
    /// Converts a gallery manifest to markdown.
    /// </summary>
    public static string ToMarkdown(HtmlCapabilityGalleryManifest manifest) {
        if (manifest == null) {
            throw new ArgumentNullException(nameof(manifest));
        }

        var builder = new StringBuilder();
        HtmlConversionProfileContract contract = HtmlConversionProfileContracts.Get(manifest.Profile);
        builder.AppendLine("# HTML Capability Gallery Scenario");
        builder.AppendLine();
        builder.AppendLine("Id: " + manifest.Result.Scenario.Id);
        builder.AppendLine("Title: " + manifest.Result.Scenario.Title);
        builder.AppendLine("Category: " + manifest.Result.Scenario.Category);
        builder.AppendLine("Profile: " + contract.Name);
        builder.AppendLine("Description: " + manifest.Result.Scenario.Description);
        builder.AppendLine("Fidelity: " + contract.FidelityGoal);
        builder.AppendLine();
        builder.AppendLine("## Artifacts");
        foreach (HtmlCapabilityGalleryArtifact artifact in manifest.Result.Artifacts) {
            builder.AppendLine("- " + artifact.Kind + ": " + Path.GetFileName(artifact.Path) + " (" + artifact.MediaType + ", " + artifact.Length + " bytes, sha256=" + artifact.Sha256 + ")");
        }

        if (manifest.RoundTripScore != null) {
            builder.AppendLine();
            builder.AppendLine("## Round Trip Score");
            builder.AppendLine("- Score: " + manifest.RoundTripScore.Score.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture));
            foreach (var metric in manifest.RoundTripScore.Metrics) {
                builder.AppendLine("- " + metric.Key + ": " + metric.Value.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        if (manifest.ResourceManifest != null) {
            builder.AppendLine();
            builder.AppendLine("## Resources");
            builder.AppendLine("- Allowed: " + manifest.ResourceManifest.AllowedCount);
            builder.AppendLine("- Blocked: " + manifest.ResourceManifest.BlockedCount);
            foreach (HtmlResourceReference resource in manifest.ResourceManifest.Resources) {
                builder.AppendLine("- " + resource.Kind + ": " + resource.Source + " => " + (resource.IsAllowed ? resource.ResolvedSource : resource.DiagnosticCode));
            }
        }

        builder.AppendLine();
        builder.AppendLine("## Diagnostics");
        foreach (HtmlDiagnostic diagnostic in manifest.Result.Diagnostics.Diagnostics) {
            HtmlDiagnosticDefinition definition = HtmlDiagnosticCatalog.GetOrCreateGeneric(diagnostic.Code);
            builder.AppendLine("- " + diagnostic.Component + ":" + diagnostic.Code + ":" + diagnostic.Severity + ": " + diagnostic.Message + " [" + definition.Category + "]");
        }

        if (manifest.ResourceManifest != null) {
            foreach (HtmlDiagnostic diagnostic in manifest.ResourceManifest.Diagnostics.Diagnostics) {
                HtmlDiagnosticDefinition definition = HtmlDiagnosticCatalog.GetOrCreateGeneric(diagnostic.Code);
                builder.AppendLine("- " + diagnostic.Component + ":" + diagnostic.Code + ":" + diagnostic.Severity + ": " + diagnostic.Message + " [" + definition.Category + "]");
            }
        }

        return builder.ToString();
    }
}
