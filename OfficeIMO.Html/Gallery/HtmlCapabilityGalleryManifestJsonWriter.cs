using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>
/// Writes deterministic JSON manifests for OfficeIMO HTML capability gallery runs.
/// </summary>
public static class HtmlCapabilityGalleryManifestJsonWriter {
    private const string SchemaId = "officeimo.html.capability-gallery";
    private const string SchemaVersion = "1.0";

    /// <summary>
    /// Converts a gallery manifest to a deterministic JSON payload.
    /// </summary>
    /// <param name="manifest">Manifest to serialize.</param>
    /// <returns>Indented JSON containing scenario, profile contract, artifacts, expectations, score, resources, and diagnostics.</returns>
    public static string ToJson(HtmlCapabilityGalleryManifest manifest) {
        if (manifest == null) {
            throw new ArgumentNullException(nameof(manifest));
        }

        HtmlConversionProfileContract contract = HtmlConversionProfileContracts.Get(manifest.Profile);
        var builder = new StringBuilder();
        builder.AppendLine("{");
        AppendStringProperty(builder, 1, "schemaId", SchemaId, comma: true);
        AppendStringProperty(builder, 1, "schemaVersion", SchemaVersion, comma: true);
        AppendScenario(builder, manifest, comma: true);
        AppendProfile(builder, contract, comma: true);
        AppendOfficeProfiles(builder, manifest.OfficeProfiles, comma: true);
        AppendExpectations(builder, manifest.Expectations, comma: true);
        AppendArtifacts(builder, manifest.Result.Artifacts, comma: true);
        AppendRoundTripScore(builder, manifest.RoundTripScore, comma: true);
        AppendResources(builder, manifest.ResourceManifest, comma: true);
        AppendDiagnostics(builder, manifest);
        builder.AppendLine();
        builder.Append('}');
        return builder.ToString();
    }

    private static void AppendScenario(StringBuilder builder, HtmlCapabilityGalleryManifest manifest, bool comma) {
        HtmlCapabilityGalleryScenario scenario = manifest.Result.Scenario;
        AppendIndent(builder, 1).AppendLine("\"scenario\": {");
        AppendStringProperty(builder, 2, "id", scenario.Id, comma: true);
        AppendStringProperty(builder, 2, "title", scenario.Title, comma: true);
        AppendStringProperty(builder, 2, "category", scenario.Category, comma: true);
        AppendStringProperty(builder, 2, "description", scenario.Description);
        AppendIndent(builder, 1).Append('}');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendProfile(StringBuilder builder, HtmlConversionProfileContract contract, bool comma) {
        AppendIndent(builder, 1).AppendLine("\"profile\": {");
        AppendStringProperty(builder, 2, "id", contract.Profile.ToString(), comma: true);
        AppendStringProperty(builder, 2, "name", contract.Name, comma: true);
        AppendStringProperty(builder, 2, "intendedUse", contract.IntendedUse, comma: true);
        AppendStringProperty(builder, 2, "fidelityGoal", contract.FidelityGoal, comma: true);
        AppendStringArrayProperty(builder, 2, "supportedHtml", contract.SupportedHtml, comma: true);
        AppendStringArrayProperty(builder, 2, "supportedCss", contract.SupportedCss, comma: true);
        AppendStringArrayProperty(builder, 2, "resourceGuarantees", contract.ResourceGuarantees, comma: true);
        AppendStringArrayProperty(builder, 2, "diagnosticGuarantees", contract.DiagnosticGuarantees);
        AppendIndent(builder, 1).Append('}');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendOfficeProfiles(StringBuilder builder, IReadOnlyList<OfficeHtmlConversionProfile> officeProfiles, bool comma) {
        AppendIndent(builder, 1).AppendLine("\"officeProfiles\": [");
        if (officeProfiles != null) {
            for (int i = 0; i < officeProfiles.Count; i++) {
                OfficeHtmlConversionProfileContract contract = OfficeHtmlConversionProfileContracts.Get(officeProfiles[i]);
                AppendIndent(builder, 2).AppendLine("{");
                AppendStringProperty(builder, 3, "id", contract.Profile.ToString(), comma: true);
                AppendStringProperty(builder, 3, "sourceFormat", contract.SourceFormat, comma: true);
                AppendStringProperty(builder, 3, "name", contract.Name, comma: true);
                AppendStringProperty(builder, 3, "sharedProfile", contract.SharedProfile.ToString(), comma: true);
                AppendStringProperty(builder, 3, "intendedUse", contract.IntendedUse, comma: true);
                AppendStringProperty(builder, 3, "fidelityGoal", contract.FidelityGoal, comma: true);
                AppendStringProperty(builder, 3, "visualPrimitiveOwner", contract.VisualPrimitiveOwner, comma: true);
                AppendStringArrayProperty(builder, 3, "supportedHtml", contract.SupportedHtml, comma: true);
                AppendStringArrayProperty(builder, 3, "resourceGuarantees", contract.ResourceGuarantees, comma: true);
                AppendStringArrayProperty(builder, 3, "diagnosticGuarantees", contract.DiagnosticGuarantees);
                AppendIndent(builder, 2).Append('}');
                AppendCommaAndLine(builder, i < officeProfiles.Count - 1);
            }
        }

        AppendIndent(builder, 1).Append(']');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendExpectations(StringBuilder builder, IReadOnlyList<HtmlCapabilityGalleryExpectation> expectations, bool comma) {
        AppendIndent(builder, 1).AppendLine("\"expectations\": [");
        for (int i = 0; i < expectations.Count; i++) {
            HtmlCapabilityGalleryExpectation expectation = expectations[i];
            AppendIndent(builder, 2).AppendLine("{");
            AppendStringProperty(builder, 3, "feature", expectation.Feature, comma: true);
            AppendStringProperty(builder, 3, "outcome", expectation.Outcome.ToString(), comma: true);
            AppendStringProperty(builder, 3, "evidence", expectation.Evidence);
            AppendIndent(builder, 2).Append('}');
            AppendCommaAndLine(builder, i < expectations.Count - 1);
        }

        AppendIndent(builder, 1).Append(']');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendArtifacts(StringBuilder builder, IReadOnlyList<HtmlCapabilityGalleryArtifact> artifacts, bool comma) {
        AppendIndent(builder, 1).AppendLine("\"artifacts\": [");
        for (int i = 0; i < artifacts.Count; i++) {
            HtmlCapabilityGalleryArtifact artifact = artifacts[i];
            AppendIndent(builder, 2).AppendLine("{");
            AppendStringProperty(builder, 3, "id", artifact.Id, comma: true);
            AppendStringProperty(builder, 3, "kind", artifact.Kind, comma: true);
            AppendStringProperty(builder, 3, "path", artifact.Path, comma: true);
            AppendStringProperty(builder, 3, "mediaType", artifact.MediaType, comma: true);
            AppendNumberProperty(builder, 3, "length", artifact.Length, comma: true);
            AppendStringProperty(builder, 3, "sha256", artifact.Sha256);
            AppendIndent(builder, 2).Append('}');
            AppendCommaAndLine(builder, i < artifacts.Count - 1);
        }

        AppendIndent(builder, 1).Append(']');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendRoundTripScore(StringBuilder builder, HtmlRoundTripScore roundTripScore, bool comma) {
        AppendIndent(builder, 1).AppendLine("\"roundTripScore\": {");
        if (roundTripScore != null) {
            AppendNumberProperty(builder, 2, "score", roundTripScore.Score, comma: true);
            AppendNumberProperty(builder, 2, "sourceNodeCount", roundTripScore.SourceNodeCount, comma: true);
            AppendNumberProperty(builder, 2, "targetNodeCount", roundTripScore.TargetNodeCount, comma: true);
            AppendNumberProperty(builder, 2, "matchedFeatureCount", roundTripScore.MatchedFeatureCount, comma: true);
            AppendNumberProperty(builder, 2, "comparedFeatureCount", roundTripScore.ComparedFeatureCount, comma: true);
            AppendIndent(builder, 2).AppendLine("\"metrics\": {");
            IReadOnlyList<KeyValuePair<string, double>> metrics = roundTripScore.Metrics
                .OrderBy(metric => metric.Key, StringComparer.Ordinal)
                .ToList();
            for (int i = 0; i < metrics.Count; i++) {
                KeyValuePair<string, double> metric = metrics[i];
                AppendNumberProperty(builder, 3, metric.Key, metric.Value, comma: i < metrics.Count - 1);
            }

            AppendIndent(builder, 2).AppendLine();
            AppendIndent(builder, 2).Append('}');
            builder.AppendLine();
        }

        AppendIndent(builder, 1).Append('}');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendResources(StringBuilder builder, HtmlResourceManifest resourceManifest, bool comma) {
        AppendIndent(builder, 1).AppendLine("\"resources\": {");
        if (resourceManifest != null) {
            AppendNumberProperty(builder, 2, "allowedCount", resourceManifest.AllowedCount, comma: true);
            AppendNumberProperty(builder, 2, "blockedCount", resourceManifest.BlockedCount, comma: true);
            AppendIndent(builder, 2).AppendLine("\"items\": [");
            for (int i = 0; i < resourceManifest.Resources.Count; i++) {
                HtmlResourceReference resource = resourceManifest.Resources[i];
                AppendIndent(builder, 3).AppendLine("{");
                AppendStringProperty(builder, 4, "kind", resource.Kind.ToString(), comma: true);
                AppendStringProperty(builder, 4, "elementName", resource.ElementName, comma: true);
                AppendStringProperty(builder, 4, "attributeName", resource.AttributeName, comma: true);
                AppendStringProperty(builder, 4, "source", resource.Source, comma: true);
                AppendStringProperty(builder, 4, "resolvedSource", resource.ResolvedSource, comma: true);
                AppendBooleanProperty(builder, 4, "isAllowed", resource.IsAllowed, comma: true);
                AppendStringProperty(builder, 4, "diagnosticCode", resource.DiagnosticCode);
                AppendIndent(builder, 3).Append('}');
                AppendCommaAndLine(builder, i < resourceManifest.Resources.Count - 1);
            }

            AppendIndent(builder, 2).AppendLine();
            AppendIndent(builder, 2).Append(']');
            builder.AppendLine();
        }

        AppendIndent(builder, 1).Append('}');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendDiagnostics(StringBuilder builder, HtmlCapabilityGalleryManifest manifest) {
        IReadOnlyList<KeyValuePair<string, HtmlDiagnostic>> diagnostics = GetDiagnostics(manifest);
        AppendIndent(builder, 1).AppendLine("\"diagnostics\": [");
        for (int i = 0; i < diagnostics.Count; i++) {
            KeyValuePair<string, HtmlDiagnostic> item = diagnostics[i];
            HtmlDiagnostic diagnostic = item.Value;
            AppendIndent(builder, 2).AppendLine("{");
            AppendStringProperty(builder, 3, "origin", item.Key, comma: true);
            AppendStringProperty(builder, 3, "component", diagnostic.Component, comma: true);
            AppendStringProperty(builder, 3, "code", diagnostic.Code, comma: true);
            AppendStringProperty(builder, 3, "severity", diagnostic.Severity.ToString(), comma: true);
            AppendStringProperty(builder, 3, "message", diagnostic.Message, comma: true);
            AppendNullableStringProperty(builder, 3, "source", diagnostic.Source, comma: true);
            AppendNullableStringProperty(builder, 3, "detail", diagnostic.Detail);
            AppendIndent(builder, 2).Append('}');
            AppendCommaAndLine(builder, i < diagnostics.Count - 1);
        }

        AppendIndent(builder, 1).Append(']');
    }

    private static IReadOnlyList<KeyValuePair<string, HtmlDiagnostic>> GetDiagnostics(HtmlCapabilityGalleryManifest manifest) {
        var diagnostics = new List<KeyValuePair<string, HtmlDiagnostic>>();
        foreach (HtmlDiagnostic diagnostic in manifest.Result.Diagnostics.Diagnostics) {
            diagnostics.Add(new KeyValuePair<string, HtmlDiagnostic>("result", diagnostic));
        }

        if (manifest.ResourceManifest != null) {
            foreach (HtmlDiagnostic diagnostic in manifest.ResourceManifest.Diagnostics.Diagnostics) {
                diagnostics.Add(new KeyValuePair<string, HtmlDiagnostic>("resource", diagnostic));
            }
        }

        return diagnostics;
    }

    private static void AppendStringArrayProperty(StringBuilder builder, int indent, string name, IReadOnlyList<string> values, bool comma = false) {
        AppendIndent(builder, indent).Append('"').Append(Escape(name)).Append("\": [");
        for (int i = 0; i < values.Count; i++) {
            if (i > 0) {
                builder.Append(", ");
            }

            builder.Append('"').Append(Escape(values[i])).Append('"');
        }

        builder.Append(']');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendStringProperty(StringBuilder builder, int indent, string name, string value, bool comma = false) {
        AppendIndent(builder, indent)
            .Append('"')
            .Append(Escape(name))
            .Append("\": \"")
            .Append(Escape(value))
            .Append('"');
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendNullableStringProperty(StringBuilder builder, int indent, string name, string? value, bool comma = false) {
        AppendIndent(builder, indent)
            .Append('"')
            .Append(Escape(name))
            .Append("\": ");
        if (value == null) {
            builder.Append("null");
        } else {
            builder.Append('"').Append(Escape(value)).Append('"');
        }

        AppendCommaAndLine(builder, comma);
    }

    private static void AppendNumberProperty(StringBuilder builder, int indent, string name, long value, bool comma = false) {
        AppendIndent(builder, indent)
            .Append('"')
            .Append(Escape(name))
            .Append("\": ")
            .Append(value.ToString(CultureInfo.InvariantCulture));
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendNumberProperty(StringBuilder builder, int indent, string name, double value, bool comma = false) {
        AppendIndent(builder, indent)
            .Append('"')
            .Append(Escape(name))
            .Append("\": ")
            .Append(value.ToString("0.###", CultureInfo.InvariantCulture));
        AppendCommaAndLine(builder, comma);
    }

    private static void AppendBooleanProperty(StringBuilder builder, int indent, string name, bool value, bool comma = false) {
        AppendIndent(builder, indent)
            .Append('"')
            .Append(Escape(name))
            .Append("\": ")
            .Append(value ? "true" : "false");
        AppendCommaAndLine(builder, comma);
    }

    private static StringBuilder AppendIndent(StringBuilder builder, int indent) {
        return builder.Append(' ', indent * 2);
    }

    private static void AppendCommaAndLine(StringBuilder builder, bool comma) {
        if (comma) {
            builder.Append(',');
        }

        builder.AppendLine();
    }

    private static string Escape(string value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length + 8);
        foreach (char ch in value) {
            switch (ch) {
                case '"':
                    builder.Append("\\\"");
                    break;
                case '\\':
                    builder.Append("\\\\");
                    break;
                case '\b':
                    builder.Append("\\b");
                    break;
                case '\f':
                    builder.Append("\\f");
                    break;
                case '\n':
                    builder.Append("\\n");
                    break;
                case '\r':
                    builder.Append("\\r");
                    break;
                case '\t':
                    builder.Append("\\t");
                    break;
                default:
                    if (ch < ' ') {
                        builder.Append("\\u").Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    } else {
                        builder.Append(ch);
                    }

                    break;
            }
        }

        return builder.ToString();
    }
}
