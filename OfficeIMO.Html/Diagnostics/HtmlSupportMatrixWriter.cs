using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <summary>Generates a support matrix from shared profile, target, renderer, and diagnostic contracts.</summary>
public static class HtmlSupportMatrixWriter {
    /// <summary>Generates deterministic Markdown describing the current executable compatibility contracts.</summary>
    public static string ToMarkdown() {
        var builder = new StringBuilder();
        builder.AppendLine("# OfficeIMO HTML support matrix");
        builder.AppendLine();
        builder.AppendLine("This file is generated from `HtmlConversionProfileContracts`, `HtmlTargetCapabilityContracts`, `HtmlEditableLayoutCapabilityContracts`, `HtmlRenderCapabilityCatalog`, and `HtmlDiagnosticCatalog`. Entries describe tested behavior and bounded fallbacks; a parsed CSS property is not treated as rendered support unless the renderer contract says so.");
        builder.AppendLine();
        builder.AppendLine("## Conversion profiles");

        foreach (HtmlConversionProfileContract contract in HtmlConversionProfileContracts.All.OrderBy(item => item.Profile)) {
            builder.AppendLine();
            builder.Append("### ").AppendLine(contract.Name);
            builder.AppendLine();
            builder.Append("- Intended use: ").AppendLine(contract.IntendedUse);
            builder.Append("- Fidelity goal: ").AppendLine(contract.FidelityGoal);
            AppendList(builder, "Supported HTML", contract.SupportedHtml);
            AppendList(builder, "Supported CSS", contract.SupportedCss);
            AppendList(builder, "Resource guarantees", contract.ResourceGuarantees);
            AppendList(builder, "Diagnostic guarantees", contract.DiagnosticGuarantees);
        }

        builder.AppendLine();
        builder.AppendLine("## Target adapter API contracts");
        builder.AppendLine();
        builder.AppendLine("| Target | Direction | Package | Artifact | Entry point | Result contract | Profiles | I/O and async boundary | Diagnostics contract |");
        builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- | --- |");
        foreach (HtmlTargetCapabilityContract contract in HtmlTargetCapabilityContracts.All.OrderBy(item => item.Target)) {
            AppendRoute(builder, contract, "HTML to target", contract.HtmlToTarget);
            if (contract.TargetToHtml != null) {
                AppendRoute(builder, contract, "Target to HTML", contract.TargetToHtml);
            }
        }

        builder.AppendLine();
        builder.AppendLine("## Native editable-layout projection");
        builder.AppendLine();
        builder.AppendLine("The shared projector accepts only bounded single-surface regions. Regions fragmented across pages or columns remain in source flow with `HtmlEditableLayoutRegionFragmented` instead of acquiring ambiguous native geometry. Destination collision avoidance is reported with `HtmlEditableLayoutPlacementSimplified`.");
        builder.AppendLine();
        builder.AppendLine("| Target | Native regions | Native geometry | Paint and picture effects | Diagnostic boundary |");
        builder.AppendLine("| --- | --- | --- | --- | --- |");
        foreach (HtmlEditableLayoutCapabilityContract contract in HtmlEditableLayoutCapabilityContracts.All) {
            builder.Append("| ").Append(contract.Target).Append(" | ")
                .Append(EscapeCell(contract.NativeRegions)).Append(" | ")
                .Append(EscapeCell(contract.NativeGeometry)).Append(" | ")
                .Append(EscapeCell(contract.NativePaintAndEffects)).Append(" | ")
                .Append(EscapeCell(contract.DiagnosticBoundary)).AppendLine(" |");
        }

        builder.AppendLine();
        builder.AppendLine("## Target semantic capability contracts");
        builder.AppendLine();
        builder.AppendLine("| Target | Direction | Supported | Approximated | Unsupported |");
        builder.AppendLine("| --- | --- | --- | --- | --- |");
        foreach (HtmlTargetCapabilityContract contract in HtmlTargetCapabilityContracts.All.OrderBy(item => item.Target)) {
            AppendRouteFeatures(builder, contract.Target, "HTML to target", contract.HtmlToTarget);
            if (contract.TargetToHtml != null) {
                AppendRouteFeatures(builder, contract.Target, "Target to HTML", contract.TargetToHtml);
            }
        }

        builder.AppendLine();
        builder.AppendLine("## Direct renderer compatibility contracts");
        builder.AppendLine();
        builder.AppendLine("| Area | ID | Kind | Support | Features | Behavior | Diagnostics |");
        builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");
        foreach (HtmlRenderCapability capability in HtmlRenderCapabilityCatalog.All) {
            builder.Append("| ").Append(EscapeCell(capability.Area)).Append(" | `")
                .Append(EscapeCode(capability.Id)).Append("` | ")
                .Append(capability.Kind).Append(" | ")
                .Append(capability.SupportLevel).Append(" | ")
                .Append(EscapeCell(string.Join(", ", capability.Features))).Append(" | ")
                .Append(EscapeCell(capability.Behavior)).Append(" | ")
                .Append(EscapeCell(FormatCodes(capability.DiagnosticCodes))).AppendLine(" |");
        }

        builder.AppendLine();
        builder.AppendLine("## Diagnostic boundaries");
        builder.AppendLine();
        builder.AppendLine("| Category | Code | Severity | Meaning | Remediation |");
        builder.AppendLine("| --- | --- | --- | --- | --- |");
        foreach (HtmlDiagnosticDefinition definition in HtmlDiagnosticCatalog.Ordered) {
            builder.Append("| ")
                .Append(EscapeCell(definition.Category)).Append(" | `")
                .Append(EscapeCode(definition.Code)).Append("` | ")
                .Append(definition.DefaultSeverity).Append(" | ")
                .Append(EscapeCell(definition.Explanation)).Append(" | ")
                .Append(EscapeCell(definition.Remediation)).AppendLine(" |");
        }

        return builder.ToString().Replace("\r\n", "\n");
    }

    /// <summary>Writes the generated Markdown support matrix to a file, replacing any existing file.</summary>
    public static void WriteMarkdown(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A support-matrix path is required.", nameof(path));
        OfficeFileCommit.WriteAllBytes(path, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(ToMarkdown()));
    }

    private static void AppendList(StringBuilder builder, string label, IReadOnlyList<string> values) {
        builder.Append("- ").Append(label).Append(": ")
            .AppendLine(values.Count == 0 ? "None" : string.Join(", ", values));
    }

    private static void AppendRoute(StringBuilder builder, HtmlTargetCapabilityContract contract, string direction,
        HtmlConversionRouteCapabilityContract route) {
        builder.Append("| ").Append(contract.Target).Append(" | ").Append(direction)
            .Append(" | `").Append(EscapeCode(contract.PackageName)).Append("` | ")
            .Append(EscapeCell(contract.ArtifactName)).Append(" | `")
            .Append(EscapeCode(route.EntryPoint)).Append("` | `")
            .Append(EscapeCode(route.ResultContract)).Append("` | ")
            .Append(EscapeCell(string.Join(", ", route.Profiles))).Append(" | ")
            .Append(EscapeCell(route.IoAndAsyncBoundary)).Append(" | ")
            .Append(EscapeCell(route.DiagnosticsContract)).AppendLine(" |");
    }

    private static void AppendRouteFeatures(StringBuilder builder, HtmlConversionTarget target, string direction,
        HtmlConversionRouteCapabilityContract route) {
        builder.Append("| ").Append(target).Append(" | ").Append(direction).Append(" | ")
            .Append(EscapeCell(FormatFeatures(route.SupportedFeatures))).Append(" | ")
            .Append(EscapeCell(FormatFeatures(route.ApproximatedFeatures))).Append(" | ")
            .Append(EscapeCell(FormatFeatures(route.UnsupportedFeatures))).AppendLine(" |");
    }

    private static string EscapeCell(string value) => (value ?? string.Empty)
        .Replace("\\", "\\\\")
        .Replace("|", "\\|")
        .Replace("\r", " ")
        .Replace("\n", " ");

    private static string EscapeCode(string value) => (value ?? string.Empty).Replace("`", "\\`");

    private static string FormatCode(string? value) => value == null ? "—" : "`" + EscapeCode(value) + "`";

    private static string FormatFeatures(IReadOnlyList<HtmlSemanticFeature> features) =>
        features.Count == 0 ? "None" : string.Join(", ", features);

    private static string FormatCodes(IReadOnlyList<string> codes) =>
        codes.Count == 0 ? "None" : string.Join(", ", codes.Select(code => "`" + code + "`"));
}
