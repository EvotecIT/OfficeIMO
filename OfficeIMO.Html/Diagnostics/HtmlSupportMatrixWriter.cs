using System.Text;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Html;

/// <summary>Generates a support matrix from the shared profile contracts and diagnostic catalog.</summary>
public static class HtmlSupportMatrixWriter {
    /// <summary>Generates deterministic Markdown describing the current profile contracts and diagnostic boundaries.</summary>
    public static string ToMarkdown() {
        var builder = new StringBuilder();
        builder.AppendLine("# OfficeIMO HTML support matrix");
        builder.AppendLine();
        builder.AppendLine("This file is generated from `HtmlConversionProfileContracts`, `HtmlTargetCapabilityContracts`, and `HtmlDiagnosticCatalog`. Profile and target entries are tested contracts; diagnostic entries describe bounded fallbacks, policy decisions, and safety limits.");
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
        builder.AppendLine("| Target | Package | Artifact | HTML import | Result contract | Reverse HTML | Reverse result | Profiles | I/O and async boundary |");
        builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- | --- |");
        foreach (HtmlTargetCapabilityContract contract in HtmlTargetCapabilityContracts.All.OrderBy(item => item.Target)) {
            builder.Append("| ").Append(contract.Target)
                .Append(" | `").Append(EscapeCode(contract.PackageName)).Append("` | ")
                .Append(EscapeCell(contract.ArtifactName)).Append(" | `")
                .Append(EscapeCode(contract.ImportEntryPoint)).Append("` | `")
                .Append(EscapeCode(contract.ImportResultContract)).Append("` | ")
                .Append(FormatCode(contract.ExportEntryPoint)).Append(" | ")
                .Append(FormatCode(contract.ExportResultContract)).Append(" | ")
                .Append(EscapeCell(string.Join(", ", contract.Profiles))).Append(" | ")
                .Append(EscapeCell(contract.IoAndAsyncBoundary)).AppendLine(" |");
        }

        builder.AppendLine();
        builder.AppendLine("## Target semantic capability contracts");
        builder.AppendLine();
        builder.AppendLine("| Target | Supported | Approximated | Unsupported |");
        builder.AppendLine("| --- | --- | --- | --- |");
        foreach (HtmlTargetCapabilityContract contract in HtmlTargetCapabilityContracts.All.OrderBy(item => item.Target)) {
            builder.Append("| ").Append(contract.Target).Append(" | ")
                .Append(EscapeCell(FormatFeatures(contract.SupportedFeatures))).Append(" | ")
                .Append(EscapeCell(FormatFeatures(contract.ApproximatedFeatures))).Append(" | ")
                .Append(EscapeCell(FormatFeatures(contract.UnsupportedFeatures))).AppendLine(" |");
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

    private static string EscapeCell(string value) => (value ?? string.Empty)
        .Replace("\\", "\\\\")
        .Replace("|", "\\|")
        .Replace("\r", " ")
        .Replace("\n", " ");

    private static string EscapeCode(string value) => (value ?? string.Empty).Replace("`", "\\`");

    private static string FormatCode(string? value) => value == null ? "—" : "`" + EscapeCode(value) + "`";

    private static string FormatFeatures(IReadOnlyList<HtmlSemanticFeature> features) =>
        features.Count == 0 ? "None" : string.Join(", ", features);
}
