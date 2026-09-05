namespace OfficeIMO.Html;

public static partial class HtmlCapabilityGalleryManifestJsonWriter {
    private static void AppendArtifactEvidence(StringBuilder builder, HtmlCapabilityGalleryArtifactEvidence? evidence) {
        if (evidence == null) {
            AppendNullProperty(builder, 3, "evidence");
            return;
        }
        AppendIndent(builder, 3).AppendLine("\"evidence\": {");
        AppendNumberProperty(builder, 4, "pageCount", evidence.PageCount, comma: true);
        if (evidence.PageNumber.HasValue) AppendNumberProperty(builder, 4, "pageNumber", evidence.PageNumber.Value, comma: true);
        else AppendNullProperty(builder, 4, "pageNumber", comma: true);
        if (evidence.Width.HasValue) AppendNumberProperty(builder, 4, "width", evidence.Width.Value, comma: true);
        else AppendNullProperty(builder, 4, "width", comma: true);
        if (evidence.Height.HasValue) AppendNumberProperty(builder, 4, "height", evidence.Height.Value, comma: true);
        else AppendNullProperty(builder, 4, "height", comma: true);
        AppendStringProperty(builder, 4, "dimensionUnit", evidence.DimensionUnit, comma: true);
        AppendBooleanProperty(builder, 4, "hasLoss", evidence.HasLoss, comma: true);
        AppendIndent(builder, 4).AppendLine("\"checks\": [");
        for (int index = 0; index < evidence.Checks.Count; index++) {
            HtmlCapabilityGalleryCheck check = evidence.Checks[index];
            AppendIndent(builder, 5).AppendLine("{");
            AppendStringProperty(builder, 6, "name", check.Name, comma: true);
            AppendBooleanProperty(builder, 6, "passed", check.Passed, comma: true);
            AppendStringProperty(builder, 6, "detail", check.Detail);
            AppendIndent(builder, 5).Append('}');
            AppendCommaAndLine(builder, index < evidence.Checks.Count - 1);
        }
        AppendIndent(builder, 4).AppendLine("],");
        AppendIndent(builder, 4).AppendLine("\"diagnostics\": [");
        for (int index = 0; index < evidence.Diagnostics.Count; index++) {
            HtmlDiagnostic diagnostic = evidence.Diagnostics[index];
            AppendIndent(builder, 5).AppendLine("{");
            AppendStringProperty(builder, 6, "component", diagnostic.Component, comma: true);
            AppendStringProperty(builder, 6, "code", diagnostic.Code, comma: true);
            AppendStringProperty(builder, 6, "severity", diagnostic.Severity.ToString(), comma: true);
            AppendStringProperty(builder, 6, "lossKind", diagnostic.LossKind.ToString(), comma: true);
            AppendNullableStringProperty(builder, 6, "source", diagnostic.Source, comma: true);
            AppendStringProperty(builder, 6, "message", diagnostic.Message);
            AppendIndent(builder, 5).Append('}');
            AppendCommaAndLine(builder, index < evidence.Diagnostics.Count - 1);
        }
        AppendIndent(builder, 4).AppendLine("]");
        AppendIndent(builder, 3).AppendLine("}");
    }
}
