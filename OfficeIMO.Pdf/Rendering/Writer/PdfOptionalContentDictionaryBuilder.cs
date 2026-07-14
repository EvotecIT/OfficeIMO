namespace OfficeIMO.Pdf;

internal static class PdfOptionalContentDictionaryBuilder {
    internal static string BuildGroup(PdfLayerDefinition definition) {
        Guard.NotNull(definition, nameof(definition));
        PdfLayerOptions options = definition.Options;
        var sb = new StringBuilder("<< /Type /OCG /Name ");
        sb.Append(PdfSyntaxEscaper.TextString(definition.Name))
            .Append(" /Intent [/View /Design] /Usage <<")
            .Append(" /View << /ViewState /").Append(options.VisibleInViewer ? "ON" : "OFF").Append(" >>")
            .Append(" /Print << /PrintState /").Append(options.VisibleWhenPrinting ? "ON" : "OFF").Append(" >>")
            .Append(" /Export << /ExportState /").Append(options.VisibleWhenExporting ? "ON" : "OFF").Append(" >>")
            .Append(" >> >>\n");
        return sb.ToString();
    }

    internal static string BuildProperties(
        IReadOnlyList<PdfLayerDefinition> definitions,
        IReadOnlyDictionary<PdfLayerDefinition, int> objectIds) {
        Guard.NotNull(definitions, nameof(definitions));
        Guard.NotNull(objectIds, nameof(objectIds));
        var sb = new StringBuilder("<< /OCGs [");
        AppendReferences(sb, definitions, objectIds, _ => true);
        sb.Append("] /D << /Name ").Append(PdfSyntaxEscaper.TextString("Layers"))
            .Append(" /Creator ").Append(PdfSyntaxEscaper.TextString("OfficeIMO.Pdf"))
            .Append(" /BaseState /ON /ON [");
        AppendReferences(sb, definitions, objectIds, definition => definition.Options.InitiallyVisible);
        sb.Append("] /OFF [");
        AppendReferences(sb, definitions, objectIds, definition => !definition.Options.InitiallyVisible);
        sb.Append("] /Locked [");
        AppendReferences(sb, definitions, objectIds, definition => definition.Options.Locked);
        sb.Append("] /Order [");
        AppendReferences(sb, definitions, objectIds, _ => true);
        sb.Append("] >> >>\n");
        return sb.ToString();
    }

    private static void AppendReferences(
        StringBuilder sb,
        IReadOnlyList<PdfLayerDefinition> definitions,
        IReadOnlyDictionary<PdfLayerDefinition, int> objectIds,
        Func<PdfLayerDefinition, bool> predicate) {
        bool first = true;
        foreach (PdfLayerDefinition definition in definitions) {
            if (!predicate(definition)) continue;
            if (!first) sb.Append(' ');
            sb.Append(PdfSyntaxEscaper.IndirectReference(objectIds[definition]));
            first = false;
        }
    }
}
