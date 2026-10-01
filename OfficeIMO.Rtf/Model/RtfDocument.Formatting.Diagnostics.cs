namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    internal RtfConversionReport GetStyleConversionDiagnostics() {
        RtfConversionReport report = GetStyleReferenceDiagnostics();
        foreach (RtfConversionDiagnostic diagnostic in _sourceNormalizationDiagnostics) {
            if (diagnostic.Code is "RtfNormalizationStyleControlOmitted" or "RtfNormalizationStyleDestinationOmitted") report.Add(diagnostic);
        }
        return report;
    }

    private RtfConversionReport GetStyleReferenceDiagnostics() {
        var report = new RtfConversionReport();
        var styles = new Dictionary<(RtfStyleKind Kind, int Id), RtfStyle>();
        foreach (RtfStyle style in Styles) styles[(style.Kind, style.Id)] = style;
        var completed = new HashSet<(RtfStyleKind Kind, int Id)>();
        var missing = new HashSet<(RtfStyleKind Kind, int Id)>();
        foreach (RtfStyle style in Styles) Inspect(style.Id, style.Kind);
        foreach (RtfParagraph paragraph in Writing.RtfDocumentWriter.EnumerateParagraphs(this)) {
            Inspect(paragraph.StyleId ?? 0, RtfStyleKind.Paragraph);
            foreach (RtfRun run in paragraph.Runs) Inspect(run.StyleId, RtfStyleKind.Character);
        }
        return report;

        void Inspect(int? id, RtfStyleKind kind) {
            var path = new List<(RtfStyleKind Kind, int Id)>();
            var positions = new Dictionary<(RtfStyleKind Kind, int Id), int>();
            while (id.HasValue) {
                var key = (kind, id.Value);
                if (completed.Contains(key)) break;
                if (positions.TryGetValue(key, out int cycleStart)) {
                    report.Add(RtfConversionSeverity.Warning, "RtfStyleInheritanceCycle",
                        "A stylesheet inheritance cycle was stopped at the last reachable style. Effective formatting uses each reachable style once.",
                        RtfConversionAction.Flattened, sourcePath: "Styles/" + kind + "/" + id.Value.ToString(CultureInfo.InvariantCulture),
                        feature: "StyleInheritance", count: path.Count - cycleStart);
                    break;
                }
                if (!styles.TryGetValue(key, out RtfStyle? style)) {
                    // Paragraph style zero is the implicit normal style when no explicit definition exists.
                    if (!(kind == RtfStyleKind.Paragraph && id.Value == 0) && missing.Add(key)) {
                        report.Add(RtfConversionSeverity.Warning, "RtfStyleReferenceMissing",
                            "A referenced style has no matching definition. Formatting resolves from the remaining reachable styles and direct values.",
                            RtfConversionAction.Flattened, sourcePath: "Styles/" + kind + "/" + id.Value.ToString(CultureInfo.InvariantCulture), feature: "StyleInheritance");
                    }
                    break;
                }
                positions.Add(key, path.Count);
                path.Add(key);
                id = style.BasedOnStyleId;
            }
            foreach (var key in path) completed.Add(key);
        }
    }
}
