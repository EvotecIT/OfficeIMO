using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    // Only resources explicitly produced for editable watermarks are candidates. Usage is
    // deliberately collected across every page/form/pattern, so shared or inherited use wins.
    private static void PruneUnusedWatermarkResources(Dictionary<int, PdfIndirectObject> objects, int[] pages,
        Dictionary<int, Dictionary<string, PdfObject>> overrides, PdfLoadOptions? readOptions) {
        if (!objects.Values.Any(item => ResourceDictionary(item.Value)?.Get<PdfStringObj>("OfficeIMOWatermarkResource") is not null)) return;
        var sequences = new List<PdfStream[]>();
        foreach (int pageNumber in pages) {
            var page = CloneDictionary((PdfDictionary)objects[pageNumber].Value);
            if (overrides.TryGetValue(pageNumber, out var changes))
                foreach (var change in changes) page.Items[change.Key] = change.Value;
            sequences.Add(GetPageContentStreams(objects, page).ToArray());
        }
        foreach (var item in objects.Values) {
            if (item.Value is PdfStream stream &&
                (stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form" || stream.Dictionary.Items.ContainsKey("PatternType")))
                sequences.Add(new[] { stream });
        }
        var used = new HashSet<string>(StringComparer.Ordinal);
        var limits = PdfLoadOptions.Resolve(readOptions).Limits;
        try {
            foreach (var sequence in sequences) {
                // Page streams share operand state, including a name in one stream followed
                // by its Do/gs operator in the next. Match the reader's ordered sequence.
                var content = new System.Text.StringBuilder();
                foreach (var stream in sequence) {
                    if (stream.DecodingFailed) return;
                    int remaining = limits.MaxPageContentBytes - content.Length - 1;
                    if (remaining <= 0) return;
                    byte[] decoded = StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects,
                        Math.Min(limits.MaxDecodedStreamBytes, remaining));
                    if (content.Length > 0) content.Append('\n');
                    content.Append(PdfEncoding.Latin1GetString(decoded));
                }
                PdfContentStreamInterpreter.Interpret(content.ToString(), limits.MaxContentOperations, operation => {
                    if (operation.Name is "Do" or "gs" && operation.Operands.Count > 0
                        && operation.Operands[operation.Operands.Count - 1] is string name)
                        used.Add(operation.Name + ":" + name);
                }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
            }
        } catch (System.IO.IOException) {
            // Pruning is optional. An unreadable or over-limit stream, including an unused
            // form, means usage is unknown: retain every resource rather than reject the edit.
            return;
        }
        foreach (int pageNumber in pages) {
            var page = (PdfDictionary)objects[pageNumber].Value;
            overrides.TryGetValue(pageNumber, out var changes);
            PdfObject? value = changes != null && changes.TryGetValue("Resources", out var modified)
                ? modified : GetInheritedPageValue(objects, page, "Resources");
            var resources = CloneDictionary(ResolveDictionary(objects, value));
            bool changed = false;
            foreach (var category in new[] { (Key: "XObject", Operator: "Do"), (Key: "ExtGState", Operator: "gs") }) {
                if (!resources.Items.TryGetValue(category.Key, out var categoryValue)) continue;
                var entries = CloneDictionary(ResolveDictionary(objects, categoryValue));
                foreach (var entry in entries.Items.ToArray()) {
                    var target = ResourceDictionary(PdfObjectLookup.Resolve(objects, entry.Value));
                    if (target?.Get<PdfStringObj>("OfficeIMOWatermarkResource") is null
                        || used.Contains(category.Operator + ":" + entry.Key)) continue;
                    entries.Items.Remove(entry.Key);
                    changed = true;
                }
                resources.Items[category.Key] = entries;
            }
            if (!changed) continue;
            if (changes is null) overrides[pageNumber] = changes = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
            changes["Resources"] = resources;
        }
    }

    private static PdfDictionary? ResourceDictionary(PdfObject? value) => value switch {
        PdfStream stream => stream.Dictionary,
        PdfDictionary dictionary => dictionary,
        _ => null
    };
}
