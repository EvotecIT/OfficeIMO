namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static PdfArray ReplaceWatermarkContent(Dictionary<int, PdfIndirectObject> objects, PdfArray contents,
        int replacementNumber, string identifier, bool behindContent) {
        var result = new PdfArray();
        bool replaced = false;
        bool preservePosition = contents.Items.Any(item => IsWatermark(objects, item, identifier, out var stream)
            && item is PdfReference reference && reference.ObjectNumber != replacementNumber
            && stream!.Dictionary.Get<PdfBoolean>("OfficeIMOWatermarkBehind")?.Value == behindContent);
        foreach (PdfObject item in contents.Items) {
            if (item is PdfReference reference && reference.ObjectNumber == replacementNumber) {
                if (!preservePosition) result.Items.Add(item);
                continue;
            }
            if (IsWatermark(objects, item, identifier, out _)) {
                if (preservePosition && !replaced) result.Items.Add(new PdfReference(replacementNumber, 0));
                replaced = true;
            } else result.Items.Add(item);
        }
        return result;
    }

    private static void RemoveWatermarksOutsideTargets(Dictionary<int, PdfIndirectObject> objects, int[] pages,
        Dictionary<int, Dictionary<string, PdfObject>> overrides, IReadOnlyList<PageStampRequest> requests) {
        var targets = new Dictionary<string, HashSet<int>>(StringComparer.Ordinal);
        foreach (var request in requests) {
            if (request.Options.ContentIdentifier is not { } identifier) continue;
            if (!targets.TryGetValue(identifier, out var selected)) targets[identifier] = selected = new HashSet<int>();
            foreach (int page in request.Options.TargetPages?.Resolve(pages.Length) ?? Enumerable.Range(1, pages.Length).ToArray())
                selected.Add(page);
        }
        if (targets.Count == 0) return;
        for (int index = 0; index < pages.Length; index++) {
            int pageObjectNumber = pages[index];
            var page = (PdfDictionary)objects[pageObjectNumber].Value;
            overrides.TryGetValue(pageObjectNumber, out var pageOverride);
            PdfObject? contents = pageOverride != null && pageOverride.TryGetValue("Contents", out var modified)
                ? modified : page.Items.TryGetValue("Contents", out var original) ? original : null;
            var entries = new PdfArray();
            AppendContentEntries(objects, entries, contents);
            int removed = entries.Items.RemoveAll(item => targets.Any(target => !target.Value.Contains(index + 1)
                && IsWatermark(objects, item, target.Key, out _)));
            if (removed == 0) continue;
            if (pageOverride is null) overrides[pageObjectNumber] = pageOverride = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
            pageOverride["Contents"] = entries;
        }
    }

    private static bool IsWatermark(Dictionary<int, PdfIndirectObject> objects, PdfObject item, string identifier,
        out PdfStream? stream) {
        stream = PdfObjectLookup.Resolve(objects, item) as PdfStream;
        return stream?.Dictionary.Get<PdfStringObj>("OfficeIMOWatermarkId")?.Value == identifier;
    }
}
