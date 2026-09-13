namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    // Resolve the root invocation, not its reusable Form/image resource. The same resource
    // may also be painted by ordinary page content that must retain normal editing behavior.
    internal Func<PdfContentOrderKey?, string?> CreateWatermarkIdentityResolver() {
        var identifiers = new Dictionary<int, string>();
        foreach (var entry in GetContentStreamObjects()) {
            if (entry.ObjectNumber is not int number) continue;
            var dictionary = entry.Stream.Dictionary;
            string? id = dictionary.Get<PdfStringObj>("OfficeIMOWatermarkId")?.Value;
            if (!Guid.TryParseExact(id, "N", out _)
                || !dictionary.Items.TryGetValue("OfficeIMOWatermarkSettings", out var settingsObject)
                || PdfObjectLookup.Resolve(_objects, settingsObject) is not PdfDictionary settings
                || settings.Get<PdfNumber>("Version")?.Value != 1) continue;
            identifiers[number] = id!;
        }
        if (identifiers.Count == 0) return static _ => null;
        var sequence = GetContentStreamSequence();
        return key => key?.RootOperatorOffset is int offset
            && sequence.GetObjectNumber(offset) is int number
            && identifiers.TryGetValue(number, out string? id) ? id : null;
    }
}
