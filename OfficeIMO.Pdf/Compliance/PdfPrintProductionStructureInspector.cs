using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static class PdfPrintProductionStructureInspector {
    internal static PdfPrintProductionStructureEvidence Inspect(PdfReadDocument document) {
        Guard.NotNull(document, nameof(document));
        int validBoxes = 0;
        int invalidBoxes = 0;
        foreach (PdfReadPage page in document.Pages) {
            if (HasValidProductionBoxes(page.GetGeometry())) validBoxes++;
            else invalidBoxes++;
        }

        var fontDictionaries = new HashSet<PdfDictionary>();
        var visitedDictionaries = new HashSet<PdfDictionary>();
        var visitedArrays = new HashSet<PdfArray>();
        foreach (PdfIndirectObject indirect in document.Objects.Values) {
            CollectFontResources(indirect.Value, document.Objects, visitedDictionaries, visitedArrays, fontDictionaries);
        }

        int unembedded = 0;
        int uninspectable = 0;
        foreach (PdfDictionary font in fontDictionaries) {
            try {
                if (!HasEmbeddedFontProgram(font, document.Objects, document.ReadOptions.Limits.MaxDecodedStreamBytes)) unembedded++;
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is NotSupportedException ||
                exception is PdfReadLimitException) {
                uninspectable++;
            }
        }

        return new PdfPrintProductionStructureEvidence(
            document.Pages.Count,
            validBoxes,
            invalidBoxes,
            fontDictionaries.Count,
            unembedded,
            uninspectable);
    }

    private static bool HasValidProductionBoxes(PdfPageGeometry geometry) {
        PdfPageBox? media = geometry.MediaBox;
        PdfPageBox? trim = geometry.TrimBox;
        PdfPageBox? bleed = geometry.BleedBox;
        if (media == null || trim == null || bleed == null || geometry.ArtBox != null) return false;
        return Contains(media, bleed) && Contains(bleed, trim);
    }

    private static bool Contains(PdfPageBox outer, PdfPageBox inner) {
        const double tolerance = 0.0001D;
        return inner.Left >= outer.Left - tolerance &&
            inner.Bottom >= outer.Bottom - tolerance &&
            inner.Right <= outer.Right + tolerance &&
            inner.Top <= outer.Top + tolerance;
    }

    private static void CollectFontResources(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfDictionary> visited,
        HashSet<PdfArray> visitedArrays,
        HashSet<PdfDictionary> fonts) {
        PdfObject? resolved = value == null ? null : PdfObjectLookup.ResolveChain(objects, value);
        PdfDictionary? dictionary = resolved switch {
            PdfDictionary direct => direct,
            PdfStream stream => stream.Dictionary,
            _ => null
        };
        if (dictionary != null) {
            if (!visited.Add(dictionary)) return;
            if (dictionary.Items.TryGetValue("Font", out PdfObject? fontObject) &&
                PdfObjectLookup.ResolveChain(objects, fontObject) is PdfDictionary fontResources) {
                foreach (PdfObject resource in fontResources.Items.Values) {
                    if (PdfObjectLookup.ResolveChain(objects, resource) is PdfDictionary font) fonts.Add(font);
                }
            }

            foreach (PdfObject child in dictionary.Items.Values) {
                CollectFontResources(child, objects, visited, visitedArrays, fonts);
            }
            return;
        }

        if (resolved is PdfArray array) {
            if (!visitedArrays.Add(array)) return;
            foreach (PdfObject child in array.Items) CollectFontResources(child, objects, visited, visitedArrays, fonts);
        }
    }

    private static bool HasEmbeddedFontProgram(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects, int maxDecodedStreamBytes) {
        string? subtype = ResolveName(font.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null, objects);
        if (string.Equals(subtype, "Type3", StringComparison.Ordinal)) {
            return font.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) &&
                PdfObjectLookup.ResolveChain(objects, charProcsObject) is PdfDictionary charProcs &&
                charProcs.Items.Count > 0 &&
                charProcs.Items.Values.All(value => PdfObjectLookup.ResolveChain(objects, value) is PdfStream);
        }

        PdfDictionary fontWithDescriptor = font;
        if (string.Equals(subtype, "Type0", StringComparison.Ordinal) &&
            font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) &&
            PdfObjectLookup.ResolveChain(objects, descendantsObject) is PdfArray descendants &&
            descendants.Items.Count > 0 &&
            PdfObjectLookup.ResolveChain(objects, descendants.Items[0]) is PdfDictionary descendant) {
            fontWithDescriptor = descendant;
        }

        if (!fontWithDescriptor.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject) ||
            PdfObjectLookup.ResolveChain(objects, descriptorObject) is not PdfDictionary descriptor) return false;

        return HasReadableFontStream(descriptor, "FontFile", objects, maxDecodedStreamBytes) ||
            HasReadableFontStream(descriptor, "FontFile2", objects, maxDecodedStreamBytes) ||
            HasReadableFontStream(descriptor, "FontFile3", objects, maxDecodedStreamBytes);
    }

    private static bool HasReadableFontStream(PdfDictionary descriptor, string key, Dictionary<int, PdfIndirectObject> objects, int maxDecodedStreamBytes) {
        if (!descriptor.Items.TryGetValue(key, out PdfObject? value) ||
            PdfObjectLookup.ResolveChain(objects, value) is not PdfStream stream ||
            StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) return false;
        return StreamDecoder.TryDecode(
            stream.Dictionary,
            stream.Data,
            maxDecodedStreamBytes,
            out byte[] decoded,
            objects) && decoded.Length > 0;
    }

    private static string? ResolveName(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) =>
        value != null && PdfObjectLookup.ResolveChain(objects, value) is PdfName name ? name.Name : null;
}
