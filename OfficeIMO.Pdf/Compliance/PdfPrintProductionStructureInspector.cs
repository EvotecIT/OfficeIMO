using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static class PdfPrintProductionStructureInspector {
    internal static PdfPrintProductionStructureEvidence Inspect(
        PdfReadDocument document,
        System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        int validBoxes = 0;
        int invalidBoxes = 0;
        foreach (PdfReadPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            if (HasValidProductionBoxes(page.GetGeometry())) validBoxes++;
            else invalidBoxes++;
        }

        var fontDictionaries = new HashSet<PdfDictionary>();
        var visitedDictionaries = new HashSet<PdfDictionary>();
        var visitedArrays = new HashSet<PdfArray>();
        foreach (PdfIndirectObject indirect in document.Objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            CollectFontResources(
                indirect.Value,
                document.Objects,
                visitedDictionaries,
                visitedArrays,
                fontDictionaries,
                document.ReadOptions.Limits.MaxObjectNestingDepth,
                cancellationToken);
        }

        int unembedded = 0;
        int uninspectable = 0;
        foreach (PdfDictionary font in fontDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                if (!HasEmbeddedFontProgram(
                        font,
                        document.Objects,
                        document.ReadOptions.Limits.MaxDecodedStreamBytes,
                        document.ReadOptions.Limits.MaxObjectNestingDepth)) unembedded++;
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
        HashSet<PdfDictionary> fonts,
        int maximumObjectDepth,
        System.Threading.CancellationToken cancellationToken) {
        if (value == null) return;
        var pending = new Stack<(PdfObject Value, int Depth)>();
        pending.Push((value, 0));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            (PdfObject candidate, int depth) = pending.Pop();
            PdfObject? resolved = ResolveObject(objects, candidate, depth, maximumObjectDepth, out int resolvedDepth);
            PdfDictionary? dictionary = resolved switch {
                PdfDictionary direct => direct,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (dictionary != null) {
                if (!visited.Add(dictionary)) continue;
                if (dictionary.Items.TryGetValue("Font", out PdfObject? fontObject) &&
                    ResolveObject(objects, fontObject, resolvedDepth + 1, maximumObjectDepth, out _) is PdfDictionary fontResources) {
                    foreach (PdfObject resource in fontResources.Items.Values) {
                        if (ResolveObject(objects, resource, resolvedDepth + 1, maximumObjectDepth, out _) is PdfDictionary font) {
                            fonts.Add(font);
                        }
                    }
                }

                foreach (PdfObject child in dictionary.Items.Values) {
                    pending.Push((child, resolvedDepth + 1));
                }
                continue;
            }

            if (resolved is PdfArray array && visitedArrays.Add(array)) {
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            }
        }
    }

    private static PdfObject? ResolveObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        int depth,
        int maximumObjectDepth,
        out int resolvedDepth) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        PdfObject? resolved = value;
        resolvedDepth = depth;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) return null;
            resolvedDepth++;
            if (resolvedDepth > maximumObjectDepth) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.ObjectNestingDepth,
                    maximumObjectDepth,
                    resolvedDepth);
            }
            resolved = indirect.Value;
        }
        return resolved;
    }

    private static bool HasEmbeddedFontProgram(
        PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int maximumObjectDepth) {
        string? subtype = ResolveName(
            font.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null,
            objects,
            maximumObjectDepth);
        if (string.Equals(subtype, "Type3", StringComparison.Ordinal)) {
            return font.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) &&
                ResolveObject(objects, charProcsObject, 0, maximumObjectDepth, out _) is PdfDictionary charProcs &&
                charProcs.Items.Count > 0 &&
                charProcs.Items.Values.All(value =>
                    ResolveObject(objects, value, 0, maximumObjectDepth, out _) is PdfStream);
        }

        PdfDictionary fontWithDescriptor = font;
        if (string.Equals(subtype, "Type0", StringComparison.Ordinal) &&
            font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) &&
            ResolveObject(objects, descendantsObject, 0, maximumObjectDepth, out _) is PdfArray descendants &&
            descendants.Items.Count > 0 &&
            ResolveObject(objects, descendants.Items[0], 0, maximumObjectDepth, out _) is PdfDictionary descendant) {
            fontWithDescriptor = descendant;
        }

        if (!fontWithDescriptor.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject) ||
            ResolveObject(objects, descriptorObject, 0, maximumObjectDepth, out _) is not PdfDictionary descriptor) return false;

        return HasReadableFontStream(descriptor, "FontFile", objects, maxDecodedStreamBytes, maximumObjectDepth) ||
            HasReadableFontStream(descriptor, "FontFile2", objects, maxDecodedStreamBytes, maximumObjectDepth) ||
            HasReadableFontStream(descriptor, "FontFile3", objects, maxDecodedStreamBytes, maximumObjectDepth);
    }

    private static bool HasReadableFontStream(
        PdfDictionary descriptor,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int maximumObjectDepth) {
        if (!descriptor.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(objects, value, 0, maximumObjectDepth, out _) is not PdfStream stream ||
            StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) return false;
        return StreamDecoder.TryDecode(
            stream.Dictionary,
            stream.Data,
            maxDecodedStreamBytes,
            out byte[] decoded,
            objects) && decoded.Length > 0;
    }

    private static string? ResolveName(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        value != null && ResolveObject(objects, value, 0, maximumObjectDepth, out _) is PdfName name ? name.Name : null;
}
