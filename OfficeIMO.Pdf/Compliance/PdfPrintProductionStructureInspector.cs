using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionStructureInspector {
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

        ReachableFontInspection fontInspection = InspectReachableFonts(document, cancellationToken);
        HashSet<PdfDictionary> fontDictionaries = fontInspection.Fonts;

        int unembedded = 0;
        int uninspectable = fontInspection.UninspectableContextCount;
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
        PdfPageBox? art = geometry.ArtBox;
        if (media == null || (trim == null) == (art == null)) return false;
        PdfPageBox productionBoundary = trim ?? art!;
        if (!Contains(media, productionBoundary)) return false;
        return bleed == null ||
            (Contains(media, bleed) && Contains(bleed, productionBoundary));
    }

    private static bool Contains(PdfPageBox outer, PdfPageBox inner) {
        const double tolerance = 0.0001D;
        return inner.Left >= outer.Left - tolerance &&
            inner.Bottom >= outer.Bottom - tolerance &&
            inner.Right <= outer.Right + tolerance &&
            inner.Top <= outer.Top + tolerance;
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
        string? programFontSubtype = subtype;
        if (string.Equals(subtype, "Type0", StringComparison.Ordinal)) {
            if (!font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) ||
                ResolveObject(objects, descendantsObject, 0, maximumObjectDepth, out _) is not PdfArray descendants ||
                descendants.Items.Count != 1 ||
                ResolveObject(objects, descendants.Items[0], 0, maximumObjectDepth, out _) is not PdfDictionary descendant) {
                return false;
            }
            string? descendantSubtype = ResolveName(
                descendant.Items.TryGetValue("Subtype", out PdfObject? descendantSubtypeObject)
                    ? descendantSubtypeObject
                    : null,
                objects,
                maximumObjectDepth);
            if (!string.Equals(descendantSubtype, "CIDFontType0", StringComparison.Ordinal) &&
                !string.Equals(descendantSubtype, "CIDFontType2", StringComparison.Ordinal)) {
                return false;
            }
            fontWithDescriptor = descendant;
            programFontSubtype = descendantSubtype;
        }

        if (!fontWithDescriptor.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject) ||
            ResolveObject(objects, descriptorObject, 0, maximumObjectDepth, out _) is not PdfDictionary descriptor) return false;

        return HasReadableFontStream(
                programFontSubtype,
                descriptor,
                "FontFile",
                objects,
                maxDecodedStreamBytes,
                maximumObjectDepth) ||
            HasReadableFontStream(
                programFontSubtype,
                descriptor,
                "FontFile2",
                objects,
                maxDecodedStreamBytes,
                maximumObjectDepth) ||
            HasReadableFontStream(
                programFontSubtype,
                descriptor,
                "FontFile3",
                objects,
                maxDecodedStreamBytes,
                maximumObjectDepth);
    }

    private static bool HasReadableFontStream(
        string? fontSubtype,
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
            objects) && IsValidFontProgram(fontSubtype, key, stream, decoded, objects, maximumObjectDepth);
    }

    private static string? ResolveName(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        value != null && ResolveObject(objects, value, 0, maximumObjectDepth, out _) is PdfName name ? name.Name : null;
}
