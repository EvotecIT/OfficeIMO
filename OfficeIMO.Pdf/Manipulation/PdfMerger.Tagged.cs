using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    /// <summary>Retains each source's semantic graph until the shared structure owner can combine it.</summary>
    private static void CollectTaggedMergeObjects(PdfReadDocument document, PdfPageExtractor.ObjectCollector collector) {
        if (!PdfTaggedPageOrganizer.IsTagged(document)) return;
        var catalog = (PdfDictionary)document.Objects[document.Security.RootObjectNumber!.Value].Value;
        foreach (string name in new[] { "StructTreeRoot", "MarkInfo", "Lang" })
            if (catalog.Items.TryGetValue(name, out var value)) collector.CollectObjectGraph(value);
    }

    /// <summary>
    /// Combines the retained semantic graphs before catalog policies run. Object numbers stay fixed here:
    /// form and resource policies still need the source-to-output maps from the initial assembly.
    /// </summary>
    private static byte[] CompleteTaggedMerge(byte[] merged, IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex, int pagesId, int catalogId, int infoId, PdfFileVersion fileVersion,
        CancellationToken cancellationToken) {
        if (!sources.Any(source => PdfTaggedPageOrganizer.IsTagged(source.Document))) return merged;
        cancellationToken.ThrowIfCancellationRequested();
        if (sources[primarySourceIndex].Document.UncheckedXmpMetadata?.HasPdfUaIdentification == true &&
            sources.Any(source => !PdfTaggedPageOrganizer.IsTagged(source.Document)))
            throw new NotSupportedException("Merging untagged pages into a PDF/UA document requires accessible tagging of those pages first.");

        // Reuse the canonical parser and serializer. Keeping the entire intermediate graph also retains
        // form resources that the existing merge policy has not attached to the final catalog yet.
        int objectCount = sources.Sum(source => source.Collector.ObjectIds.Count) + 3;
        PdfLoadOptions readOptions = PdfLoadOptions.ForComposedOutput(sources[primarySourceIndex].Document.ReadOptions,
            sources.Select(source => source.Document.ReadOptions), merged.LongLength, objectCount);
        var read = PdfReadDocument.OpenBorrowed(merged, readOptions, cancellationToken);
        Dictionary<int, PdfIndirectObject> objects = read.Objects;
        var catalog = (PdfDictionary)objects[catalogId].Value;
        var precedingPages = new List<int>();
        foreach (ImportedSource source in sources) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyDictionary<int, int> map = source.OutputNumberMap!;
            var sourceCatalog = (PdfDictionary)source.Objects[source.Document.Security.RootObjectNumber!.Value].Value;
            if (PdfTaggedPageOrganizer.IsTagged(source.Document)) {
                var importedCatalog = new PdfDictionary();
                foreach (string name in new[] { "StructTreeRoot", "MarkInfo", "Lang" })
                    if (sourceCatalog.Items.TryGetValue(name, out var value))
                        importedCatalog.Items[name] = MapCatalogValue(value, map, cancellationToken);
                PdfTaggedPageOrganizer.MergeImportedStructure(objects, catalog, importedCatalog, map.Values,
                    precedingPages.ToArray(), precedingPages.Count + 1, cancellationToken);
            }
            precedingPages.AddRange(source.PageObjectNumbers.Select(number => map[number]));
        }

        PdfTaggedPageOrganizer.PruneStructure(objects, catalog, new HashSet<int>(), new HashSet<int>(), cancellationToken);
        var identity = objects.Keys.ToDictionary(number => number, number => number);
        var context = new PdfPageExtractor.SerializationContext(identity, pagesId,
            new Dictionary<int, Dictionary<string, PdfObject>>(), objects,
            preserveRawStringBytes: true, cancellationToken: cancellationToken);
        var serialized = new List<PdfSerializedObject>(objects.Count);
        foreach (var pair in objects.OrderBy(pair => pair.Key)) {
            cancellationToken.ThrowIfCancellationRequested();
            serialized.Add(PdfPageExtractor.SerializeIndirectObjectForAssembly(pair.Key, pair.Value.Value, context));
        }
        return PdfPageExtractor.Assemble(serialized, catalogId, infoId, fileVersion, cancellationToken);
    }

    private static void ValidateTaggedFormMergePolicy(IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex, PdfMergePolicy policy) {
        bool taggedFields = sources.Any(source => source.FormFieldCount > 0 && PdfTaggedPageOrganizer.IsTagged(source.Document));
        if (!taggedFields) return;
        bool dropsTaggedFields = policy.Forms == PdfMergeStructureMode.Drop ||
            policy.Forms == PdfMergeStructureMode.KeepPrimary && sources.Where((source, index) => index != primarySourceIndex)
                .Any(source => source.FormFieldCount > 0 && PdfTaggedPageOrganizer.IsTagged(source.Document));
        if (dropsTaggedFields || policy.Forms == PdfMergeStructureMode.Combine && policy.FormFieldCollisions == PdfMergeCollisionMode.KeepFirst)
            throw new NotSupportedException("Tagged form fields must keep their accessibility links. Use a Combine form policy with RenameIncoming or Reject collisions.");
    }

    private static PdfObject MapCatalogValue(PdfObject value, IReadOnlyDictionary<int, int> map,
        CancellationToken cancellationToken, int depth = 0) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > 256) throw new InvalidOperationException("Tagged merge catalog nesting exceeds the supported limit.");
        if (value is PdfReference reference) {
            if (!map.TryGetValue(reference.ObjectNumber, out int number))
                throw new InvalidOperationException("A tagged merge catalog entry references an object that was not retained.");
            return new PdfReference(number, 0);
        }
        if (value is PdfDictionary dictionary) {
            var copy = new PdfDictionary();
            foreach (var pair in dictionary.Items) copy.Items[pair.Key] = MapCatalogValue(pair.Value, map, cancellationToken, depth + 1);
            return copy;
        }
        if (value is PdfArray array) {
            var copy = new PdfArray();
            foreach (var item in array.Items) copy.Items.Add(MapCatalogValue(item, map, cancellationToken, depth + 1));
            return copy;
        }
        if (value is PdfStream) throw new InvalidOperationException("Tagged merge catalog entries must be dictionaries or values, not streams.");
        return value;
    }
}
