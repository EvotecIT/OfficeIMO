using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

/// <summary>
/// AI source declarations outside the C2PA attachment: IPTC digital source types and AI generator fields in the catalog XMP packet
/// and in image XMP packets, AI generator names in the Info dictionary's <c>Creator</c> and <c>Producer</c>, and every carrier
/// inside JPEG image streams (<c>DCTDecode</c>), which are complete JPEG files. Removal edits only those objects, in the same
/// verified full rewrite that removes C2PA attachments.
/// </summary>
public static partial class PdfProvenance {
    private const string CatalogMetadataLocation = "PDF/Catalog/Metadata";
    private static readonly string[] InfoGeneratorKeys = { "Creator", "Producer" };

    private delegate void ObjectGraphEdit(Dictionary<int, PdfIndirectObject> objects, PdfDocumentSecurityInfo security);

    /// <summary>Adds declaration evidence and returns the diagnostics of images that were preserved because they could not be inspected.</summary>
    private static string[] InspectDeclarations(
        PdfReadDocument document,
        OfficeProvenanceOptions options,
        List<OfficeProvenanceEvidence> evidence,
        out long prefixDecodedBytes) {
        var context = new OfficeProvenanceContext(OfficeProvenanceAssetFormat.Pdf, options);
        var tally = new DecodeTally();
        WalkDeclarations(document, options, context, removal: null, changes: null, edits: null, tally);
        prefixDecodedBytes = tally.Bytes;
        if (evidence.Count + context.Evidence.Count > options.MaxCarriers) {
            throw OfficeProvenanceLimitException.Create($"The asset exceeds the configured carrier limit of {options.MaxCarriers}.");
        }
        evidence.AddRange(context.Evidence);
        return context.Diagnostics.ToArray();
    }

    /// <summary>Plans the selected declaration removals; each edit replaces or trims one object in the rewritten graph.</summary>
    private static List<ObjectGraphEdit> PlanDeclarationRemoval(
        PdfReadDocument document,
        OfficeProvenanceRemovalOptions removal,
        List<OfficeProvenanceChange> changes) {
        var edits = new List<ObjectGraphEdit>();
        WalkDeclarations(document, removal.Limits, context: null, removal, changes, edits, new DecodeTally());
        return edits;
    }

    private static void WalkDeclarations(
        PdfReadDocument document,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        OfficeProvenanceRemovalOptions? removal,
        List<OfficeProvenanceChange>? changes,
        List<ObjectGraphEdit>? edits,
        DecodeTally tally) {
        CancellationToken cancellationToken = options.CancellationToken;
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        PdfDocumentSecurityInfo security = document.Security;
        bool removeDeclarations = removal?.RemoveAiSourceMetadata == true;

        if (security.RootObjectNumber is int rootNumber &&
            objects.TryGetValue(rootNumber, out PdfIndirectObject? root) && root.Value is PdfDictionary catalog &&
            catalog.Items.TryGetValue("Metadata", out PdfObject? catalogMetadata)) {
            WalkXmp(document, options, catalogMetadata, CatalogMetadataLocation, context, removeDeclarations ? removal : null, changes, edits,
                replaceDirect: (graph, graphSecurity, stream) => {
                    if (graphSecurity.RootObjectNumber is int number && graph.TryGetValue(number, out PdfIndirectObject? item) &&
                        item.Value is PdfDictionary graphCatalog) graphCatalog.Items["Metadata"] = stream;
                });
        }

        if (security.InfoObjectNumber is int infoNumber &&
            objects.TryGetValue(infoNumber, out PdfIndirectObject? infoObject) && infoObject.Value is PdfDictionary info) {
            foreach (string key in InfoGeneratorKeys) {
                if (!info.Items.TryGetValue(key, out PdfObject? item) ||
                    PdfObjectLookup.Resolve(objects, item) is not PdfStringObj text ||
                    !OfficeProvenanceAiGenerators.TryMatch(text.Value, out string product)) continue;
                string location = "PDF/Info/" + key;
                context?.Add(new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.AiGeneratorMetadata, location, true,
                    text.RawBytes.Length, product, OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia));
                if (!removeDeclarations || changes == null || edits == null) continue;
                changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.AiGeneratorMetadata, location, text.RawBytes.Length));
                string removedKey = key;
                edits.Add((graph, graphSecurity) => {
                    if (graphSecurity.InfoObjectNumber is int number && graph.TryGetValue(number, out PdfIndirectObject? graphInfo) &&
                        graphInfo.Value is PdfDictionary dictionary) dictionary.Items.Remove(removedKey);
                });
            }
        }

        bool processImages = removal?.ProcessEmbeddedAssets ?? options.ProcessEmbeddedAssets;
        int maximumImages = removal?.MaxEmbeddedAssets ?? options.MaxEmbeddedAssets;
        int images = 0;
        var visitedMetadata = new HashSet<int>();
        foreach (KeyValuePair<int, PdfIndirectObject> pair in objects.OrderBy(static item => item.Key).ToArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!processImages) break;
            if (pair.Value.Value is not PdfStream image || !IsName(objects, image.Dictionary, "Subtype", "Image")) continue;
            int objectNumber = pair.Key;
            bool hasMetadata = image.Dictionary.Items.TryGetValue("Metadata", out PdfObject? imageMetadata);
            bool isJpeg = IsJpegStream(objects, image);
            if (!hasMetadata && !isJpeg) continue;
            // Each image with a carrier counts once against the embedded-asset limit, before any of it is read.
            if (++images > maximumImages) throw new InvalidDataException("The PDF exceeds the configured embedded-asset limit.");
            if (hasMetadata && imageMetadata != null &&
                (imageMetadata is not PdfReference metadataReference || visitedMetadata.Add(metadataReference.ObjectNumber))) {
                WalkXmp(document, options, imageMetadata, $"PDF/Object[{objectNumber}]/Metadata", context,
                    removeDeclarations ? removal : null, changes, edits,
                    replaceDirect: (graph, _, stream) => {
                        if (graph.TryGetValue(objectNumber, out PdfIndirectObject? item) && item.Value is PdfStream graphImage) {
                            graphImage.Dictionary.Items["Metadata"] = stream;
                        }
                    });
            }
            if (!isJpeg) continue;
            WalkJpegImage(document, objectNumber, pair.Value.Generation, image, options, context, removal, changes, edits, tally);
        }
    }

    private static void WalkXmp(
        PdfReadDocument document,
        OfficeProvenanceOptions options,
        PdfObject metadata,
        string location,
        OfficeProvenanceContext? context,
        OfficeProvenanceRemovalOptions? removal,
        List<OfficeProvenanceChange>? changes,
        List<ObjectGraphEdit>? edits,
        Action<Dictionary<int, PdfIndirectObject>, PdfDocumentSecurityInfo, PdfStream> replaceDirect) {
        if (PdfObjectLookup.Resolve(document.Objects, metadata) is not PdfStream stream) return;
        byte[] packet;
        try {
            packet = document.DecodedStreamBudget.DecodeRequired(stream, document.Objects, PdfReadDocument.MaxXmpMetadataBytes, options.CancellationToken);
        } catch (InvalidDataException) {
            return;
        }
        if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location);
        if (removal == null || changes == null || edits == null ||
            !OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removal, location, changes, out byte[] cleaned)) return;
        PdfDictionary dictionary = PdfXmpMetadataSynchronizer.CloneUnfilteredMetadataDictionary(stream.Dictionary);
        if (metadata is PdfReference reference) {
            edits.Add((graph, _) => {
                int generation = graph.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? existing) ? existing.Generation : reference.Generation;
                graph[reference.ObjectNumber] = new PdfIndirectObject(reference.ObjectNumber, generation, new PdfStream(dictionary, cleaned));
            });
        } else {
            edits.Add((graph, security) => replaceDirect(graph, security, new PdfStream(dictionary, cleaned)));
        }
    }

    private static void WalkJpegImage(
        PdfReadDocument document,
        int objectNumber,
        int generation,
        PdfStream image,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        OfficeProvenanceRemovalOptions? removal,
        List<OfficeProvenanceChange>? changes,
        List<ObjectGraphEdit>? edits,
        DecodeTally tally) {
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        string prefix = $"PDF/Object[{objectNumber}]/Image/";
        long maximumBytes = options.MaxAssetBytes;
        bool chained = HasFilterChain(objects, image.Dictionary);
        if (chained) {
            // Decoding the filters ahead of DCTDecode expands data, so it draws from the same budget as every other decoded stream.
            long remaining = options.MaxExpandedContainerBytes - document.DecodedStreamBudget.UsedBytes - tally.Bytes;
            if (remaining <= 0) throw OfficeProvenanceLimitException.Create("The PDF exceeds the configured expanded-byte limit.");
            maximumBytes = Math.Min(maximumBytes, remaining);
        }
        // Filters ahead of DCTDecode (ASCII85, Flate) are decoded so the JPEG engine sees the complete JPEG file.
        if (!ResourceResolver.TryGetDctPayload(image, objects, (int)Math.Min(int.MaxValue, maximumBytes), out byte[] jpeg)) {
            context?.Diagnostics.Add($"{prefix.TrimEnd('/')}: image was preserved because its filters could not be decoded.");
            return;
        }
        if (jpeg.LongLength > options.MaxAssetBytes) throw new InvalidDataException("A PDF image exceeds the configured asset limit.");
        if (chained) tally.Bytes += jpeg.LongLength;
        string name = $"object-{objectNumber}.jpg";
        try {
            if (context != null) {
                OfficeProvenanceReport nested = OfficeProvenanceInspector.Inspect(jpeg, name, CreateNestedOptions(options));
                foreach (OfficeProvenanceEvidence item in nested.Evidence) {
                    context.Add(new OfficeProvenanceEvidence(item.Carrier, prefix + item.Location, item.IsStructurallyValid, item.PayloadLength,
                        item.Value, item.DigitalSourceKind).WithManifest(item.Manifest));
                }
            }
            if (removal == null || changes == null || edits == null) return;
            OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, name, CreateNestedRemoval(removal));
            if (!result.WasChanged) return;
            foreach (OfficeProvenanceChange change in result.Changes) {
                changes.Add(new OfficeProvenanceChange(change.Carrier, prefix + change.Location, change.RemovedBytes));
            }
            byte[] cleaned = result.ToArray();
            PdfDictionary dictionary = CreateJpegDictionary(image.Dictionary, objects);
            edits.Add((graph, _) => graph[objectNumber] = new PdfIndirectObject(objectNumber, generation, new PdfStream(dictionary, cleaned)));
        } catch (InvalidDataException exception) when (!OfficeProvenanceLimitException.Is(exception)) {
            // A JPEG stream the image reader cannot parse is left exactly as it is.
            context?.Diagnostics.Add($"{prefix.TrimEnd('/')}: image was preserved because inspection failed: {exception.Message}");
        }
    }

    /// <summary>Whether the image's last filter is DCTDecode, so its decoded payload is a JPEG file.</summary>
    private static bool IsJpegStream(Dictionary<int, PdfIndirectObject> objects, PdfStream image) {
        PdfObject? filter = image.Dictionary.Items.TryGetValue("Filter", out PdfObject? full) ? full
            : image.Dictionary.Items.TryGetValue("F", out PdfObject? abbreviated) ? abbreviated : null;
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, filter);
        if (resolved is PdfArray { Items.Count: > 0 } array) resolved = PdfObjectLookup.Resolve(objects, array.Items[array.Items.Count - 1]);
        return resolved is PdfName name && name.Name is "DCTDecode" or "DCT";
    }

    /// <summary>
    /// The image dictionary for a cleaned JPEG written without the filters that preceded DCTDecode: the filter becomes DCTDecode alone
    /// and only DCTDecode's own decode parameters stay.
    /// </summary>
    private static PdfDictionary CreateJpegDictionary(PdfDictionary source, Dictionary<int, PdfIndirectObject> objects) {
        var dictionary = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> entry in source.Items) {
            if (entry.Key is not ("Length" or "Filter" or "F" or "DecodeParms" or "DP")) dictionary.Items[entry.Key] = entry.Value;
        }
        dictionary.Items["Filter"] = new PdfName("DCTDecode");
        PdfObject? parameters = source.Items.TryGetValue("DecodeParms", out PdfObject? full) ? full
            : source.Items.TryGetValue("DP", out PdfObject? abbreviated) ? abbreviated : null;
        PdfObject? filter = source.Items.TryGetValue("Filter", out PdfObject? fullFilter) ? fullFilter
            : source.Items.TryGetValue("F", out PdfObject? abbreviatedFilter) ? abbreviatedFilter : null;
        // A filter array, even of one entry, pairs with a parameter array; DCTDecode's own parameters are its last entry.
        bool filterArray = PdfObjectLookup.Resolve(objects, filter) is PdfArray;
        PdfObject? resolvedParameters = PdfObjectLookup.Resolve(objects, parameters);
        if (!filterArray && parameters != null) {
            dictionary.Items["DecodeParms"] = parameters;
        } else if (filterArray && resolvedParameters is PdfArray { Items.Count: > 0 } list &&
            PdfObjectLookup.Resolve(objects, list.Items[list.Items.Count - 1]) is PdfDictionary dctParameters) {
            dictionary.Items["DecodeParms"] = dctParameters;
        }
        return dictionary;
    }

    private static bool HasFilterChain(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        PdfObject? filter = dictionary.Items.TryGetValue("Filter", out PdfObject? full) ? full
            : dictionary.Items.TryGetValue("F", out PdfObject? abbreviated) ? abbreviated : null;
        return PdfObjectLookup.Resolve(objects, filter) is PdfArray { Items.Count: > 1 };
    }

    /// <summary>Bytes decoded from filters ahead of DCTDecode during one inspection or removal.</summary>
    private sealed class DecodeTally {
        internal long Bytes { get; set; }
    }

    private static bool IsName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key, string expected) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && PdfObjectLookup.Resolve(objects, value) is PdfName name && name.Name == expected;

    private static OfficeProvenanceOptions CreateNestedOptions(OfficeProvenanceOptions source) => new OfficeProvenanceOptions {
        ProcessEmbeddedAssets = false,
        MaxAssetBytes = source.MaxAssetBytes,
        MaxManifestBytes = source.MaxManifestBytes,
        MaxCarriers = source.MaxCarriers,
        MaxContainerEntries = source.MaxContainerEntries,
        MaxExpandedContainerBytes = source.MaxExpandedContainerBytes,
        CancellationToken = source.CancellationToken
    };

    private static OfficeProvenanceRemovalOptions CreateNestedRemoval(OfficeProvenanceRemovalOptions source) {
        var nested = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            ProcessEmbeddedAssets = false
        };
        nested.Limits.MaxAssetBytes = source.Limits.MaxAssetBytes;
        nested.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        nested.Limits.MaxCarriers = source.Limits.MaxCarriers;
        nested.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        nested.Limits.MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes;
        nested.Limits.CancellationToken = source.Limits.CancellationToken;
        return nested;
    }
}
