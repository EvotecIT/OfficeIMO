namespace OfficeIMO.Pdf;

/// <summary>Walks the resource graph of a PDF page and records embedded color assets.</summary>
internal sealed class PdfAssetResourceInventory {
    private readonly PdfReadDocument _document;
    private readonly PdfAssetOutput? _output;
    private readonly PdfAssetReport? _report;
    private readonly PdfAssetInspectionBudget _work;
    private CancellationToken Token => _work.CancellationToken;
    private readonly Dictionary<PdfStream, ProfileFile?> _profiles = new();
    private long _decodedBytes;

    internal PdfAssetResourceInventory(PdfReadDocument document, PdfAssetOutput? output, PdfAssetReport? report, PdfAssetInspectionBudget work) {
        _document = document;
        _output = output;
        _report = report;
        _work = work;
    }

    internal void InspectOutputIntents() {
        if (_document.CatalogDictionary is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("OutputIntents", out PdfObject? value) ||
            Resolve(value) is not PdfArray intents) return;
        foreach (PdfObject item in intents.Items) {
            _work.Consume(1);
            if (Resolve(item) is PdfDictionary intent && intent.Items.TryGetValue("DestOutputProfile", out PdfObject? profile)) {
                AddProfile(null, "OutputIntent", profile);
            }
        }
    }

    internal IReadOnlyList<PdfDictionary> InspectPage(PdfReadPage page, int pageNumber, IReadOnlyList<PdfImagePlacement>? placements = null) {
        var resourceSets = new List<PdfDictionary>();
        var addedResourceSets = new HashSet<PdfDictionary>();
        bool hasUnmappedAppearance = false;
        var visited = new HashSet<PdfObject>();
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        var inks = new HashSet<(string Space, string Name)>();
        var profiles = new HashSet<PdfStream>();
        var stack = new Stack<PdfObject>();
        AddResources(page.GetFontInspectionResources());
        if (page.PageDictionary.Items.TryGetValue("Annots", out PdfObject? annotations) && Resolve(annotations) is PdfArray annots) {
            foreach (PdfObject item in annots.Items) {
                _work.Consume(1);
                if (Resolve(item) is PdfDictionary annotation && annotation.Items.TryGetValue("AP", out PdfObject? appearance)) {
                    hasUnmappedAppearance = true;
                    Schedule(appearance);
                }
            }
        }
        if (placements is not null) {
            foreach (PdfImagePlacement placement in placements) Schedule(placement.InlineImageStream?.Dictionary);
        }

        while (stack.Count > 0) {
            PdfObject value = stack.Pop();
            switch (value) {
                case PdfReference reference when PdfObjectLookup.TryGet(_document.Objects, reference, out PdfIndirectObject indirect):
                    Schedule(indirect.Value);
                    break;
                case PdfArray array:
                    if (_report is not null) InspectColorSpace(array);
                    foreach (PdfObject item in array.Items) Schedule(item);
                    break;
                case PdfStream stream:
                    Schedule(stream.Dictionary);
                    break;
                case PdfDictionary dictionary when dictionary.Get<PdfName>("Type")?.Name is not ("Catalog" or "Pages" or "Page"):
                    foreach (var entry in dictionary.Items) {
                        if (entry.Key == "Resources") AddResources(Resolve(entry.Value) as PdfDictionary);
                        else Schedule(entry.Value);
                    }
                    hasUnmappedAppearance |= dictionary.Items.ContainsKey("CharProcs") ||
                        dictionary.Get<PdfNumber>("PatternType")?.Value == 1 ||
                        dictionary.Get<PdfName>("S")?.Name is "Alpha" or "Luminosity";
                    break;
            }
        }

        if (hasUnmappedAppearance) {
            _report?.Diagnostics.Add(pageNumber, "Image placements", "AdditionalImagePlacementContexts",
                "Placement DPI is available for page and Form images. Annotation, pattern, glyph and graphics-state appearance resources have no placement DPI. Inline images inside those appearances may be omitted.");
        }
        return resourceSets;

        void AddResources(PdfDictionary? resources) {
            if (resources is not null && addedResourceSets.Add(resources)) {
                resourceSets.Add(resources);
                Schedule(resources);
            }
        }

        void Schedule(PdfObject? value) {
            _work.Consume(1);
            if (value is PdfReference reference) {
                if (visitedReferences.Add((reference.ObjectNumber, reference.Generation))) stack.Push(reference);
            } else if (value is PdfArray or PdfStream or PdfDictionary && visited.Add(value)) {
                stack.Push(value);
            }
        }

        void InspectColorSpace(PdfArray array) {
            if (array.Items.Count < 2 || Resolve(array.Items[0]) is not PdfName kind) return;
            switch (kind.Name) {
                case "ICCBased" when Resolve(array.Items[1]) is PdfStream profile && profiles.Add(profile):
                    AddProfile(pageNumber, "ICCBased", profile);
                    break;
                case "Separation" when array.Items.Count == 4:
                    AddInk(kind.Name, array.Items[1]);
                    break;
                case "DeviceN" when array.Items.Count is 4 or 5 && Resolve(array.Items[1]) is PdfArray names:
                    foreach (PdfObject name in names.Items) {
                        _work.Consume(1);
                        AddInk(kind.Name, name);
                    }
                    break;
            }
        }

        void AddInk(string space, PdfObject value) {
            if (Resolve(value) is PdfName name && name.Name is not ("All" or "None") && inks.Add((space, name.Name))) {
                _report!.Inks.Add(pageNumber, space, name.Name);
            }
        }
    }

    private void AddProfile(int? page, string source, PdfObject value) {
        if (Resolve(value) is not PdfStream profile) {
            _report!.Diagnostics.Add(page, source, "IccProfileUnavailable", "The embedded ICC profile stream is unavailable.");
            return;
        }

        if (!_profiles.TryGetValue(profile, out ProfileFile? file)) {
            PdfReadLimits limits = _document.ReadOptions.Limits;
            if (PdfImageStreamDecoder.TryDecode(profile, _document.Objects, out byte[] bytes,
                    limits.MaxDecodedStreamBytes, Token) && bytes.Length != 0) {
                _decodedBytes += bytes.LongLength;
                if (_decodedBytes > limits.MaxTotalDecodedStreamBytes) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, limits.MaxTotalDecodedStreamBytes, _decodedBytes);
                }
                string hash = PdfArtifactInfo.ComputeSha256Hex(bytes, Token);
                file = new ProfileFile(_output!.WriteDistinct("profiles", "icc", bytes, hash), bytes.Length, hash);
            }
            _profiles.Add(profile, file);
        }
        if (file is null) {
            _report!.Diagnostics.Add(page, source, "IccProfileUnavailable", "The embedded ICC profile could not be decoded.");
        } else {
            _report!.Profiles.Add(page, source, file.Path, file.Size, file.Hash);
        }
    }

    private PdfObject? Resolve(PdfObject value) => PdfObjectLookup.ResolveChain(_document.Objects, value);

    private sealed record ProfileFile(string Path, int Size, string Hash);
}
