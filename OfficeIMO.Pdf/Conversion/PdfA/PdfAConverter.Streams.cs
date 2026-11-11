using System.Globalization;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private readonly List<(PdfDictionary FileSpec, string Location)> _fileSpecifications = new();
    private readonly HashSet<PdfObject> _associatedFiles = new(ReferenceComparer.Instance);

    /// <summary>Visits every object reachable from the catalog and Info dictionary once.</summary>
    private void ScanObjects(PdfDictionary catalog) {
        var visited = new HashSet<PdfObject>(ReferenceComparer.Instance);
        var pending = new Stack<(PdfObject Value, int? ObjectNumber)>();
        pending.Push((catalog, ObjectNumberOf(catalog)));
        if (_document.Security.InfoObjectNumber is int info && _objects.TryGetValue(info, out PdfIndirectObject? infoObject)) {
            pending.Push((infoObject.Value, info));
        }
        var limitReported = new HashSet<int>();
        while (pending.Count > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            (PdfObject value, int? objectNumber) = pending.Pop();
            if (value is PdfReference reference) {
                if (PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject indirect)) pending.Push((indirect.Value, reference.ObjectNumber));
                continue;
            }
            if (_part != 4 && !CheckValueLimit(value) && objectNumber is int limitObject && limitReported.Add(limitObject)) {
                AddViolation(PdfAConversionRefusalReason.ImplementationLimitExceeded, "object " + limitObject.ToString(CultureInfo.InvariantCulture),
                    "The object holds a number, string or name beyond the PDF/A implementation limits.", limitObject);
            }
            if (value is PdfArray array) {
                if (!visited.Add(array)) continue;
                foreach (PdfObject item in array.Items) pending.Push((item, objectNumber));
                continue;
            }
            PdfStream? stream = value as PdfStream;
            PdfDictionary? dictionary = stream?.Dictionary ?? value as PdfDictionary;
            if (dictionary is null || !visited.Add(dictionary)) continue;

            if (stream is not null) ScanStream(stream, objectNumber);
            ScanDictionary(dictionary, stream, objectNumber, catalog);
            foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
                if (stream is not null && entry.Key == "Length") continue;
                if (entry.Key == "Parent" || entry.Key == "P") {
                    // Back references reach objects already reachable through their owners.
                    if (entry.Value is PdfReference) continue;
                }
                pending.Push((entry.Value, entry.Value is PdfReference child ? child.ObjectNumber : objectNumber));
            }
        }
    }

    private void ScanDictionary(PdfDictionary dictionary, PdfStream? stream, int? objectNumber, PdfDictionary catalog) {
        string location = "object " + (objectNumber?.ToString(CultureInfo.InvariantCulture) ?? "(direct)");
        string? type = GetName(dictionary, "Type");
        string? subtype = GetName(dictionary, "Subtype");
        if (stream is not null && subtype == "Image") {
            if (Resolve(dictionary.Items.TryGetValue("Interpolate", out PdfObject? interpolate) ? interpolate : null) is PdfBoolean { Value: true } &&
                Change(PdfAConversionChangeKind.XObjectUpdated, location + " (image)",
                    "Turned off image interpolation, which PDF/A forbids. Image pixels and placement are unchanged.", stream, OfficeConversionLossKind.Approximation)) {
                dictionary.Items["Interpolate"] = new PdfBoolean(false);
            }
            RemoveXObjectEntry(dictionary, stream, "Alternates", location + " (image)", "Removed alternate images, which PDF/A forbids; the image drawn on the page is kept.");
            RemoveXObjectEntry(dictionary, stream, "OPI", location + " (image)", "Removed OPI print-proxy information, which PDF/A forbids.");
        } else if (stream is not null && subtype == "Form") {
            RemoveXObjectEntry(dictionary, stream, "OPI", location + " (form)", "Removed OPI print-proxy information, which PDF/A forbids.");
            RemoveXObjectEntry(dictionary, stream, "PS", location + " (form)", "Removed a PostScript print fragment, which PDF/A forbids; viewers do not draw it.");
            if (GetName(dictionary, "Subtype2") == "PS") {
                AddViolation(PdfAConversionRefusalReason.ForbiddenXObject, location, "A form XObject is a PostScript XObject (Subtype2 /PS), which PDF/A forbids.", stream);
            }
        }

        if (dictionary.Items.ContainsKey("EF") && (type is null or "Filespec" or "F")) {
            _fileSpecifications.Add((dictionary, location));
        }
        if (dictionary.Items.TryGetValue("AF", out PdfObject? associated)) {
            PdfObject? resolved = Resolve(associated);
            IEnumerable<PdfObject> entries = resolved is PdfArray array ? array.Items : new[] { associated };
            foreach (PdfObject entry in entries) {
                if (ResolveDictionary(entry) is PdfDictionary fileSpec) _associatedFiles.Add(fileSpec);
            }
        }
        if (_part == 4 && Resolve(dictionary.Items.TryGetValue("ActualText", out PdfObject? actualText) ? actualText : null) is PdfStringObj replacement &&
            ContainsPrivateUse(replacement.Value) &&
            Change(PdfAConversionChangeKind.StructureUpdated, location,
                "Removed replacement text (ActualText) made of private-use characters, which PDF/A-4 forbids; the page draws the same.", dictionary, OfficeConversionLossKind.Omission)) {
            dictionary.Items.Remove("ActualText");
        }
        if (!ReferenceEquals(dictionary, catalog) && dictionary.Items.TryGetValue("Metadata", out PdfObject? objectMetadata) &&
            Resolve(objectMetadata) is PdfStream &&
            Change(PdfAConversionChangeKind.StructureUpdated, location,
                "Removed object-level XMP metadata, whose properties PDF/A cannot validate without extension schemas. Document metadata is kept.",
                dictionary, OfficeConversionLossKind.Omission)) {
            dictionary.Items.Remove("Metadata");
        }
    }

    private static bool ContainsPrivateUse(string text) {
        for (int index = 0; index < text.Length; index++) {
            char character = text[index];
            if (character >= '\uE000' && character <= '\uF8FF') return true;
            if (char.IsHighSurrogate(character) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                int scalar = char.ConvertToUtf32(character, text[index + 1]);
                if (scalar >= 0xF0000) return true;
                index++;
            }
        }
        return false;
    }

    private void RemoveXObjectEntry(PdfDictionary dictionary, PdfStream owner, string key, string location, string message) {
        if (dictionary.Items.ContainsKey(key) && Change(PdfAConversionChangeKind.XObjectUpdated, location, message, owner, OfficeConversionLossKind.Omission)) {
            dictionary.Items.Remove(key);
        }
    }

    private void ScanStream(PdfStream stream, int? objectNumber) {
        PdfDictionary dictionary = stream.Dictionary;
        string location = "object " + (objectNumber?.ToString(CultureInfo.InvariantCulture) ?? "(direct)");
        if (dictionary.Items.ContainsKey("F") || dictionary.Items.ContainsKey("FFilter") || dictionary.Items.ContainsKey("FDecodeParms")) {
            AddViolation(PdfAConversionRefusalReason.UnsupportedStream, location, "The stream takes its data from an external file, which PDF/A forbids.", stream);
            return;
        }
        if (!dictionary.Items.TryGetValue("Filter", out PdfObject? filterValue)) return;
        PdfObject? resolvedFilter = Resolve(filterValue);
        var filters = new List<string>();
        if (resolvedFilter is PdfName single) {
            filters.Add(single.Name);
        } else if (resolvedFilter is PdfArray array) {
            foreach (PdfObject item in array.Items) filters.Add(NameOf(Resolve(item)) ?? string.Empty);
        }
        int crypt = filters.IndexOf("Crypt");
        if (crypt >= 0) {
            AddViolation(PdfAConversionRefusalReason.UnsupportedStream, location, "The stream uses a Crypt filter, which PDF/A forbids.", stream);
            return;
        }
        int lastLzw = filters.LastIndexOf("LZWDecode");
        if (lastLzw < 0) return;
        for (int index = 0; index <= lastLzw; index++) {
            if (filters[index] is not ("LZWDecode" or "FlateDecode" or "ASCIIHexDecode" or "ASCII85Decode" or "RunLengthDecode")) {
                AddViolation(PdfAConversionRefusalReason.UnsupportedStream, location,
                    "The stream is LZW-compressed behind a " + filters[index] + " filter and cannot be re-encoded without decoding the image.", stream);
                return;
            }
        }

        var prefix = new PdfDictionary();
        PdfObject? parameters = dictionary.Items.TryGetValue("DecodeParms", out PdfObject? parms) ? Resolve(parms) : null;
        var prefixFilters = new PdfArray();
        var prefixParameters = new PdfArray();
        for (int index = 0; index <= lastLzw; index++) {
            prefixFilters.Items.Add(new PdfName(filters[index]));
            prefixParameters.Items.Add((PdfObject?)GetFilterParameter(parameters, index, filters.Count) ?? PdfNull.Instance);
        }
        prefix.Items["Filter"] = prefixFilters;
        prefix.Items["DecodeParms"] = prefixParameters;
        if (!StreamDecoder.TryDecode(prefix, stream.Data, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] decoded, _objects)) {
            AddViolation(PdfAConversionRefusalReason.UnsupportedStream, location, "The LZW-compressed stream cannot be decoded for re-encoding.", stream);
            return;
        }
        if (!Change(PdfAConversionChangeKind.StreamReencoded, location, "Re-encoded an LZW-compressed stream with Flate, as PDF/A forbids LZW. The decoded data is unchanged.", stream)) {
            return;
        }
        var remainingFilters = new PdfArray();
        var remainingParameters = new PdfArray();
        remainingFilters.Items.Add(new PdfName("FlateDecode"));
        remainingParameters.Items.Add(PdfNull.Instance);
        bool hasParameters = false;
        for (int index = lastLzw + 1; index < filters.Count; index++) {
            remainingFilters.Items.Add(new PdfName(filters[index]));
            PdfObject? parameter = GetFilterParameter(parameters, index, filters.Count);
            hasParameters |= parameter is not null;
            remainingParameters.Items.Add(parameter ?? PdfNull.Instance);
        }
        stream.Dictionary.Items.Remove("DecodeParms");
        if (hasParameters) stream.Dictionary.Items["DecodeParms"] = remainingParameters;
        ReplaceStreamData(stream, objectNumber, PdfFlateEncoder.Compress(decoded, _cancellationToken),
            remainingFilters.Items.Count == 1 ? new PdfName("FlateDecode") : remainingFilters);
    }

    private PdfDictionary? GetFilterParameter(PdfObject? parameters, int index, int filterCount) {
        if (parameters is PdfArray array) return index < array.Items.Count && Resolve(array.Items[index]) is PdfDictionary dictionary ? dictionary : null;
        return filterCount == 1 && index == 0 && parameters is PdfDictionary single ? single : null;
    }

    /// <summary>Replaces a stream's data. Streams are immutable, so the owning indirect object is replaced.</summary>
    private void ReplaceStreamData(PdfStream stream, int? objectNumber, byte[] data, PdfObject? filter) {
        stream.Dictionary.Items.Remove("Filter");
        if (filter is null) stream.Dictionary.Items.Remove("DecodeParms");
        else stream.Dictionary.Items["Filter"] = filter;
        stream.Dictionary.Items.Remove("DL");
        var replacement = new PdfStream(stream.Dictionary, data);
        if (objectNumber is int number && _objects.TryGetValue(number, out PdfIndirectObject? indirect) && ReferenceEquals(indirect.Value, stream)) {
            _objects[number] = new PdfIndirectObject(number, indirect.Generation, replacement);
            _objectNumbers[replacement] = number;
        } else {
            AddViolation(PdfAConversionRefusalReason.UnsupportedStream, "object " + (objectNumber?.ToString(CultureInfo.InvariantCulture) ?? "(direct)"),
                "A direct stream could not be re-encoded in place.", objectNumber);
        }
    }

    private static bool CheckValueLimit(PdfObject value) {
        switch (value) {
            case PdfNumber number:
                return !IsBeyondNumberLimit(number.Value);
            case PdfStringObj text:
                return text.RawBytes.Length <= MaximumStringBytes;
            case PdfName name:
                return Encoding.UTF8.GetByteCount(name.Name) <= MaximumNameBytes;
            case PdfDictionary dictionary:
                foreach (string key in dictionary.Items.Keys) {
                    if (Encoding.UTF8.GetByteCount(key) > MaximumNameBytes) return false;
                }
                return true;
            case PdfStream stream:
                return CheckValueLimit(stream.Dictionary);
            default:
                return true;
        }
    }

    /// <summary>
    /// PDF/A-2 and PDF/A-4 keep only attachments that are PDF/A themselves, so attachments are removed; PDF/A-3 and PDF/A-4f
    /// keep all, with their relationships. PDF/A-4f also requires at least one attachment in the embedded-files name tree.
    /// </summary>
    private void ProcessEmbeddedFiles(PdfDictionary catalog) {
        PdfDictionary? names = ResolveDictionary(catalog.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null);
        string partName = PdfAConversionReport.ToProfileName(_options.Profile);
        if (_part == 2 || (_part == 4 && !_attachmentsProfile)) {
            foreach ((PdfDictionary fileSpec, string location) in _fileSpecifications) {
                string fileName = DescribeFileSpecification(fileSpec);
                if (!_options.RemoveEmbeddedFiles) {
                    AddViolation(PdfAConversionRefusalReason.EmbeddedFileNotAllowed, location,
                        "The embedded file " + fileName + " cannot be kept in " + partName + ", and RemoveEmbeddedFiles is false.", fileSpec);
                    continue;
                }
                if (Change(PdfAConversionChangeKind.EmbeddedFileRemoved, location,
                        "Removed the embedded file " + fileName + "; " + partName + " cannot carry it. Choose " + (_part == 4 ? "PDF/A-4f" : "PDF/A-3") +
                        " to keep attachments.", fileSpec, OfficeConversionLossKind.Omission)) {
                    fileSpec.Items.Remove("EF");
                    fileSpec.Items.Remove("RF");
                }
            }
            if (_options.RemoveEmbeddedFiles && names is not null && names.Items.ContainsKey("EmbeddedFiles") &&
                Change(PdfAConversionChangeKind.EmbeddedFileRemoved, "catalog names", "Removed the embedded-files name tree.", names, OfficeConversionLossKind.Omission)) {
                names.Items.Remove("EmbeddedFiles");
            }
            if (_options.RemoveEmbeddedFiles && catalog.Items.ContainsKey("Collection") &&
                Change(PdfAConversionChangeKind.StructureUpdated, "catalog", "Removed the portfolio (Collection) view of the removed attachments.", catalog, OfficeConversionLossKind.Omission)) {
                catalog.Items.Remove("Collection");
            }
            return;
        }

        PdfArray? catalogFiles = Resolve(catalog.Items.TryGetValue("AF", out PdfObject? af) ? af : null) as PdfArray;
        foreach ((PdfDictionary fileSpec, string location) in _fileSpecifications) {
            string fileName = DescribeFileSpecification(fileSpec);
            string? f = TextOf(fileSpec, "F");
            string? uf = TextOf(fileSpec, "UF");
            if ((f is null || uf is null) && (f ?? uf) is string known &&
                Change(PdfAConversionChangeKind.EmbeddedFileUpdated, location, "Added the " + (f is null ? "F" : "UF") + " file name of " + fileName + ", as " + partName + " requires.", fileSpec)) {
                if (f is null) fileSpec.Items["F"] = new PdfStringObj(known);
                else fileSpec.Items["UF"] = new PdfStringObj(known, useTextStringEncoding: true);
            }
            if (!fileSpec.Items.ContainsKey("AFRelationship") &&
                Change(PdfAConversionChangeKind.EmbeddedFileUpdated, location, "Marked the relationship of " + fileName + " as Unspecified, as " + partName + " requires.", fileSpec)) {
                fileSpec.Items["AFRelationship"] = new PdfName("Unspecified");
            }
            if (ResolveDictionary(fileSpec.Items.TryGetValue("EF", out PdfObject? ef) ? ef : null) is PdfDictionary embedded) {
                foreach (PdfObject fileValue in embedded.Items.Values) {
                    if (Resolve(fileValue) is PdfStream file && GetName(file.Dictionary, "Subtype") is null &&
                        Change(PdfAConversionChangeKind.EmbeddedFileUpdated, location, "Marked the media type of " + fileName + " as application/octet-stream, as " + partName + " requires.", file)) {
                        file.Dictionary.Items["Subtype"] = new PdfName("application/octet-stream");
                    }
                }
            }
            if (!_associatedFiles.Contains(fileSpec) &&
                Change(PdfAConversionChangeKind.EmbeddedFileUpdated, location, "Associated " + fileName + " with the document (AF), as " + partName + " requires.", fileSpec)) {
                if (catalogFiles is null) {
                    catalogFiles = new PdfArray();
                    catalog.Items["AF"] = catalogFiles;
                }
                int? number = ObjectNumberOf(fileSpec);
                catalogFiles.Items.Add(number is int fileSpecNumber ? new PdfReference(fileSpecNumber, _objects[fileSpecNumber].Generation) : AddObject(fileSpec));
                _associatedFiles.Add(fileSpec);
            }
        }
        if (_attachmentsProfile) EnsureEmbeddedFilesNameTree(catalog, names);
    }

    /// <summary>
    /// PDF/A-4f requires the EmbeddedFiles name tree (ISO 19005-4, Annex A). Attachments reachable only from annotations are
    /// listed in a new tree; a source without any attachment is refused, because the converter cannot invent one.
    /// </summary>
    private void EnsureEmbeddedFilesNameTree(PdfDictionary catalog, PdfDictionary? names) {
        if (names is not null && ResolveDictionary(names.Items.TryGetValue("EmbeddedFiles", out PdfObject? tree) ? tree : null) is PdfDictionary existing &&
            NameTreeHasEntries(existing, 0)) return;
        if (_fileSpecifications.Count == 0) {
            AddViolation(PdfAConversionRefusalReason.EmbeddedFileRequired, "catalog names",
                "PDF/A-4f requires at least one embedded file, and the PDF has none. Choose PDF/A-4 for a file without attachments.", catalog);
            return;
        }
        if (!Change(PdfAConversionChangeKind.EmbeddedFileUpdated, "catalog names",
                "Listed " + _fileSpecifications.Count.ToString(CultureInfo.InvariantCulture) +
                " attachment(s) in the embedded-files name tree, as PDF/A-4f requires.", catalog)) return;
        var entries = new SortedDictionary<string, PdfObject>(StringComparer.Ordinal);
        foreach ((PdfDictionary fileSpec, _) in _fileSpecifications) {
            string baseName = TextOf(fileSpec, "UF") ?? TextOf(fileSpec, "F") ?? "attachment";
            string key = baseName;
            for (int suffix = 2; entries.ContainsKey(key); suffix++) key = baseName + " (" + suffix.ToString(CultureInfo.InvariantCulture) + ")";
            int? number = ObjectNumberOf(fileSpec);
            entries[key] = number is int fileSpecNumber ? new PdfReference(fileSpecNumber, _objects[fileSpecNumber].Generation) : AddObject(fileSpec);
        }
        var array = new PdfArray();
        foreach (KeyValuePair<string, PdfObject> entry in entries) {
            array.Items.Add(new PdfStringObj(entry.Key, useTextStringEncoding: true));
            array.Items.Add(entry.Value);
        }
        var nameTree = new PdfDictionary();
        nameTree.Items["Names"] = array;
        if (names is null) {
            names = new PdfDictionary();
            catalog.Items["Names"] = names;
        }
        names.Items["EmbeddedFiles"] = AddObject(nameTree);
    }

    private bool NameTreeHasEntries(PdfDictionary node, int depth) {
        if (depth > 32) return false;
        if (Resolve(node.Items.TryGetValue("Names", out PdfObject? values) ? values : null) is PdfArray { Items.Count: >= 2 }) return true;
        if (Resolve(node.Items.TryGetValue("Kids", out PdfObject? kids) ? kids : null) is not PdfArray children) return false;
        foreach (PdfObject child in children.Items) {
            if (ResolveDictionary(child) is PdfDictionary kid && NameTreeHasEntries(kid, depth + 1)) return true;
        }
        return false;
    }

    private string? TextOf(PdfDictionary dictionary, string key) =>
        Resolve(dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null) is PdfStringObj text ? text.Value : null;

    private string DescribeFileSpecification(PdfDictionary fileSpec) {
        string? name = TextOf(fileSpec, "UF") ?? TextOf(fileSpec, "F");
        return name is null ? "(unnamed)" : "'" + name + "'";
    }
}
