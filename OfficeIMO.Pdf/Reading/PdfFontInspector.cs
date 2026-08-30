using System.Runtime.CompilerServices;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static class PdfFontInspector {
    internal static PdfFontInventory Inspect(
        PdfReadDocument document,
        PdfFontInspectionOptions? options = null,
        PdfPageSelection? selection = null) {
        Guard.NotNull(document, nameof(document));
        PdfFontInspectionOptions effective = PdfFontInspectionOptions.Resolve(options);
        document.DemandContentExtraction("font resource");

        int[] pageNumbers = selection is null
            ? Enumerable.Range(1, document.Pages.Count).ToArray()
            : selection.ToPageNumbers(document.Pages.Count, nameof(selection));
        var context = new InspectionContext(document.Objects, effective);
        for (int i = 0; i < pageNumbers.Length && !context.IsStopped; i++) {
            int pageNumber = pageNumbers[i];
            PdfDictionary? resources = document.Pages[pageNumber - 1].GetFontInspectionResources();
            if (resources is not null) {
                context.InspectResources(
                    resources,
                    pageNumber,
                    "Page " + pageNumber,
                    0,
                    new HashSet<PdfStream>(ReferenceComparer<PdfStream>.Instance),
                    new Dictionary<FormResourceContext, int>(FormResourceContextComparer.Instance));
            }
        }

        return context.BuildInventory();
    }

    private sealed class InspectionContext {
        private readonly Dictionary<int, PdfIndirectObject> _objects;
        private readonly PdfFontInspectionOptions _options;
        private readonly Dictionary<PdfDictionary, FontBuilder> _fonts = new(ReferenceComparer<PdfDictionary>.Instance);
        private readonly List<FontBuilder> _fontOrder = new();
        private readonly List<PdfFontInspectionDiagnostic> _diagnostics = new();
        private readonly FontStreamDecodeBudget _decodeBudget;
        private int _referenceCount;
        private int _formTraversalCount;

        internal InspectionContext(Dictionary<int, PdfIndirectObject> objects, PdfFontInspectionOptions options) {
            _objects = objects;
            _options = options;
            _decodeBudget = new FontStreamDecodeBudget(options.MaxTotalDecodedFontBytes);
        }

        internal bool IsStopped { get; private set; }

        internal void InspectResources(
            PdfDictionary resources,
            int pageNumber,
            string resourcePath,
            int depth,
            HashSet<PdfStream> activeForms,
            Dictionary<FormResourceContext, int> visitedFormContexts) {
            if (IsStopped) return;
            InspectFonts(resources, pageNumber, resourcePath);
            if (IsStopped) return;

            if (!resources.Items.TryGetValue("XObject", out PdfObject? xObjectValue) ||
                Resolve(xObjectValue) is not PdfDictionary xObjects) return;
            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
                if (Resolve(entry.Value) is not PdfStream form ||
                    !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) continue;
                string formPath = resourcePath + "/XObject/" + entry.Key;
                if (depth >= _options.MaxResourceDepth) {
                    _diagnostics.Add(new PdfFontInspectionDiagnostic(
                        PdfFontInspectionDiagnosticCode.ResourceDepthExceeded,
                        "Font resource traversal stopped at the configured nested Form XObject depth.",
                        pageNumber,
                        formPath));
                    continue;
                }
                if (!activeForms.Add(form)) {
                    _diagnostics.Add(new PdfFontInspectionDiagnostic(
                        PdfFontInspectionDiagnosticCode.CyclicResourceGraph,
                        "Font resource traversal stopped at a cyclic Form XObject path.",
                        pageNumber,
                        formPath));
                    continue;
                }

                PdfDictionary formResources = form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceValue) && Resolve(formResourceValue) is PdfDictionary declared
                    ? declared
                    : resources;
                var context = new FormResourceContext(form, formResources);
                int contextDepth = depth + 1;
                bool previouslyVisited = visitedFormContexts.TryGetValue(context, out int previousDepth);
                if (!TryConsumeFormTraversal(pageNumber, formPath)) {
                    activeForms.Remove(form);
                    return;
                }
                if (previouslyVisited && previousDepth <= contextDepth) {
                    InspectResourceReferences(formResources, pageNumber, formPath, contextDepth, activeForms);
                    activeForms.Remove(form);
                    if (IsStopped) return;
                    continue;
                }
                visitedFormContexts[context] = contextDepth;
                InspectResources(
                    formResources,
                    pageNumber,
                    formPath,
                    contextDepth,
                    activeForms,
                    visitedFormContexts);
                activeForms.Remove(form);
                if (IsStopped) return;
            }
        }

        private void InspectResourceReferences(
            PdfDictionary resources,
            int pageNumber,
            string resourcePath,
            int depth,
            HashSet<PdfStream> activeForms) {
            if (IsStopped) return;
            InspectFonts(resources, pageNumber, resourcePath);
            if (IsStopped ||
                !resources.Items.TryGetValue("XObject", out PdfObject? xObjectValue) ||
                Resolve(xObjectValue) is not PdfDictionary xObjects) return;

            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
                if (Resolve(entry.Value) is not PdfStream form ||
                    !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) continue;
                string formPath = resourcePath + "/XObject/" + entry.Key;
                if (depth >= _options.MaxResourceDepth) {
                    _diagnostics.Add(new PdfFontInspectionDiagnostic(
                        PdfFontInspectionDiagnosticCode.ResourceDepthExceeded,
                        "Font resource traversal stopped at the configured nested Form XObject depth.",
                        pageNumber,
                        formPath));
                    continue;
                }
                if (!activeForms.Add(form)) {
                    _diagnostics.Add(new PdfFontInspectionDiagnostic(
                        PdfFontInspectionDiagnosticCode.CyclicResourceGraph,
                        "Font resource traversal stopped at a cyclic Form XObject path.",
                        pageNumber,
                        formPath));
                    continue;
                }

                PdfDictionary formResources = form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceValue) && Resolve(formResourceValue) is PdfDictionary declared
                    ? declared
                    : resources;
                if (!TryConsumeFormTraversal(pageNumber, formPath)) {
                    activeForms.Remove(form);
                    return;
                }
                InspectResourceReferences(formResources, pageNumber, formPath, depth + 1, activeForms);
                activeForms.Remove(form);
                if (IsStopped) return;
            }
        }

        private bool TryConsumeFormTraversal(int pageNumber, string resourcePath) {
            if (_formTraversalCount >= _options.MaxFormResourceTraversals) {
                AddLimitDiagnostic(
                    PdfFontInspectionDiagnosticCode.FormResourceTraversalLimitExceeded,
                    "Font inspection stopped at the configured Form XObject resource-context traversal limit.",
                    pageNumber,
                    resourcePath);
                return false;
            }
            _formTraversalCount++;
            return true;
        }

        private void InspectFonts(PdfDictionary resources, int pageNumber, string resourcePath) {
            if (!resources.Items.TryGetValue("Font", out PdfObject? fontValue) ||
                Resolve(fontValue) is not PdfDictionary fontDictionary) return;
            foreach (KeyValuePair<string, PdfObject> entry in fontDictionary.Items) {
                if (Resolve(entry.Value) is not PdfDictionary font) continue;
                if (_referenceCount >= _options.MaxResourceReferences) {
                    AddLimitDiagnostic(
                        PdfFontInspectionDiagnosticCode.ResourceReferenceLimitExceeded,
                        "Font inspection stopped at the configured resource-reference limit.",
                        pageNumber,
                        resourcePath + "/Font/" + entry.Key);
                    return;
                }
                _referenceCount++;

                if (!_fonts.TryGetValue(font, out FontBuilder? builder)) {
                    if (_fonts.Count >= _options.MaxFonts) {
                        AddLimitDiagnostic(
                            PdfFontInspectionDiagnosticCode.FontLimitExceeded,
                            "Font inspection stopped at the configured unique-font limit.",
                            pageNumber,
                            resourcePath + "/Font/" + entry.Key);
                        return;
                    }
                    GetReferenceIdentity(entry.Value, out int? objectNumber, out int? generation);
                    builder = FontBuilder.Create(font, entry.Key, objectNumber, generation, _objects, _options, _decodeBudget);
                    _fonts.Add(font, builder);
                    _fontOrder.Add(builder);
                }
                builder.AddReference(pageNumber, entry.Key, resourcePath + "/Font/" + entry.Key);
            }
        }

        private void AddLimitDiagnostic(
            PdfFontInspectionDiagnosticCode code,
            string message,
            int pageNumber,
            string resourcePath) {
            _diagnostics.Add(new PdfFontInspectionDiagnostic(code, message, pageNumber, resourcePath));
            IsStopped = true;
        }

        private PdfObject? Resolve(PdfObject? value) =>
            PdfObjectLookup.ResolveChain(_objects, value);

        private void GetReferenceIdentity(PdfObject value, out int? objectNumber, out int? generation) {
            objectNumber = null;
            generation = null;
            var visited = new HashSet<(int ObjectNumber, int Generation)>();
            PdfObject? current = value;
            while (current is PdfReference reference && visited.Add((reference.ObjectNumber, reference.Generation))) {
                objectNumber ??= reference.ObjectNumber;
                generation ??= reference.Generation;
                current = PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject indirect)
                    ? indirect.Value
                    : null;
            }
        }

        internal PdfFontInventory BuildInventory() {
            PdfFontInfo[] fonts = _fontOrder.Select(static builder => builder.Build()).ToArray();
            return new PdfFontInventory(Array.AsReadOnly(fonts), _diagnostics.AsReadOnly());
        }
    }

    private sealed class FontBuilder {
        private readonly int? _objectNumber;
        private readonly int? _generation;
        private readonly PdfFontResource _resource;
        private readonly string _familyName;
        private readonly string? _subsetTag;
        private readonly bool _isEmbedded;
        private readonly string? _embeddedProgramSubtype;
        private readonly int? _embeddedProgramEncodedLength;
        private readonly PdfOpenTypeFontInfo? _embeddedOpenTypeInfo;
        private readonly byte[]? _embeddedProgramBytes;
        private readonly List<PdfFontResourceReference> _references = new();
        private readonly List<PdfFontInspectionDiagnostic> _diagnostics = new();

        private FontBuilder(
            int? objectNumber,
            int? generation,
            PdfFontResource resource,
            string familyName,
            string? subsetTag,
            bool isEmbedded,
            string? embeddedProgramSubtype,
            int? embeddedProgramEncodedLength,
            PdfOpenTypeFontInfo? embeddedOpenTypeInfo,
            byte[]? embeddedProgramBytes) {
            _objectNumber = objectNumber;
            _generation = generation;
            _resource = resource;
            _familyName = familyName;
            _subsetTag = subsetTag;
            _isEmbedded = isEmbedded;
            _embeddedProgramSubtype = embeddedProgramSubtype;
            _embeddedProgramEncodedLength = embeddedProgramEncodedLength;
            _embeddedOpenTypeInfo = embeddedOpenTypeInfo;
            _embeddedProgramBytes = embeddedProgramBytes;
        }

        internal static FontBuilder Create(
            PdfDictionary font,
            string resourceName,
            int? objectNumber,
            int? generation,
            Dictionary<int, PdfIndirectObject> objects,
            PdfFontInspectionOptions options,
            FontStreamDecodeBudget decodeBudget) {
            FontStreamDecodeFailure toUnicodeDecodeFailure = FontStreamDecodeFailure.None;
            PdfFontResource resource = ResourceResolver.CreateFontResource(
                resourceName,
                font,
                objects,
                stream => {
                    bool decoded = decodeBudget.TryDecode(
                        stream,
                        objects,
                        options.MaxToUnicodeBytes,
                        out byte[]? bytes,
                        out toUnicodeDecodeFailure);
                    return decoded ? bytes : null;
                },
                includeEmbeddedTrueTypeFont: false);
            SplitSubsetName(resource.BaseFont, out string familyName, out string? subsetTag);
            FindEmbeddedProgram(font, objects, out PdfStream? program, out string? programSubtype);
            byte[]? decodedProgram = null;
            byte[]? programBytes = null;
            PdfOpenTypeFontInfo? openTypeInfo = null;
            bool programUnavailable = false;
            FontStreamDecodeFailure programDecodeFailure = FontStreamDecodeFailure.None;
            bool unreadableOpenTypeProgram = false;
            if (program is not null && (options.IncludeEmbeddedProgramBytes || options.InspectEmbeddedProgramMetadata)) {
                if (!decodeBudget.TryDecode(
                        program,
                        objects,
                        options.MaxEmbeddedProgramBytes,
                        out decodedProgram,
                        out programDecodeFailure)) {
                    decodedProgram = null;
                    programUnavailable = true;
                }
            }
            if (decodedProgram is not null && options.InspectEmbeddedProgramMetadata && IsOpenTypeProgram(programSubtype)) {
                if (!PdfOpenTypeFontInspector.TryInspect(decodedProgram, out openTypeInfo, out _, familyName)) {
                    unreadableOpenTypeProgram = true;
                }
            }
            if (options.IncludeEmbeddedProgramBytes) {
                programBytes = decodedProgram;
            }

            var builder = new FontBuilder(
                objectNumber,
                generation,
                resource,
                familyName,
                subsetTag,
                program is not null,
                programSubtype,
                program?.Data.Length,
                openTypeInfo,
                programBytes);
            if (string.IsNullOrWhiteSpace(resource.BaseFont)) {
                builder._diagnostics.Add(new PdfFontInspectionDiagnostic(
                    PdfFontInspectionDiagnosticCode.MissingBaseFont,
                    "Font dictionary does not declare a BaseFont name."));
            }
            if (!resource.HasToUnicode) {
                builder._diagnostics.Add(new PdfFontInspectionDiagnostic(
                    PdfFontInspectionDiagnosticCode.MissingToUnicode,
                    "Font dictionary does not declare a ToUnicode mapping."));
            } else if (resource.CMap is null) {
                builder._diagnostics.Add(new PdfFontInspectionDiagnostic(
                    toUnicodeDecodeFailure == FontStreamDecodeFailure.AggregateLimit
                        ? PdfFontInspectionDiagnosticCode.ToUnicodeTotalLimitExceeded
                        : toUnicodeDecodeFailure == FontStreamDecodeFailure.PerStreamLimit
                            ? PdfFontInspectionDiagnosticCode.ToUnicodeLimitExceeded
                            : PdfFontInspectionDiagnosticCode.UnreadableToUnicode,
                    toUnicodeDecodeFailure == FontStreamDecodeFailure.AggregateLimit
                        ? "ToUnicode mapping was not decoded because the aggregate font-stream byte allowance was exhausted."
                        : toUnicodeDecodeFailure == FontStreamDecodeFailure.PerStreamLimit
                            ? "ToUnicode mapping exceeded the configured per-map decoded-byte limit."
                            : "Font dictionary declares a ToUnicode mapping that could not be decoded."));
            }
            if (programUnavailable) {
                builder._diagnostics.Add(new PdfFontInspectionDiagnostic(
                    programDecodeFailure == FontStreamDecodeFailure.AggregateLimit
                        ? PdfFontInspectionDiagnosticCode.EmbeddedProgramTotalLimitExceeded
                        : PdfFontInspectionDiagnosticCode.EmbeddedProgramUnavailable,
                    programDecodeFailure == FontStreamDecodeFailure.AggregateLimit
                        ? "Embedded font program was not decoded because the aggregate font-stream byte allowance was exhausted."
                        : "Embedded font program could not be decoded within the configured byte limit."));
            }
            if (unreadableOpenTypeProgram) {
                builder._diagnostics.Add(new PdfFontInspectionDiagnostic(
                    PdfFontInspectionDiagnosticCode.UnreadableEmbeddedOpenTypeProgram,
                    "Embedded OpenType or TrueType font program was decoded but its table directory could not be inspected."));
            }
            return builder;
        }

        internal void AddReference(int pageNumber, string resourceName, string resourcePath) {
            _references.Add(new PdfFontResourceReference(pageNumber, resourceName, resourcePath));
        }

        internal PdfFontInfo Build() => new PdfFontInfo(
            _objectNumber,
            _generation,
            _resource.BaseFont,
            _familyName,
            _subsetTag,
            _resource.FontSubtype,
            _resource.Encoding,
            _resource.HasToUnicode,
            _resource.CMap is not null,
            _resource.CMap?.MappingCount ?? 0,
            _resource.Differences?.Count ?? 0,
            _isEmbedded,
            _embeddedProgramSubtype,
            _embeddedProgramEncodedLength,
            _embeddedOpenTypeInfo,
            _embeddedProgramBytes,
            _resource.Type3 is not null,
            _references.AsReadOnly(),
            _diagnostics.AsReadOnly());

        private static bool IsOpenTypeProgram(string? subtype) =>
            string.Equals(subtype, "TrueType", StringComparison.Ordinal) ||
            string.Equals(subtype, "OpenType", StringComparison.Ordinal);

        private static void SplitSubsetName(string baseFontName, out string familyName, out string? subsetTag) {
            familyName = baseFontName;
            subsetTag = null;
            if (baseFontName.Length <= 7 || baseFontName[6] != '+') return;
            for (int i = 0; i < 6; i++) {
                char value = baseFontName[i];
                if (value < 'A' || value > 'Z') return;
            }
            subsetTag = baseFontName.Substring(0, 6);
            familyName = baseFontName.Substring(7);
        }

        private static void FindEmbeddedProgram(
            PdfDictionary font,
            Dictionary<int, PdfIndirectObject> objects,
            out PdfStream? program,
            out string? subtype) {
            program = null;
            subtype = null;
            PdfDictionary descriptorOwner = font;
            if (string.Equals(font.Get<PdfName>("Subtype")?.Name, "Type0", StringComparison.Ordinal) &&
                font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsValue) &&
                PdfObjectLookup.ResolveChain(objects, descendantsValue) is PdfArray descendants &&
                descendants.Items.Count > 0 &&
                PdfObjectLookup.ResolveChain(objects, descendants.Items[0]) is PdfDictionary descendant) {
                descriptorOwner = descendant;
            }
            if (!descriptorOwner.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorValue) ||
                PdfObjectLookup.ResolveChain(objects, descriptorValue) is not PdfDictionary descriptor) return;

            if (TryGetProgram(descriptor, "FontFile", "Type1", objects, out program, out subtype)) return;
            if (TryGetProgram(descriptor, "FontFile2", "TrueType", objects, out program, out subtype)) return;
            if (!TryGetProgram(descriptor, "FontFile3", null, objects, out program, out subtype) || program is null) return;
            subtype = program.Dictionary.Get<PdfName>("Subtype")?.Name ?? "FontFile3";
        }

        private static bool TryGetProgram(
            PdfDictionary descriptor,
            string key,
            string? declaredSubtype,
            Dictionary<int, PdfIndirectObject> objects,
            out PdfStream? program,
            out string? subtype) {
            program = descriptor.Items.TryGetValue(key, out PdfObject? value)
                ? PdfObjectLookup.ResolveChain(objects, value) as PdfStream
                : null;
            subtype = program is null ? null : declaredSubtype;
            return program is not null;
        }
    }

    private sealed class FontStreamDecodeBudget {
        private long _remainingBytes;

        internal FontStreamDecodeBudget(long maximumBytes) {
            _remainingBytes = maximumBytes;
        }

        internal bool TryDecode(
            PdfStream program,
            Dictionary<int, PdfIndirectObject> objects,
            int maximumStreamBytes,
            out byte[]? decoded,
            out FontStreamDecodeFailure failure) {
            decoded = null;
            failure = FontStreamDecodeFailure.None;
            if (_remainingBytes <= 0L) {
                failure = FontStreamDecodeFailure.AggregateLimit;
                return false;
            }

            long remainingBeforeDecode = _remainingBytes;
            int maximumOutput = (int)Math.Min(maximumStreamBytes, Math.Min(remainingBeforeDecode, int.MaxValue));
            if (!StreamDecoder.TryDecode(
                    program.Dictionary,
                    program.Data,
                    maximumOutput,
                    out byte[] candidate,
                    out bool decodedLimitExceeded,
                    objects)) {
                _remainingBytes -= maximumOutput;
                failure = decodedLimitExceeded
                    ? remainingBeforeDecode <= maximumStreamBytes
                        ? FontStreamDecodeFailure.AggregateLimit
                        : FontStreamDecodeFailure.PerStreamLimit
                    : FontStreamDecodeFailure.DecodeFailure;
                return false;
            }

            _remainingBytes -= candidate.LongLength;
            decoded = candidate;
            return true;
        }
    }

    private enum FontStreamDecodeFailure {
        None,
        DecodeFailure,
        PerStreamLimit,
        AggregateLimit
    }

    private readonly struct FormResourceContext {
        internal FormResourceContext(PdfStream form, PdfDictionary resources) {
            Form = form;
            Resources = resources;
        }

        internal PdfStream Form { get; }
        internal PdfDictionary Resources { get; }
    }

    private sealed class FormResourceContextComparer : IEqualityComparer<FormResourceContext> {
        internal static FormResourceContextComparer Instance { get; } = new FormResourceContextComparer();

        public bool Equals(FormResourceContext x, FormResourceContext y) =>
            ReferenceEquals(x.Form, y.Form) && ReferenceEquals(x.Resources, y.Resources);

        public int GetHashCode(FormResourceContext value) =>
            unchecked((RuntimeHelpers.GetHashCode(value.Form) * 397) ^ RuntimeHelpers.GetHashCode(value.Resources));
    }

    private sealed class ReferenceComparer<T> : IEqualityComparer<T> where T : class {
        internal static ReferenceComparer<T> Instance { get; } = new ReferenceComparer<T>();
        public bool Equals(T? x, T? y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
