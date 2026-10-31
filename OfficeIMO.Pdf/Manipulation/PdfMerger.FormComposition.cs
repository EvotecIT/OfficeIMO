using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    private static PdfMergeOptions PageImportMergeOptions() => new PdfMergeOptions {
        Policy = new PdfMergePolicy { Forms = PdfMergeStructureMode.Combine }
    };

    /// <summary>Reuses merge field/resource policy inside a catalog-preserving page insertion.</summary>
    internal static void CombineImportedForms(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        PdfReadDocument target, PdfReadDocument incoming, IReadOnlyDictionary<int, int> incomingMap, CancellationToken token) {
        if (incoming.FormFields.Count == 0) return;
        if (PdfTaggedPageOrganizer.IsTagged(target) || PdfTaggedPageOrganizer.IsTagged(incoming))
            throw new NotSupportedException("Importing AcroForm fields with tagged pages requires preservation of field structure ownership; this combination is not supported.");
        var sources = new[] {
            FormCompositionSource.FromDocument(target, target.Objects.Keys.ToDictionary(number => number, number => number)),
            FormCompositionSource.FromDocument(incoming, incomingMap)
        };
        int dropped = 0;
        ComposeFormGraph(objects, catalog, target, sources, 0, PdfMergeStructureMode.Combine,
            PdfMergeCollisionMode.RenameIncoming, new List<string>(), new List<string>(), ref dropped, token);
    }

    private sealed class FormCompositionSource {
        internal FormCompositionSource(PdfReadDocument document, Dictionary<int, PdfIndirectObject> objects,
            int[] roots, IReadOnlyDictionary<int, int> numberMap) {
            Document = document; Objects = objects; FormFieldRootObjectNumbers = roots; NumberMap = numberMap;
        }
        internal PdfReadDocument Document { get; }
        internal Dictionary<int, PdfIndirectObject> Objects { get; }
        internal int[] FormFieldRootObjectNumbers { get; }
        internal IReadOnlyDictionary<int, int> NumberMap { get; }
        internal PdfDictionary? Form => Objects.TryGetValue(Document.Security.RootObjectNumber!.Value, out var root) &&
            root.Value is PdfDictionary catalog && catalog.Items.TryGetValue("AcroForm", out var value)
                ? ResolveDictionary(Objects, value) : null;

        internal static FormCompositionSource FromDocument(PdfReadDocument document, IReadOnlyDictionary<int, int> map) {
            var result = new FormCompositionSource(document, document.Objects, Array.Empty<int>(), map);
            var roots = new List<int>();
            if (result.Form?.Items.TryGetValue("Fields", out var value) == true && ResolveObject(document.Objects, value) is PdfArray fields) {
                foreach (PdfObject field in fields.Items) {
                    if (field is not PdfReference reference)
                        throw new NotSupportedException("Page import requires indirect AcroForm field roots.");
                    roots.Add(reference.ObjectNumber);
                }
            }
            if (document.FormFields.Count > 0 && roots.Count == 0)
                throw new NotSupportedException("Page import requires indirect AcroForm field roots.");
            return new FormCompositionSource(document, document.Objects, roots.ToArray(), map);
        }
    }

    private static PdfDictionary BuildMergedAcroForm(Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<MergedFormRoot> roots, IReadOnlyList<FormCompositionSource> sources, int primarySourceIndex, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var form = new PdfDictionary();
        var resources = new PdfDictionary();
        var calculations = new PdfArray();
        bool needAppearances = false;
        var selected = new HashSet<int>(roots.Select(root => root.RootObjectNumber));
        foreach (int sourceIndex in Enumerable.Range(0, sources.Count).OrderBy(index => index == primarySourceIndex ? 0 : 1)) {
            token.ThrowIfCancellationRequested();
            FormCompositionSource source = sources[sourceIndex];
            MergedFormRoot[] sourceRoots = roots.Where(root => root.SourceIndex == sourceIndex).ToArray();
            if (sourceRoots.Length == 0 || source.Form is not PdfDictionary original) continue;
            needAppearances |= source.Document.AcroFormNeedAppearances == true;
            var fontNames = MergeFormResources(objects, resources, sourceRoots[0].SourceDefaultResources, sourceIndex, token);
            foreach (MergedFormRoot root in sourceRoots) {
                token.ThrowIfCancellationRequested();
                foreach (int widget in root.WidgetObjectNumbers) {
                    if (objects[widget].Value is not PdfDictionary dictionary ||
                        !dictionary.Items.TryGetValue("AP", out var appearanceValue) ||
                        ResolveDictionary(objects, appearanceValue) is not PdfDictionary appearances) continue;
                    foreach (var appearance in appearances.Items.Values) MaterializeAppearanceResources(appearance);
                }
                if (root.SourceDefaultResources != null) root.FieldDictionary.Items["DR"] = resources;
                RemapFormDefaultAppearance(objects, root.FieldDictionary, fontNames, token);
                foreach (int widget in root.WidgetObjectNumbers)
                    if (widget != root.RootObjectNumber && objects[widget].Value is PdfDictionary dictionary)
                        RemapFormDefaultAppearance(objects, dictionary, fontNames, token);
            }
            if (!form.Items.ContainsKey("DA") && !string.IsNullOrEmpty(source.Document.AcroFormDefaultAppearance)) {
                form.Items["DA"] = new PdfStringObj(source.Document.AcroFormDefaultAppearance!, true);
                RemapFormDefaultAppearance(objects, form, fontNames, token);
            }
            if (!form.Items.ContainsKey("Q") && source.Document.AcroFormQuadding.HasValue)
                form.Items["Q"] = new PdfNumber(source.Document.AcroFormQuadding.Value);
            void MaterializeAppearanceResources(PdfObject value) {
                PdfObject? resolved = ResolveObject(objects, value);
                if (resolved is PdfStream stream) {
                    if (!stream.Dictionary.Items.ContainsKey("Resources") && sourceRoots[0].SourceDefaultResources != null)
                        stream.Dictionary.Items["Resources"] = sourceRoots[0].SourceDefaultResources!;
                } else if (resolved is PdfDictionary states)
                    foreach (var state in states.Items.Values)
                        if (ResolveObject(objects, state) is PdfStream stateStream && !stateStream.Dictionary.Items.ContainsKey("Resources") && sourceRoots[0].SourceDefaultResources != null)
                            stateStream.Dictionary.Items["Resources"] = sourceRoots[0].SourceDefaultResources!;
            }
            if (original.Items.TryGetValue("CO", out var order) && ResolveObject(source.Objects, order) is PdfArray co) {
                foreach (PdfReference reference in co.Items.OfType<PdfReference>())
                    if (source.NumberMap.TryGetValue(reference.ObjectNumber, out int mapped) && selected.Contains(mapped))
                        calculations.Items.Add(new PdfReference(mapped, objects[mapped].Generation));
            }
        }
        form.Items["Fields"] = RetainFormHierarchy(objects, roots.Select(root => root.RootObjectNumber));
        form.Items["NeedAppearances"] = new PdfBoolean(needAppearances);
        if (resources.Items.Count > 0) form.Items["DR"] = resources;
        if (calculations.Items.Count > 0) form.Items["CO"] = calculations;
        return form;
    }

    private static Dictionary<string, string> MergeFormResources(Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources, PdfObject? sourceResources, int sourceIndex, CancellationToken token) {
        var fontNames = new Dictionary<string, string>(StringComparer.Ordinal);
        if (sourceResources is null || ResolveDictionary(objects, sourceResources) is not PdfDictionary incoming) return fontNames;
        foreach (var category in incoming.Items) {
            token.ThrowIfCancellationRequested();
            if (ResolveDictionary(objects, category.Value) is not PdfDictionary entries) {
                if (!resources.Items.ContainsKey(category.Key)) resources.Items[category.Key] = category.Value;
                continue;
            }
            if (!resources.Items.TryGetValue(category.Key, out var existing) || existing is not PdfDictionary destination)
                resources.Items[category.Key] = destination = new PdfDictionary();
            foreach (var entry in entries.Items) {
                token.ThrowIfCancellationRequested();
                string name = entry.Key;
                if (destination.Items.ContainsKey(name)) {
                    string prefix = name + "_source" + (sourceIndex + 1).ToString(CultureInfo.InvariantCulture);
                    name = prefix; int sequence = 2;
                    while (destination.Items.ContainsKey(name)) name = prefix + "_" + (sequence++).ToString(CultureInfo.InvariantCulture);
                }
                destination.Items[name] = entry.Value;
                fontNames[category.Key + "/" + entry.Key] = name;
            }
        }
        return fontNames;
    }

    // DA is a content stream. Use the shared lexer so escaped resource names and comments keep their PDF meaning.
    private static void RemapFormDefaultAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field,
        Dictionary<string, string> fontNames, CancellationToken token) {
        if (!fontNames.Any(pair => pair.Key.Substring(pair.Key.IndexOf('/') + 1) != pair.Value) || !field.Items.TryGetValue("DA", out var value) ||
            ResolveObject(objects, value) is not PdfStringObj text) return;
        var result = new StringBuilder();
        PdfContentStreamInterpreter.Interpret(text.Value, int.MaxValue, operation => {
            token.ThrowIfCancellationRequested();
            if (operation.HasInvalidOperands) throw new NotSupportedException("AcroForm default appearance has invalid operands.");
            for (int index = 0; index < operation.Operands.Count; index++) {
                object operand = operation.Operands[index];
                if (operand is string name) {
                    string? category = operation.Name switch {
                        "Tf" => "Font", "gs" => "ExtGState", "cs" or "CS" => "ColorSpace",
                        "scn" or "SCN" => "Pattern", "sh" => "Shading", _ => null
                    };
                    if (category != null && fontNames.TryGetValue(category + "/" + name, out string? mapped)) name = mapped;
                    result.Append('/').Append(PdfSyntaxEscaper.Name(name));
                } else if (operand is double number) result.Append(PdfSyntaxEscaper.Number(number));
                else throw new NotSupportedException("Combining AcroForm default appearances with non-name/non-numeric operands is not supported.");
                result.Append(' ');
            }
            result.Append(operation.Name).Append('\n');
        }, dispatchInvalidOperations: true);
        field.Items["DA"] = new PdfStringObj(result.ToString(), true);
    }
}
