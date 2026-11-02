namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    // A collision belongs to the highest named ancestor, so all its descendants move together.
    private static (int ObjectNumber, string Name) FindFormNamespace(Dictionary<int, PdfIndirectObject> objects, int terminal) {
        var result = (ObjectNumber: terminal, Name: string.Empty);
        foreach (int number in FormAncestorChain(objects, terminal)) {
            var dictionary = (PdfDictionary)objects[number].Value;
            if (dictionary.Items.TryGetValue("T", out var value) && ResolveObject(objects, value) is PdfStringObj name) {
                if (name.Value.Contains('.'))
                    throw new NotSupportedException("AcroForm partial field names containing periods require repair before composition.");
                if (name.Value.Length > 0) result = (number, name.Value);
            }
        }
        return result;
    }

    /// <summary>Keeps real field parents and their metadata, pruning only children outside the selection.</summary>
    private static PdfArray RetainFormHierarchy(Dictionary<int, PdfIndirectObject> objects, IEnumerable<int> terminals) {
        int[] terminalNumbers = terminals.Distinct().ToArray();
        var terminalSet = new HashSet<int>(terminalNumbers);
        var retained = new HashSet<int>();
        var roots = new List<int>();
        var rootSet = new HashSet<int>();
        foreach (int terminal in terminalNumbers) {
            int root = terminal;
            foreach (int number in FormAncestorChain(objects, terminal)) {
                retained.Add(number);
                root = number;
            }
            if (rootSet.Add(root)) roots.Add(root);
        }
        foreach (int number in retained) {
            PdfIndirectObject original = objects[number];
            var dictionary = (PdfDictionary)original.Value;
            var clone = new PdfDictionary();
            foreach (var entry in dictionary.Items) clone.Items[entry.Key] = entry.Value;
            if (!terminalSet.Contains(number) && dictionary.Items.TryGetValue("Kids", out var value) && ResolveObject(objects, value) is PdfArray kids) {
                var kept = new PdfArray();
                foreach (PdfReference child in kids.Items.OfType<PdfReference>())
                    if (retained.Contains(child.ObjectNumber)) kept.Items.Add(child);
                clone.Items["Kids"] = kept;
            }
            objects[number] = new PdfIndirectObject(number, original.Generation, clone);
        }
        var fields = new PdfArray();
        foreach (int root in roots) fields.Items.Add(new PdfReference(root, objects[root].Generation));
        return fields;
    }

    // Names come from /Kids, so every /Parent walk must see that same hierarchy.
    private static void AlignFormParents(Dictionary<int, PdfIndirectObject> objects, PdfReadDocument document) {
        if (!document.Security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(document.Security.RootObjectNumber.Value, out PdfIndirectObject? catalog) ||
            catalog.Value is not PdfDictionary catalogDictionary ||
            !catalogDictionary.Items.TryGetValue("AcroForm", out var formObject) ||
            ResolveDictionary(objects, formObject) is not PdfDictionary form ||
            !form.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            return;
        }
        if (fields.Items.Any(static item => item is not PdfReference))
            throw new NotSupportedException("Form composition requires indirect AcroForm field roots.");
        var visited = new HashSet<int>();
        foreach (PdfReference root in fields.Items.OfType<PdfReference>()) AlignFormParents(objects, root.ObjectNumber, visited);
    }

    private static void AlignFormParents(Dictionary<int, PdfIndirectObject> objects, int field, HashSet<int> visited) {
        if (!visited.Add(field) || !objects.TryGetValue(field, out var indirect) || indirect.Value is not PdfDictionary dictionary ||
            !dictionary.Items.TryGetValue("Kids", out var value) || ResolveObject(objects, value) is not PdfArray kids) {
            return;
        }
        foreach (PdfReference kid in kids.Items.OfType<PdfReference>()) {
            AlignFormParent(objects, kid.ObjectNumber, field);
            AlignFormParents(objects, kid.ObjectNumber, visited);
        }
    }

    private static void AlignFormParent(Dictionary<int, PdfIndirectObject> objects, int child, int parent) {
        if (!objects.TryGetValue(child, out var indirect) || indirect.Value is not PdfDictionary field) return;
        if (field.Items.TryGetValue("Parent", out var current) && current is PdfReference reference && reference.ObjectNumber == parent) return;
        var aligned = new PdfDictionary();
        foreach (var entry in field.Items) aligned.Items[entry.Key] = entry.Value;
        aligned.Items["Parent"] = new PdfReference(parent, objects[parent].Generation);
        objects[child] = new PdfIndirectObject(child, indirect.Generation, aligned);
    }

    private static IEnumerable<int> FormAncestorChain(Dictionary<int, PdfIndirectObject> objects, int terminal) {
        var visited = new HashSet<int>();
        int number = terminal;
        while (true) {
            if (!visited.Add(number) || !objects.TryGetValue(number, out var indirect) || indirect.Value is not PdfDictionary dictionary)
                throw new NotSupportedException("AcroForm composition requires an acyclic indirect field hierarchy.");
            yield return number;
            if (!dictionary.Items.TryGetValue("Parent", out var parent)) yield break;
            if (parent is not PdfReference reference)
                throw new NotSupportedException("AcroForm composition requires indirect field parents.");
            number = reference.ObjectNumber;
        }
    }
}
