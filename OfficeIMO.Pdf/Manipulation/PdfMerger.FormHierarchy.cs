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
