using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio;

public static partial class VisioDuplicationExtensions {
    private static void RemapCopiedFormulaReferences(VisioPage page,
        IReadOnlyDictionary<VisioShape, VisioShape> shapes,
        IReadOnlyDictionary<VisioConnector, VisioConnector> connectors,
        IReadOnlyDictionary<string, string> sourceIdentifiers, bool remapPageSheet = false) {
        VisioDocument.ValidateCopiedNativeCellMetadata(shapes.Values, connectors.Values, page.OwnerDocument, page);
        var destination = VisioDocument.AssignPageElementIdentifiers(page);
        var pairs = shapes.Select(pair => (sourceIdentifiers[pair.Key.Id], destination[pair.Value.Id]))
            .Concat(connectors.Select(pair => (sourceIdentifiers[pair.Key.Id], destination[pair.Value.Id])));
        var references = BuildFormulaReferenceMap(pairs);
        foreach (VisioShape shape in shapes.Values) VisioShapeFormulaReferences.Remap(shape, references);
        foreach (VisioConnector connector in connectors.Values) VisioShapeFormulaReferences.Remap(connector, references);
        if (remapPageSheet) VisioShapeFormulaReferences.RemapPageSheet(page, references);
    }

    private static Dictionary<string, string> BuildFormulaReferenceMap(IEnumerable<(string Source, string Target)> pairs) {
        var references = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in pairs) {
            references.Add(VisioShapeFormulaReferences.NormalizeSheetId(pair.Source), pair.Target);
        }
        return references;
    }
}
