using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Visio;

/// <summary>Resolves serialized hyperlink identities without changing the source model.</summary>
internal static class VisioHyperlinkRowNames {
    /// <summary>Gets the corresponding inherited shape's rows, including nested master-shape references.</summary>
    internal static IList<VisioHyperlink>? Inherited(VisioShape shape, VisioDocument? document = null) =>
        FromMaster(shape, document?.ResolveEffectiveMaster(shape));

    /// <summary>Uses the writer's resolved master while retaining the corresponding nested master shape.</summary>
    internal static IList<VisioHyperlink>? FromMaster(VisioShape shape, VisioMaster? master) =>
        (shape.MasterShape ?? master?.Shape ?? shape.Master?.Shape)?.Hyperlinks;

    internal static IList<VisioHyperlink>? Inherited(VisioConnector connector, VisioDocument? document) =>
        document?.ResolveEffectiveMaster(connector)?.Shape.Hyperlinks;

    /// <summary>
    /// Keeps explicit names and indexed unnamed source rows; generated names avoid all local
    /// and inherited names. Explicit local names can intentionally override inherited rows.
    /// </summary>
    internal static string?[] Create(IList<VisioHyperlink> hyperlinks, IList<VisioHyperlink>? inherited = null) {
        var names = new string?[hyperlinks.Count];
        var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (inherited != null) {
            foreach (string? name in Create(inherited))
                if (name != null) reserved.Add(name);
        }
        foreach (VisioHyperlink hyperlink in hyperlinks)
            if (!string.IsNullOrWhiteSpace(hyperlink.RowName)) reserved.Add(hyperlink.RowName!);

        for (int i = 0; i < hyperlinks.Count; i++) {
            VisioHyperlink hyperlink = hyperlinks[i];
            if (!string.IsNullOrWhiteSpace(hyperlink.RowName)) {
                names[i] = hyperlink.RowName;
            } else if (!hyperlink.RowIndex.HasValue) {
                int number = i + 1;
                string name;
                do {
                    name = "Row_" + number.ToString(CultureInfo.InvariantCulture);
                    number++;
                } while (!reserved.Add(name));
                names[i] = name;
            }
        }
        return names;
    }
}
