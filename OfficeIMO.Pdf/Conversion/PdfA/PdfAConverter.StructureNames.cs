using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private const int ShortenedNamePrefixBytes = 100;

    /// <summary>
    /// PDF/A-2 and PDF/A-3 limit names to 127 bytes. Producers that derive structure types from style names can exceed it.
    /// Structure type names drive no drawing, so over-long ones are renamed consistently in the role map and every structure
    /// element; the role map keeps each renamed type's standard role.
    /// </summary>
    private void ShortenStructureTypeNames(PdfDictionary catalog) {
        if (_part == 4 || ResolveDictionary(catalog.Items.TryGetValue("StructTreeRoot", out PdfObject? rootValue) ? rootValue : null) is not PdfDictionary root) return;
        PdfDictionary? roleMap = ResolveDictionary(root.Items.TryGetValue("RoleMap", out PdfObject? roleMapValue) ? roleMapValue : null);
        var elements = new List<PdfDictionary>();
        var visited = new HashSet<PdfObject>(ReferenceComparer.Instance);
        var pending = new Stack<PdfObject>();
        if (root.Items.TryGetValue("K", out PdfObject? kids)) pending.Push(kids);
        while (pending.Count > 0 && visited.Count < 1_000_000) {
            PdfObject? item = Resolve(pending.Pop());
            if (item is PdfArray array) {
                if (!visited.Add(array)) continue;
                foreach (PdfObject child in array.Items) pending.Push(child);
            } else if (item is PdfDictionary element && visited.Add(element)) {
                if (element.Items.ContainsKey("S")) elements.Add(element);
                if (element.Items.TryGetValue("K", out PdfObject? children)) pending.Push(children);
            }
        }

        var longNames = new SortedSet<string>(StringComparer.Ordinal);
        if (roleMap is not null) {
            foreach (KeyValuePair<string, PdfObject> entry in roleMap.Items) {
                if (IsOverLongName(entry.Key)) longNames.Add(entry.Key);
                if (NameOf(Resolve(entry.Value)) is string target && IsOverLongName(target)) longNames.Add(target);
            }
        }
        foreach (PdfDictionary element in elements) {
            if (GetName(element, "S") is string type && IsOverLongName(type)) longNames.Add(type);
        }
        if (longNames.Count == 0) return;
        if (!Change(PdfAConversionChangeKind.StructureUpdated, "structure tree",
                "Shortened " + longNames.Count.ToString(CultureInfo.InvariantCulture) +
                " structure type name(s) longer than the 127-byte PDF/A limit; the role map keeps their standard roles and the page draws the same.",
                root, OfficeConversionLossKind.Approximation)) return;

        var used = new HashSet<string>(roleMap?.Items.Keys ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
        var renamed = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (string name in longNames) {
            string shortName = ShortenName(name);
            while (!used.Add(shortName)) shortName += "_";
            renamed[name] = shortName;
        }
        if (roleMap is not null) {
            foreach (KeyValuePair<string, PdfObject> entry in roleMap.Items.ToList()) {
                PdfObject value = NameOf(Resolve(entry.Value)) is string target && renamed.TryGetValue(target, out string? newTarget) ? new PdfName(newTarget) : entry.Value;
                if (renamed.TryGetValue(entry.Key, out string? newKey)) {
                    roleMap.Items.Remove(entry.Key);
                    roleMap.Items[newKey] = value;
                } else {
                    roleMap.Items[entry.Key] = value;
                }
            }
        }
        foreach (PdfDictionary element in elements) {
            if (GetName(element, "S") is string type && renamed.TryGetValue(type, out string? newType)) element.Items["S"] = new PdfName(newType);
        }
    }

    private static bool IsOverLongName(string name) => Encoding.UTF8.GetByteCount(name) > MaximumNameBytes;

    /// <summary>Keeps the readable start of the name and appends a hash of the whole name, so distinct names stay distinct.</summary>
    private static string ShortenName(string name) {
        var prefix = new StringBuilder();
        int bytes = 0;
        foreach (char character in name) {
            int size = Encoding.UTF8.GetByteCount(new[] { character });
            if (char.IsSurrogate(character) || bytes + size > ShortenedNamePrefixBytes) break;
            prefix.Append(character);
            bytes += size;
        }
        uint hash = 2166136261;
        foreach (char character in name) hash = unchecked((hash ^ character) * 16777619);
        return prefix.ToString().TrimEnd() + "-" + hash.ToString("X8", CultureInfo.InvariantCulture);
    }
}
