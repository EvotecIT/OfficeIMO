using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfRenderingIntentResolver {
    internal static readonly OfficeIccRenderingIntent[] All = {
        OfficeIccRenderingIntent.Perceptual,
        OfficeIccRenderingIntent.RelativeColorimetric,
        OfficeIccRenderingIntent.Saturation,
        OfficeIccRenderingIntent.AbsoluteColorimetric
    };

    internal static OfficeIccRenderingIntent FromName(string? name) => name switch {
        "Perceptual" => OfficeIccRenderingIntent.Perceptual,
        "Saturation" => OfficeIccRenderingIntent.Saturation,
        "AbsoluteColorimetric" => OfficeIccRenderingIntent.AbsoluteColorimetric,
        _ => OfficeIccRenderingIntent.RelativeColorimetric
    };

    internal static string BuildResourceKey(string name, OfficeIccRenderingIntent renderingIntent) =>
        name + "|intent:" + ((int)renderingIntent).ToString(System.Globalization.CultureInfo.InvariantCulture);

    internal static bool TryGetResource<T>(
        IReadOnlyDictionary<string, T>? resources,
        string name,
        OfficeIccRenderingIntent renderingIntent,
        out T resource) {
        resource = default!;
        if (resources == null) return false;
        if (resources.TryGetValue(BuildResourceKey(name, renderingIntent), out resource!)) return true;
        foreach (OfficeIccRenderingIntent intent in All) {
            if (resources.ContainsKey(BuildResourceKey(name, intent))) return false;
        }
        return resources.TryGetValue(name, out resource!);
    }

    internal static OfficeIccRenderingIntent Read(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        OfficeIccRenderingIntent inherited) =>
        TryRead(dictionary, key, objects, out OfficeIccRenderingIntent renderingIntent)
            ? renderingIntent
            : inherited;

    internal static bool TryRead(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        out OfficeIccRenderingIntent renderingIntent) {
        renderingIntent = OfficeIccRenderingIntent.RelativeColorimetric;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) return false;
        PdfObject? resolved = ResolveObject(value, objects);
        if (resolved is PdfNull) return false;
        renderingIntent = resolved is PdfName name
            ? FromName(name.Name)
            : OfficeIccRenderingIntent.RelativeColorimetric;
        return true;
    }

    private static PdfObject? ResolveObject(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        while (value is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) {
                return null;
            }
            value = indirect.Value;
        }
        return value;
    }
}
