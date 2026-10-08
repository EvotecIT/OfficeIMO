using OfficeIMO.Core.Internal;

namespace OfficeIMO.Xps;

internal static class XpsPackage {
    internal const string XpsNamespace = "http://schemas.microsoft.com/xps/2005/06";
    internal const string OxpsNamespace = "http://schemas.openxps.org/oxps/v1.0";
    internal static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    internal static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    internal static string Namespace(XpsFormat format) => format == XpsFormat.Xps ? XpsNamespace : OxpsNamespace;
    internal static string StartRelationship(XpsFormat format) => Namespace(format) + "/fixedrepresentation";
    internal static string Type(string kind) => "application/vnd.ms-package.xps-" + kind + "+xml";
    internal static byte[] Serialize(XElement root) {
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false })) root.Save(writer);
        return stream.ToArray();
    }
    internal static XElement Xml(byte[] bytes, XpsReadOptions limits, CancellationToken token) {
        using var stream = new MemoryStream(bytes, false);
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = limits.MaximumPartBytes };
        // Inspect depth before constructing a tree; deeply nested input must not reach recursive consumers.
        using (var reader = XmlReader.Create(stream, settings)) {
            while (reader.Read()) {
                token.ThrowIfCancellationRequested();
                if (reader.Depth > limits.MaximumXmlDepth) throw new InvalidDataException("XPS XML nesting limit exceeded.");
            }
        }
        stream.Position = 0;
        using var safeReader = XmlReader.Create(stream, settings);
        return XElement.Load(safeReader, LoadOptions.PreserveWhitespace);
    }
    internal static string PartName(string name) {
        if (string.IsNullOrEmpty(name) || name.IndexOfAny(new[] { '\\', '?', '#', '\0', ':' }) >= 0 ||
            name.Contains("//") || name.EndsWith("/", StringComparison.Ordinal)) throw new InvalidDataException("Invalid XPS part name.");
        string value = name.TrimStart('/');
        if (name.StartsWith("//", StringComparison.Ordinal) || OfficeArchiveSafety.IsUnsafePath(value)) throw new InvalidDataException("Unsafe XPS part name.");
        // Percent-encoded separators or dot segments must not alias another package part.
        string decoded = Uri.UnescapeDataString(value);
        if (decoded.IndexOfAny(new[] { '\\', '?', '#', '\0', ':' }) >= 0 || decoded.Contains("//") || OfficeArchiveSafety.IsUnsafePath(decoded) ||
            decoded.Count(c => c == '/') != value.Count(c => c == '/')) throw new InvalidDataException("Ambiguous XPS part name.");
        return value;
    }
    internal static string Resolve(string source, string target) {
        if (string.IsNullOrWhiteSpace(target) || target.IndexOfAny(new[] { '\\', '?', '#', '\0', ':' }) >= 0 || target.StartsWith("//", StringComparison.Ordinal))
            throw new InvalidDataException("XPS resource references must be package-local part URIs.");
        var segments = new List<string>();
        if (!target.StartsWith("/", StringComparison.Ordinal)) segments.AddRange(source.Substring(0, Math.Max(0, source.LastIndexOf('/') + 1)).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries));
        foreach (string segment in target.Split('/')) {
            if (segment == "" || segment == ".") continue;
            if (segment == "..") {
                if (segments.Count == 0) throw new InvalidDataException("XPS reference escapes the package root.");
                segments.RemoveAt(segments.Count - 1);
            } else segments.Add(segment);
        }
        return PartName(string.Join("/", segments));
    }
    internal static double Number(string? value, double fallback = 0) {
        if (value == null) return fallback;
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ||
            double.IsNaN(result) || double.IsInfinity(result) || Math.Abs(result) > 10000000)
            throw new InvalidDataException("Invalid or excessive XPS numeric value.");
        return result;
    }
    internal static string N(double value) => value.ToString("R", CultureInfo.InvariantCulture);
}
