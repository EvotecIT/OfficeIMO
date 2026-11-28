using System.Collections.Concurrent;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Adds the parts of a built-in SmartArt diagram. Each layout ships as the part markup it writes, in
    /// <c>SmartArt/Templates</c>; <c>$IDn$</c> marks a model identifier that is new for every diagram and
    /// <c>$PERSIST$</c> the relationship to the persisted layout drawing.
    /// </summary>
    internal static class SmartArtBuiltIn {
        private static readonly ConcurrentDictionary<string, string> Templates = new(StringComparer.Ordinal);

        internal static (string relLayout, string relColors, string relStyle, string relData) AddParts(object owner, WordSmartArtType type) {
            string name = type switch {
                WordSmartArtType.Cycle => "Cycle",
                WordSmartArtType.CustomSmartArt1 => "Custom1",
                WordSmartArtType.CustomSmartArt2 => "Custom2",
                WordSmartArtType.Hierarchy => "Hierarchy",
                WordSmartArtType.PictureOrgChart => "PictureOrgChart",
                WordSmartArtType.ContinuousBlockProcess => "ContinuousBlockProcess",
                _ => "BasicProcess"
            };
            bool custom = type is WordSmartArtType.CustomSmartArt1 or WordSmartArtType.CustomSmartArt2;
            string styleSet = custom ? name : "Common";
            var ids = new Dictionary<string, string>(StringComparer.Ordinal);

            string relLayout = WordPartOwnership.AddDiagramPart(owner, WordDiagramPartKind.Layout, Fill(name + ".layout.xml", ids, null));
            string relColors = WordPartOwnership.AddDiagramPart(owner, WordDiagramPartKind.Colors, Fill(styleSet + ".colors.xml", ids, null));
            string relStyle = WordPartOwnership.AddDiagramPart(owner, WordDiagramPartKind.Style, Fill(styleSet + ".quickStyle.xml", ids, null));
            string? persist = custom
                ? WordPartOwnership.AddDiagramPart(owner, WordDiagramPartKind.PersistLayout, Fill(name + ".drawing.xml", ids, null))
                : null;
            string relData = WordPartOwnership.AddDiagramPart(owner, WordDiagramPartKind.Data, Fill(name + ".data.xml", ids, persist));
            return (relLayout, relColors, relStyle, relData);
        }

        private static byte[] Fill(string template, Dictionary<string, string> ids, string? persistRelationshipId) {
            string markup = Templates.GetOrAdd(template, Load);
            if (markup.IndexOf('$') < 0) return Encoding.UTF8.GetBytes(markup);
            var builder = new StringBuilder(markup.Length + 64);
            int position = 0;
            while (position < markup.Length) {
                int start = markup.IndexOf('$', position);
                int end = start < 0 ? -1 : markup.IndexOf('$', start + 1);
                if (end < 0) {
                    builder.Append(markup, position, markup.Length - position);
                    break;
                }
                builder.Append(markup, position, start - position);
                string token = markup.Substring(start + 1, end - start - 1);
                if (token == "PERSIST") {
                    builder.Append(persistRelationshipId ?? throw new InvalidOperationException("The SmartArt template needs a persisted layout part."));
                } else {
                    if (!ids.TryGetValue(token, out string? id)) {
                        id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
                        ids.Add(token, id);
                    }
                    builder.Append(id);
                }
                position = end + 1;
            }
            return Encoding.UTF8.GetBytes(builder.ToString());
        }

        private static string Load(string template) {
            string resource = "OfficeIMO.Word.SmartArt.Templates." + template;
            using Stream stream = typeof(SmartArtBuiltIn).Assembly.GetManifestResourceStream(resource)
                ?? throw new InvalidOperationException("The SmartArt template " + resource + " is missing.");
            using var reader = new StreamReader(stream, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: false);
            return reader.ReadToEnd();
        }
    }
}
