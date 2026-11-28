using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

// The built-in table style definitions, read from WordTableStyles.xml.
internal static partial class WordTableStyles {
    private const string MarkupResource = "OfficeIMO.Word.TableStyles.WordTableStyles.xml";
    private static readonly Lazy<Dictionary<string, string>> MarkupById = new(LoadMarkup);

    /// <summary>
    /// The definition of a built-in table style as a <c>w:style</c> element with its namespace declared on it, or null
    /// for <see cref="WordTableStyle.OfficeIMOModern"/>, which is built from the document theme.
    /// </summary>
    internal static string? GetStyleMarkup(WordTableStyle style) =>
        style == WordTableStyle.OfficeIMOModern ? null : MarkupById.Value[GetStyleId(style)];

    private static Dictionary<string, string> LoadMarkup() {
        byte[] bytes;
        using (Stream stream = typeof(WordTableStyles).Assembly.GetManifestResourceStream(MarkupResource)
            ?? throw new InvalidOperationException("The built-in table styles are missing from the assembly.")) {
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            bytes = buffer.ToArray();
        }
        OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
        const int W = OfficeXmlNamespaces.WordprocessingML;
        var markup = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int node = tree.FirstChildOf(tree.Root); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
            if (!tree.IsElement(node, W, "style")) continue;
            markup[tree.GetAttribute(node, W, "styleId")!] = Encoding.UTF8.GetString(tree.WriteElement(node));
        }
        return markup;
    }
}
