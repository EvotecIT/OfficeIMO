using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

/// <summary>
/// The built-in list definitions, kept as WordprocessingML in <c>ListStyles/WordListStyles.xml</c>.
/// </summary>
internal static class WordListStyles {
    private const string MarkupResource = "OfficeIMO.Word.ListStyles.WordListStyles.xml";
    private const string CustomTemplateCode = "FFFFFFFE";
    private const string PictureBulletTemplateCode = "FFFFFFFF";
    private static readonly Lazy<Dictionary<string, string>> MarkupByTemplateCode = new(LoadMarkup);

    /// <summary>
    /// The template code (<c>w:tmpl</c>) that identifies each built-in style, or null for <see cref="WordListStyle.Custom"/>,
    /// which gets a fresh code every time it is added.
    /// </summary>
    private static string? GetTemplateCode(WordListStyle style) => style switch {
        WordListStyle.Bulleted => "934E79A6",
        WordListStyle.ArticleSections => "04090023",
        WordListStyle.Headings111 => "04090025",
        WordListStyle.HeadingIA1 => "04090027",
        WordListStyle.Chapters => "04090029",
        WordListStyle.BulletedChars => "04090021",
        WordListStyle.Heading1ai => "0409001D",
        WordListStyle.Headings111Shifted => "0409001F",
        WordListStyle.LowerLetterWithBracket => "BB9E481E",
        WordListStyle.LowerLetterWithDot => "73ECA528",
        WordListStyle.UpperLetterWithDot => "76643E8A",
        WordListStyle.UpperLetterWithBracket => "76643E8C",
        WordListStyle.Numbered => "25BE4216",
        WordListStyle.Custom => null,
        _ => throw new ArgumentOutOfRangeException(nameof(style))
    };

    /// <summary>
    /// Matches a list definition to one of the built-in styles by its template code, or returns
    /// <see cref="WordListStyle.Custom"/> when none matches. Template codes are hexadecimal, so case doesn't matter.
    /// </summary>
    internal static WordListStyle MatchStyle(string? templateCode) => templateCode?.ToUpperInvariant() switch {
        "934E79A6" => WordListStyle.Bulleted,
        "04090023" => WordListStyle.ArticleSections,
        "04090025" => WordListStyle.Headings111,
        "04090027" => WordListStyle.HeadingIA1,
        "04090029" => WordListStyle.Chapters,
        "04090021" => WordListStyle.BulletedChars,
        "0409001D" => WordListStyle.Heading1ai,
        "0409001F" => WordListStyle.Headings111Shifted,
        "BB9E481E" => WordListStyle.LowerLetterWithBracket,
        "73ECA528" => WordListStyle.LowerLetterWithDot,
        "76643E8A" => WordListStyle.UpperLetterWithDot,
        "76643E8C" => WordListStyle.UpperLetterWithBracket,
        "25BE4216" => WordListStyle.Numbered,
        _ => WordListStyle.Custom
    };

    /// <summary>
    /// The <c>w:abstractNum</c> template for a built-in style, declaring the WordprocessingML and Word 2012 namespaces it
    /// uses. Its <c>w:abstractNumId</c> and <c>w:nsid</c> are placeholders, as is the <c>w:tmpl</c> of
    /// <see cref="WordListStyle.Custom"/>; the caller sets them for each list it adds. The same string instance is
    /// returned on every call.
    /// </summary>
    internal static string GetDefinitionMarkup(WordListStyle style) => MarkupByTemplateCode.Value[GetTemplateCode(style) ?? CustomTemplateCode];

    /// <summary>
    /// The template of a single-level bulleted definition that uses a picture bullet, with placeholders as in
    /// <see cref="GetDefinitionMarkup"/> and a <c>w:lvlPicBulletId</c> of 0.
    /// </summary>
    internal static string PictureBulletMarkup => MarkupByTemplateCode.Value[PictureBulletTemplateCode];

    /// <summary>Whether a template code is a placeholder that each added definition replaces with a fresh code.</summary>
    internal static bool IsPlaceholderTemplateCode(string? templateCode) =>
        templateCode == CustomTemplateCode || templateCode == PictureBulletTemplateCode;

    /// <summary>A new eight-digit hexadecimal identifier, as Word uses for <c>w:nsid</c> and <c>w:tmpl</c>.</summary>
    internal static string GenerateNsidValue() => Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();

    private static Dictionary<string, string> LoadMarkup() {
        byte[] bytes;
        using (Stream stream = typeof(WordListStyles).Assembly.GetManifestResourceStream(MarkupResource)
            ?? throw new InvalidOperationException("The built-in list definitions are missing from the assembly.")) {
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            bytes = buffer.ToArray();
        }
        OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
        const int W = OfficeXmlNamespaces.WordprocessingML;
        var markup = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int node = tree.FirstChildOf(tree.Root); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
            if (!tree.IsElement(node, W, "abstractNum")) continue;
            int templateCode = tree.FindChild(node, W, "tmpl");
            markup[tree.GetAttribute(templateCode, W, "val")!] = Encoding.UTF8.GetString(tree.WriteElement(node));
        }
        return markup;
    }
}
