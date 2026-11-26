using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void BuiltInTableStyleDefinitionsAreTypedStylesUnderTheirOwnIdentifiers() {
        foreach (WordTableStyle style in Enum.GetValues(typeof(WordTableStyle))) {
            Style definition = WordTableStyles.GetStyleDefinition(style);

            // Every element reads back as the SDK's typed class: a misspelled or misplaced element would be unknown.
            Assert.Empty(definition.Descendants<OpenXmlUnknownElement>());
            Assert.Equal(StyleValues.Table, definition.Type!.Value);
            Assert.Equal(WordTableStyles.GetStyleId(style), definition.StyleId!.Value);
            Assert.Equal(style, WordTableStyles.GetStyle(definition.StyleId.Value!));

            // The name Word shows is the enumeration name in words, so a definition filed under another style shows up.
            string expectedName = style switch {
                WordTableStyle.TableNormal => "Normal Table",
                WordTableStyle.OfficeIMOModern => "OfficeIMO Modern",
                _ => string.Join(" ", System.Text.RegularExpressions.Regex.Matches(style.ToString(), "[A-Z][a-z]*|[0-9]+").Cast<System.Text.RegularExpressions.Match>().Select(match => match.Value))
            };
            Assert.Equal(expectedName, definition.StyleName!.Val!.Value);

            // Callers add the definition to a document, so each call gets its own copy.
            Assert.NotSame(definition, WordTableStyles.GetStyleDefinition(style));
        }
    }
}
