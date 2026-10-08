using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void MaterializeDocumentDefaultCharacterScale(Dictionary<string, Style> paragraphStyles, Styles? styles) {
            CharacterScale? defaults = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<CharacterScale>();
            if (defaults == null) return;
            _ = ReadSupportedCharacterScale(defaults);

            // DOC has no docDefaults record. Roots retain the document default,
            // while descendants inherit their base style's authored overrides.
            foreach (string styleId in paragraphStyles.Keys.ToArray()) {
                Style original = paragraphStyles[styleId];
                string? baseId = original.BasedOn?.Val?.Value;
                if (!string.IsNullOrWhiteSpace(baseId) && paragraphStyles.ContainsKey(baseId!)) continue;
                if (original.StyleRunProperties?.GetFirstChild<CharacterScale>() != null) continue;
                Style style = (Style)original.CloneNode(true);
                style.StyleRunProperties ??= new StyleRunProperties();
                style.StyleRunProperties.AddChild(defaults.CloneNode(true), true);
                paragraphStyles[styleId] = style;
            }
        }

        private static int ReadSupportedCharacterScale(CharacterScale scale) {
            long? percentage = scale.Val?.Value;
            if (!percentage.HasValue || percentage.Value < 1 || percentage.Value > 600) {
                throw new NotSupportedException("Native DOC saving supports character scale percentages from 1 through 600.");
            }

            return (int)percentage.Value;
        }
    }
}
