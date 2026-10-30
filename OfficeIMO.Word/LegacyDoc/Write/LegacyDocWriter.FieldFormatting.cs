using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static readonly ConditionalWeakTable<Styles, FieldComparisonStyleCache> FieldComparisonStyles = new();

        private static LegacyDocWritableFormatting ReadFieldComparisonFormatting(
            Run run, LegacyDocWritableFormatting direct, LegacyDocWritableFormatting inherited) {
            // A field is flattened to one native display format. Compare the
            // inherited values before doing so, or an explicit off override can
            // erase visible content when another display run inherits hidden text.
            OpenXmlPartRootElement? root = run.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
            MainDocumentPart? main = (root?.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart;
            Styles? styles = main?.StyleDefinitionsPart?.Styles;
            if (main == null || styles == null) return direct.WithInheritedFormatting(inherited);

            Paragraph? paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
            string? styleId = paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) {
                styleId = run.Ancestors<Footnote>().Any() ? "FootnoteText"
                    : run.Ancestors<Endnote>().Any() ? "EndnoteText" : "Normal";
            }
            FieldComparisonStyleCache cache = FieldComparisonStyles.GetValue(styles, _ => new FieldComparisonStyleCache(main, styles));
            (LegacyDocWritableFormatting paragraphStyle, bool rootWithoutNormal) = cache.Resolve(styleId!);
            LegacyDocWritableFormatting normal = rootWithoutNormal ? cache.RootDefaults : cache.NormalDefaults;
            return direct.WithInheritedFormatting(paragraphStyle.WithInheritedFormatting(inherited).WithInheritedFormatting(normal));
        }

        private sealed class FieldComparisonStyleCache {
            private const int MaxFieldStyleDepth = 256;
            private readonly Dictionary<string, Style> _definitions = new(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, (LegacyDocWritableFormatting Formatting, bool RootWithoutNormal)> _resolved = new(StringComparer.OrdinalIgnoreCase);

            internal LegacyDocWritableFormatting NormalDefaults { get; }
            internal LegacyDocWritableFormatting RootDefaults { get; }

            internal FieldComparisonStyleCache(MainDocumentPart main, Styles styles) {
                foreach (Style style in styles.Elements<Style>()) {
                    string? id = style.StyleId?.Value;
                    if (style.Type?.Value == StyleValues.Paragraph && !string.IsNullOrWhiteSpace(id) && !_definitions.ContainsKey(id!))
                        _definitions.Add(id!, style);
                }
                StyleRunProperties defaults = CreateDefaultParagraphStyle(main, _definitions, styles).StyleRunProperties!;
                CharacterScale? documentScale = styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<CharacterScale>();
                if (defaults.GetFirstChild<CharacterScale>() == null)
                    defaults.AddChild(documentScale?.CloneNode(true) ?? new CharacterScale { Val = 100L }, true);
                NormalDefaults = ReadSupportedRunFormatting(defaults);
                var rootDefaults = (StyleRunProperties)defaults.CloneNode(true);
                rootDefaults.RemoveAllChildren<CharacterScale>();
                rootDefaults.AddChild(documentScale?.CloneNode(true) ?? new CharacterScale { Val = 100L }, true);
                RootDefaults = ReadSupportedRunFormatting(rootDefaults);
            }

            internal (LegacyDocWritableFormatting Formatting, bool RootWithoutNormal) Resolve(string styleId) {
                if (_resolved.TryGetValue(styleId, out var cached)) return cached;
                var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                LegacyDocWritableFormatting paragraphStyle = LegacyDocWritableFormatting.Plain;
                string? current = styleId;
                while (!string.IsNullOrWhiteSpace(current) && !string.Equals(current, "Normal", StringComparison.OrdinalIgnoreCase)
                    && visited.Add(current!)) {
                    if (visited.Count > MaxFieldStyleDepth) throw new NotSupportedException("Native DOC saving does not support field paragraph style chains beyond 256 levels.");
                    if (!_definitions.TryGetValue(current!, out Style? style)) break;
                    paragraphStyle = paragraphStyle.WithInheritedFormatting(ReadSupportedRunFormatting(GetSupportedTableStyleRunProperties(style)));
                    current = style.BasedOn?.Val?.Value;
                }
                var result = (paragraphStyle, !string.Equals(current, "Normal", StringComparison.OrdinalIgnoreCase) && visited.Count != 0);
                _resolved.Add(styleId, result);
                return result;
            }
        }
    }
}
