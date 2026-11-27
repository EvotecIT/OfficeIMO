using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void MaterializeDocumentDefaultCharacterScale(Dictionary<string, LegacyDocElement> paragraphStyles, LegacyDocElement documentDefaults) {
            LegacyDocElement defaults = documentDefaults.Child("rPrDefault").Child("rPr").Child("w");
            if (defaults.IsNone) return;
            _ = ReadSupportedCharacterScale(defaults);

            // DOC has no docDefaults record. Roots retain the document default,
            // while descendants inherit their base style's authored overrides.
            foreach (string styleId in paragraphStyles.Keys.ToArray()) {
                LegacyDocElement original = paragraphStyles[styleId];
                string? baseId = original.Child("basedOn").Val;
                if (!string.IsNullOrWhiteSpace(baseId) && paragraphStyles.ContainsKey(baseId!)) continue;
                if (original.Child("rPr").Child("w").Exists) continue;
                OfficeXmlTree style = CopyStyleDefinition(original);
                int runProperties = GetOrInsertStyleChild(style, style.Root, "rPr", StyleChildOrder);
                style.InsertOrdered(runProperties, style.ImportElement(defaults.Tree, defaults.Node), WRPrView.Order);
                paragraphStyles[styleId] = new LegacyDocElement(style, style.Root);
            }
        }

        private static int ReadSupportedCharacterScale(LegacyDocElement scale) {
            long? percentage = LegacyDocValues.Int64(scale.Val);
            if (!percentage.HasValue || percentage.Value < 1 || percentage.Value > 600) {
                throw new NotSupportedException("Native DOC saving supports character scale percentages from 1 through 600.");
            }

            return (int)percentage.Value;
        }
    }
}
