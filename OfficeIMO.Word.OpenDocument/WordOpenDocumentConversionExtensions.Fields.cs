using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static bool TryMapWordField(WordInlineFieldSnapshot field, out OdtFieldKind kind) {
        kind = default;
        if (field.HasUnsupportedResultContent || field.HasUnsupportedContainer) return false;
        string[] tokens = field.Instruction.Split(new[] { ' ', '\t', '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 1 && !(tokens.Length == 3 && tokens[1] == "\\*" &&
            string.Equals(tokens[2], "MERGEFORMAT", StringComparison.OrdinalIgnoreCase))) return false;
        if (tokens.Length == 0) return false;
        switch (tokens[0].ToUpperInvariant()) {
            case "PAGE": kind = OdtFieldKind.PageNumber; return true;
            case "NUMPAGES":
                if (field.IsLocked) return false;
                kind = OdtFieldKind.PageCount; return true;
            case "DATE": kind = OdtFieldKind.Date; return true;
            case "TIME": kind = OdtFieldKind.Time; return true;
            default: return false;
        }
    }

    private static bool TryMapOdtField(OdtField field, out WordFieldType type) {
        type = default;
        if (!field.IsBasic) return false;
        type = field.Kind switch {
            OdtFieldKind.PageNumber => WordFieldType.Page,
            OdtFieldKind.PageCount => WordFieldType.NumPages,
            OdtFieldKind.Date => WordFieldType.Date,
            OdtFieldKind.Time => WordFieldType.Time,
            _ => default
        };
        return Enum.IsDefined(typeof(OdtFieldKind), field.Kind);
    }

    private static int CountOdtFields(OdtDocument document) =>
        new[] { "content.xml", "styles.xml" }.Sum(part => document.Package.GetXml(part)
            .Descendants().Count(element => OdtField.TryGetKind(element.Name, out _)));
}
